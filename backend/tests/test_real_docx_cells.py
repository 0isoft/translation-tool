from pathlib import Path
import types
import unittest
from dataclasses import replace
from unittest.mock import patch
from zipfile import ZipFile

from app.domain.models import (
    CellParagraphInput,
    ChangedSourceParagraphInput,
    ClaudeCellEdit,
    ClaudeTargetCellPlan,
    ClaudeNewParagraphPlan,
    ClaudeNewParagraphTarget,
    RevisionInput,
    TargetCellInput,
    TranslateCellChangesRequest,
)
from app.domain.translation_rules import (
    PlanValidationError,
    deterministic_numeric_edit,
    is_effectively_new_paragraph,
    validate_target_plan,
)
from app.main import claude_planner, translate_cell_changes
from tests.docx_fixture import DocxParagraph, TrackedTableFixture


FIXTURE_PATH = Path(__file__).with_name("Tab_file_for_testing.docx")
ISOLATED_FIXTURE_PATH = Path(__file__).with_name("isolated_issue.docx")


def changed_input(paragraph: DocxParagraph) -> ChangedSourceParagraphInput:
    assert paragraph.current_index is not None
    changes = []
    if paragraph.deleted_text:
        changes.append(RevisionInput(type="Deleted", text=paragraph.deleted_text))
    if paragraph.added_text:
        changes.append(RevisionInput(type="Added", text=paragraph.added_text))
    return ChangedSourceParagraphInput(
        index=paragraph.current_index,
        original_text=paragraph.original_text,
        current_text=paragraph.current_text,
        changes=changes,
    )


def cell_input(
    fixture: TrackedTableFixture,
    row: int,
    column: int,
    language: str,
) -> TargetCellInput:
    return TargetCellInput(
        column=column,
        expected_language=language,
        paragraphs=[
            CellParagraphInput(index=paragraph.current_index, text=paragraph.current_text)
            for paragraph in fixture.paragraphs(row, column)
            if paragraph.current_index is not None
        ],
    )


class FakeNewParagraphMessages:
    async def parse(self, **_kwargs):
        return types.SimpleNamespace(
            parsed_output=ClaudeNewParagraphPlan(
                target_1=ClaudeNewParagraphTarget(
                    language="German",
                    insertion_index=999,
                    translated_text=(
                        "Außer Arbeitslosen, die nur in öffentlichen "
                        "Gesundheitsdiensten Anspruch auf Versorgung haben."
                    ),
                ),
                target_2=ClaudeNewParagraphTarget(
                    language="French",
                    insertion_index=999,
                    translated_text=(
                        "Sauf pour les chômeurs qui ont droit aux soins "
                        "uniquement dans les services de santé publics."
                    ),
                ),
            )
        )


class FakeNewParagraphClient:
    def __init__(self):
        self.messages = FakeNewParagraphMessages()


class PartiallyFailingNewParagraphMessages:
    def __init__(self):
        self.calls = 0

    async def parse(self, **_kwargs):
        self.calls += 1
        if 3 <= self.calls <= 5:
            raise RuntimeError("simulated failure for one new legal decision")

        return types.SimpleNamespace(
            parsed_output=ClaudeNewParagraphPlan(
                target_1=ClaudeNewParagraphTarget(
                    language="German",
                    insertion_index=999,
                    translated_text=f"Deutsche Entscheidung {self.calls}",
                ),
                target_2=ClaudeNewParagraphTarget(
                    language="French",
                    insertion_index=999,
                    translated_text=f"Décision française {self.calls}",
                ),
            )
        )


class PartiallyFailingNewParagraphClient:
    def __init__(self):
        self.messages = PartiallyFailingNewParagraphMessages()


class DuplicateAwareMessages:
    async def parse(self, **_kwargs):
        return types.SimpleNamespace(
            parsed_output=ClaudeNewParagraphPlan(
                target_1=ClaudeNewParagraphTarget(
                    language="German",
                    operation="none",
                    insertion_index=19,
                    translated_text="",
                ),
                target_2=ClaudeNewParagraphTarget(
                    language="French",
                    operation="insert",
                    insertion_index=20,
                    translated_text="Nouvelle décision.",
                ),
            )
        )


class DuplicateAwareClient:
    def __init__(self):
        self.messages = DuplicateAwareMessages()


class RealDocxCellTests(unittest.IsolatedAsyncioTestCase):
    @classmethod
    def setUpClass(cls):
        cls.fixture = TrackedTableFixture(FIXTURE_PATH)

    def test_fixture_has_real_seven_column_table(self):
        self.assertEqual(len(self.fixture.rows), 439)
        self.assertTrue(all(
            self.fixture.cell_count(row) == 7
            for row in range(len(self.fixture.rows))
        ))

    def test_isolated_fixture_has_shifted_columns_and_saved_baseline(self):
        isolated = TrackedTableFixture(ISOLATED_FIXTURE_PATH)
        self.assertEqual(len(isolated.rows), 2)
        self.assertTrue(all(
            isolated.cell_count(row) == 6
            for row in range(len(isolated.rows))
        ))
        self.assertEqual(
            isolated.paragraphs(0, 2)[0].current_text,
            "E-II-01-EL",
        )
        self.assertEqual(
            isolated.paragraphs(1, 2)[0].current_text,
            "E-II-02-EL",
        )
        with ZipFile(ISOLATED_FIXTURE_PATH) as archive:
            web_extension = archive.read(
                "word/webextensions/webextension1.xml"
            )
        self.assertIn(b"translationTool.revisionBaseline.v1", web_extension)

    def test_real_whitespace_only_revision_is_semantically_empty(self):
        paragraph = self.fixture.paragraphs(1, 3)[6]

        self.assertEqual(paragraph.deleted_text, " ")
        self.assertEqual(paragraph.added_text, "")
        self.assertEqual(
            "".join(
                character
                for character in paragraph.deleted_text
                if not character.isspace()
            ),
            "",
        )

    def test_real_single_amount_change_updates_both_language_cells(self):
        source = self.fixture.paragraphs(78, 3)[0]
        change = changed_input(source)

        german = deterministic_numeric_edit(
            change,
            cell_input(self.fixture, 78, 1, "German"),
            set(),
        )
        french = deterministic_numeric_edit(
            change,
            cell_input(self.fixture, 78, 2, "French"),
            set(),
        )

        self.assertIsNotNone(german)
        self.assertIsNotNone(french)
        assert german is not None and french is not None
        self.assertIn("€71,65", german.translated_text)
        self.assertNotIn("€69,90", german.translated_text)
        self.assertIn("€71,65", french.translated_text)
        self.assertNotIn("€69,90", french.translated_text)

    def test_real_repeated_amount_table_uses_ordinal_context(self):
        source_paragraphs = self.fixture.paragraphs(15, 3)
        changed = [
            paragraph
            for paragraph in source_paragraphs
            if paragraph.added_text or paragraph.deleted_text
        ]
        self.assertEqual(len(changed), 13)

        for column, language in ((1, "German"), (2, "French")):
            target = cell_input(self.fixture, 15, column, language)
            occupied: set[int] = set()
            edits = [
                deterministic_numeric_edit(
                    changed_input(paragraph),
                    target,
                    occupied,
                )
                for paragraph in changed
            ]

            self.assertTrue(all(edit is not None for edit in edits))
            self.assertEqual(len(occupied), 13)
            translated = [edit.translated_text for edit in edits if edit]
            self.assertTrue(any("€5,97" in text for text in translated))
            self.assertTrue(any("€7,16" in text for text in translated))
            self.assertTrue(any("€3,58" in text for text in translated))

    def test_real_complex_decision_change_falls_back_without_number_shift(self):
        source = self.fixture.paragraphs(157, 3)[3]
        change = changed_input(source)

        for column, language in ((1, "German"), (2, "French")):
            edit = deterministic_numeric_edit(
                change,
                cell_input(self.fixture, 157, column, language),
                set(),
            )
            self.assertIsNone(edit)

    async def test_real_new_paragraph_is_inserted_in_both_target_cells(self):
        source_paragraphs = self.fixture.paragraphs(54, 3)
        new_paragraph = source_paragraphs[6]
        self.assertEqual(new_paragraph.original_text, "")
        self.assertIn("Except for the unemployed", new_paragraph.current_text)

        source_cell = [
            CellParagraphInput(index=paragraph.current_index, text=paragraph.current_text)
            for paragraph in source_paragraphs
            if paragraph.current_index is not None
        ]
        targets = [
            cell_input(self.fixture, 54, 1, "German"),
            cell_input(self.fixture, 54, 2, "French"),
        ]
        request = TranslateCellChangesRequest(
            source_column=3,
            source_cell=source_cell,
            changed_source_paragraphs=[changed_input(new_paragraph)],
            targets=targets,
        )

        with patch(
            "app.adapters.anthropic.get_anthropic_client",
            return_value=FakeNewParagraphClient(),
        ), patch("app.adapters.anthropic.get_anthropic_model", return_value="fixture-model"):
            response = await translate_cell_changes(
                request,
                x_request_id="real-docx-new-paragraph",
            )

        self.assertEqual(response.partial_errors, [])
        self.assertEqual(response.failed_source_paragraph_indices, [])
        self.assertEqual(len(response.translations), 2)
        for translation, target in zip(
            response.translations,
            targets,
            strict=True,
        ):
            self.assertEqual(len(translation.edits), 1)
            edit = translation.edits[0]
            self.assertEqual(edit.operation, "insert")
            self.assertEqual(edit.target_paragraph_index, len(target.paragraphs))

    async def test_real_e_ii_02_addition_survives_misreported_original_text(self):
        source_paragraphs = self.fixture.paragraphs(53, 3)
        added_paragraph = source_paragraphs[5]
        assert added_paragraph.current_index is not None
        misleading_change = ChangedSourceParagraphInput(
            index=added_paragraph.current_index,
            # This simulates Office.js returning current text for both
            # reviewed versions of a wholly inserted paragraph.
            original_text=added_paragraph.current_text,
            current_text=added_paragraph.current_text,
            changes=[RevisionInput(
                type="Added",
                text=added_paragraph.current_text,
            )],
        )
        self.assertTrue(is_effectively_new_paragraph(misleading_change))

        request = TranslateCellChangesRequest(
            source_column=3,
            source_cell=[
                CellParagraphInput(
                    index=paragraph.current_index,
                    text=paragraph.current_text,
                )
                for paragraph in source_paragraphs
                if paragraph.current_index is not None
            ],
            changed_source_paragraphs=[misleading_change],
            targets=[
                cell_input(self.fixture, 53, 1, "German"),
                cell_input(self.fixture, 53, 2, "French"),
            ],
        )

        with patch(
            "app.adapters.anthropic.get_anthropic_client",
            return_value=FakeNewParagraphClient(),
        ), patch("app.adapters.anthropic.get_anthropic_model", return_value="fixture-model"):
            response = await translate_cell_changes(
                request,
                x_request_id="real-docx-e-ii-02",
            )

        self.assertEqual(response.failed_source_paragraph_indices, [])
        self.assertTrue(all(
            translation.edits[0].operation == "insert"
            for translation in response.translations
        ))

    def test_general_plan_cannot_silently_omit_real_legal_decisions(self):
        source_paragraphs = self.fixture.paragraphs(52, 3)
        new_decisions = [
            paragraph
            for paragraph in source_paragraphs
            if not paragraph.original_text.strip()
            and paragraph.current_text.startswith("Decision No.")
        ]
        source_changes = {
            paragraph.current_index: changed_input(paragraph)
            for paragraph in new_decisions
        }
        target = cell_input(self.fixture, 52, 2, "French")
        incomplete_plan = ClaudeTargetCellPlan(
            language="French",
            edits=[ClaudeCellEdit(
                source_paragraph_indices=[new_decisions[0].current_index],
                operation="insert",
                target_paragraph_index=len(target.paragraphs),
                translated_text="Décision n° 238/31-07-25 (Β' 4737)",
            )],
        )

        with self.assertRaises(PlanValidationError) as raised:
            validate_target_plan(target, incomplete_plan, source_changes)

        self.assertIn("omitted source paragraphs", str(raised.exception))

    async def test_real_legal_decisions_report_only_failed_paragraph_for_retry(self):
        source_paragraphs = self.fixture.paragraphs(52, 3)
        new_decisions = [
            paragraph
            for paragraph in source_paragraphs
            if not paragraph.original_text.strip()
            and paragraph.current_text.startswith("Decision No.")
        ]
        self.assertEqual(len(new_decisions), 5)

        request = TranslateCellChangesRequest(
            source_column=3,
            source_cell=[
                CellParagraphInput(
                    index=paragraph.current_index,
                    text=paragraph.current_text,
                )
                for paragraph in source_paragraphs
                if paragraph.current_index is not None
            ],
            changed_source_paragraphs=[
                changed_input(paragraph) for paragraph in new_decisions
            ],
            targets=[
                cell_input(self.fixture, 52, 1, "German"),
                cell_input(self.fixture, 52, 2, "French"),
            ],
        )

        retry_settings = replace(
            claude_planner.settings,
            claude_max_attempts=3,
            claude_backoff_initial_seconds=0,
        )
        with patch(
            "app.adapters.anthropic.get_anthropic_client",
            return_value=PartiallyFailingNewParagraphClient(),
        ), patch(
            "app.adapters.anthropic.get_anthropic_model",
            return_value="fixture-model",
        ), patch.object(
            claude_planner,
            "settings",
            retry_settings,
        ):
            response = await translate_cell_changes(
                request,
                x_request_id="real-docx-legal-decisions",
            )

        failed_index = new_decisions[2].current_index
        self.assertEqual(
            response.failed_source_paragraph_indices,
            [failed_index],
        )
        self.assertEqual(len(response.partial_errors), 1)
        self.assertTrue(all(
            len(translation.edits) == 4
            for translation in response.translations
        ))

    async def test_retry_can_skip_existing_target_without_duplication(self):
        source = self.fixture.paragraphs(52, 3)
        first_new_decision = source[21]
        request = TranslateCellChangesRequest(
            source_column=3,
            source_cell=[
                CellParagraphInput(
                    index=paragraph.current_index,
                    text=paragraph.current_text,
                )
                for paragraph in source
                if paragraph.current_index is not None
            ],
            changed_source_paragraphs=[changed_input(first_new_decision)],
            targets=[
                cell_input(self.fixture, 52, 1, "German"),
                cell_input(self.fixture, 52, 2, "French"),
            ],
        )

        with patch(
            "app.adapters.anthropic.get_anthropic_client",
            return_value=DuplicateAwareClient(),
        ), patch("app.adapters.anthropic.get_anthropic_model", return_value="fixture-model"):
            response = await translate_cell_changes(
                request,
                x_request_id="duplicate-aware-retry",
            )

        self.assertEqual(
            response.translations[0].edits[0].operation,
            "none",
        )
        self.assertEqual(
            response.translations[1].edits[0].operation,
            "insert",
        )


if __name__ == "__main__":
    unittest.main()
