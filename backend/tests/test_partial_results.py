import json
import types
import unittest
from dataclasses import replace
from unittest.mock import patch

from app.domain.models import (
    CellParagraphInput,
    ChangedSourceParagraphInput,
    ClaudeNewParagraphPlan,
    ClaudeNewParagraphTarget,
    RevisionInput,
    TargetCellInput,
    TranslateCellChangesRequest,
)
from app.main import claude_planner, translate_cell_changes


class FakeMessages:
    def __init__(self):
        self.calls = 0

    async def parse(self, **_kwargs):
        self.calls += 1
        if self.calls == 1:
            return types.SimpleNamespace(
                parsed_output=ClaudeNewParagraphPlan(
                    target_1=ClaudeNewParagraphTarget(
                        language="German",
                        insertion_index=1,
                        translated_text="Neuer Absatz.",
                    ),
                    target_2=ClaudeNewParagraphTarget(
                        language="French",
                        insertion_index=1,
                        translated_text="Nouveau paragraphe.",
                    ),
                )
            )

        raise RuntimeError("simulated existing-paragraph planner failure")


class FakeClient:
    def __init__(self):
        self.messages = FakeMessages()


class GermanReferenceMessages:
    def __init__(self):
        self.source_language = None

    async def parse(self, **kwargs):
        payload = json.loads(kwargs["messages"][0]["content"])
        self.source_language = payload["source_language"]
        return types.SimpleNamespace(
            parsed_output=ClaudeNewParagraphPlan(
                target_1=ClaudeNewParagraphTarget(
                    language="English",
                    insertion_index=1,
                    translated_text="New paragraph.",
                ),
                target_2=ClaudeNewParagraphTarget(
                    language="French",
                    insertion_index=1,
                    translated_text="Nouveau paragraphe.",
                ),
            )
        )


class GermanReferenceClient:
    def __init__(self):
        self.messages = GermanReferenceMessages()


class PartialResultTests(unittest.IsolatedAsyncioTestCase):
    async def test_german_reference_translates_to_english_and_french(self):
        request = TranslateCellChangesRequest(
            source_column=1,
            source_language="German",
            source_cell=[
                CellParagraphInput(index=0, text="Bestehend."),
                CellParagraphInput(index=1, text="Neuer Absatz."),
            ],
            changed_source_paragraphs=[
                ChangedSourceParagraphInput(
                    index=1,
                    original_text="",
                    current_text="Neuer Absatz.",
                    changes=[
                        RevisionInput(type="Added", text="Neuer Absatz.")
                    ],
                )
            ],
            targets=[
                TargetCellInput(
                    column=2,
                    expected_language="English",
                    paragraphs=[
                        CellParagraphInput(index=0, text="Existing.")
                    ],
                ),
                TargetCellInput(
                    column=3,
                    expected_language="French",
                    paragraphs=[
                        CellParagraphInput(index=0, text="Existant.")
                    ],
                ),
            ],
        )
        fake_client = GermanReferenceClient()

        with patch(
            "app.adapters.anthropic.get_anthropic_client",
            return_value=fake_client,
        ):
            response = await translate_cell_changes(
                request,
                x_request_id="german-reference-test",
            )

        self.assertEqual(fake_client.messages.source_language, "German")
        self.assertEqual(
            [translation.language for translation in response.translations],
            ["English", "French"],
        )
        self.assertEqual(
            response.translations[0].edits[0].translated_text,
            "New paragraph.",
        )

    async def test_new_paragraph_edits_survive_later_planner_failure(self):
        request = TranslateCellChangesRequest(
            source_column=3,
            source_language="English",
            source_cell=[
                CellParagraphInput(index=0, text="Changed existing text."),
                CellParagraphInput(index=1, text="New paragraph."),
            ],
            changed_source_paragraphs=[
                ChangedSourceParagraphInput(
                    index=0,
                    original_text="Existing text.",
                    current_text="Changed existing text.",
                    changes=[RevisionInput(type="Added", text="Changed")],
                ),
                ChangedSourceParagraphInput(
                    index=1,
                    original_text="",
                    current_text="New paragraph.",
                    changes=[
                        RevisionInput(type="Added", text="New paragraph.")
                    ],
                ),
            ],
            targets=[
                TargetCellInput(
                    column=1,
                    expected_language="German",
                    paragraphs=[CellParagraphInput(index=0, text="Alt.")],
                ),
                TargetCellInput(
                    column=2,
                    expected_language="French",
                    paragraphs=[CellParagraphInput(index=0, text="Ancien.")],
                ),
            ],
        )
        fake_client = FakeClient()

        single_attempt_settings = replace(
            claude_planner.settings,
            claude_max_attempts=1,
        )
        with patch.object(
            claude_planner,
            "settings",
            single_attempt_settings,
        ), patch(
            "app.adapters.anthropic.get_anthropic_client",
            return_value=fake_client,
        ):
            response = await translate_cell_changes(
                request,
                x_request_id="partial-result-test",
            )

        self.assertEqual(len(response.partial_errors), 1)
        self.assertEqual(response.failed_source_paragraph_indices, [0])
        self.assertEqual(response.translations[0].edits[0].operation, "insert")
        self.assertEqual(
            response.translations[0].edits[0].translated_text,
            "Neuer Absatz.",
        )
        self.assertEqual(response.translations[1].edits[0].operation, "insert")
        self.assertEqual(
            response.translations[1].edits[0].translated_text,
            "Nouveau paragraphe.",
        )


if __name__ == "__main__":
    unittest.main()
