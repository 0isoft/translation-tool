import unittest

from app.main import (
    CellParagraphInput,
    ChangedSourceParagraphInput,
    RevisionInput,
    TargetCellInput,
    deterministic_numeric_edit,
)


def numeric_change(
    original_text: str,
    current_text: str,
    index: int = 1,
) -> ChangedSourceParagraphInput:
    return ChangedSourceParagraphInput(
        index=index,
        original_text=original_text,
        current_text=current_text,
        changes=[
            RevisionInput(type="Deleted", text="old numbers"),
            RevisionInput(type="Added", text="new numbers"),
        ],
    )


def french_target(*paragraphs: str) -> TargetCellInput:
    return TargetCellInput(
        column=2,
        expected_language="French",
        paragraphs=[
            CellParagraphInput(index=index, text=text)
            for index, text in enumerate(paragraphs)
        ],
    )


class DeterministicNumericEditTests(unittest.TestCase):
    def test_number_word_does_not_shift_later_replacements(self):
        change = numeric_change(
            "Adjustment based on the sum divided by 2. For 2025, the "
            "coefficient is 2.4% (Decision 53457; Law 4670/2020).",
            "Adjustment based on the sum divided by 2. For 2026, the "
            "coefficient is 2.4% (Decision 31854; Law 4670/2020).",
        )
        target = french_target(
            "Aucun chiffre pertinent ici.",
            "Revalorisation calculée en divisant la somme par deux. Pour "
            "2025, le coefficient est de 2,4% (décision 53457; loi "
            "4670/2020).",
        )

        edit = deterministic_numeric_edit(change, target, set())

        self.assertIsNotNone(edit)
        assert edit is not None
        self.assertEqual(edit.target_paragraph_index, 1)
        self.assertEqual(
            edit.translated_text,
            "Revalorisation calculée en divisant la somme par deux. Pour "
            "2026, le coefficient est de 2,4% (décision 31854; loi "
            "4670/2020).",
        )

    def test_missing_old_value_never_replaces_the_next_number(self):
        change = numeric_change(
            "For 2025, Decision 53457 applies.",
            "For 2026, Decision 31854 applies.",
        )
        target = french_target(
            "Pour 2025, la décision 99999 s’applique ensuite en 2030."
        )

        edit = deterministic_numeric_edit(change, target, set())

        self.assertIsNone(edit)

    def test_changed_numeric_structure_falls_back_instead_of_shifting(self):
        change = numeric_change(
            "Decision 53457/12.12.2024.",
            "Decision 31854/05.12.2025 and 6519/08.12.2025.",
        )
        target = french_target(
            "Décision 53457 du 12 décembre 2024, article 25."
        )

        edit = deterministic_numeric_edit(change, target, set())

        self.assertIsNone(edit)

    def test_repeated_amount_uses_unchanged_ordinal_context(self):
        change = numeric_change(
            "2nd contribution €6.99",
            "2nd contribution €7.16",
            index=1,
        )
        target = french_target("1ère €6,99", "2e €6,99")

        edit = deterministic_numeric_edit(change, target, set())

        self.assertIsNotNone(edit)
        assert edit is not None
        self.assertEqual(edit.target_paragraph_index, 1)
        self.assertEqual(edit.translated_text, "2e €7,16")


if __name__ == "__main__":
    unittest.main()
