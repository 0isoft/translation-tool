from collections import Counter
import logging
import re

from app.domain.models import (
    CellEditOutput,
    CellParagraphInput,
    CellTranslationOutput,
    ChangedSourceParagraphInput,
    ClaudeTargetCellPlan,
    TargetCellInput,
    TranslateCellChangesRequest,
)


class PlanValidationError(ValueError):
    """Claude returned a structurally unsafe edit plan."""


logger = logging.getLogger("uvicorn.error")


def normalized_paragraph_text(text: str) -> str:
    return re.sub(r"[\s\x07]+", " ", text).strip()


def is_effectively_new_paragraph(
    paragraph: ChangedSourceParagraphInput,
) -> bool:
    current_text = normalized_paragraph_text(paragraph.current_text)
    if not current_text:
        return False

    if not normalized_paragraph_text(paragraph.original_text):
        return True

    added_text = normalized_paragraph_text("".join(
        change.text
        for change in paragraph.changes
        if change.type == "Added"
    ))
    has_substantive_deletion = any(
        change.type == "Deleted"
        and normalized_paragraph_text(change.text)
        for change in paragraph.changes
    )

    # Some Word versions expose the current text for both reviewed versions
    # when an entire paragraph was inserted. The tracked Added content is the
    # authoritative fallback in that case.
    return not has_substantive_deletion and added_text == current_text


NUMBER_PATTERN = re.compile(
    r"(?<!\w)[+-]?(?:\d{1,3}(?:[ .,\u00a0\u202f]\d{3})+|\d+)"
    r"(?:[.,]\d+)?%?"
)


def numeric_signature(text: str) -> Counter[str]:
    signature: Counter[str] = Counter()

    for match in NUMBER_PATTERN.findall(text):
        digits = "".join(character for character in match if character.isdigit())

        if digits:
            signature[digits] += 1

    return signature


def numeric_delta(before: str, after: str) -> tuple[Counter[str], Counter[str]]:
    before_signature = numeric_signature(before)
    after_signature = numeric_signature(after)
    return before_signature - after_signature, after_signature - before_signature


def is_numeric_only_change(before: str, after: str) -> bool:
    removed, added = numeric_delta(before, after)
    return bool(removed or added) and (
        NUMBER_PATTERN.sub("<NUMBER>", before)
        == NUMBER_PATTERN.sub("<NUMBER>", after)
    )


def numeric_matches(text: str) -> list[re.Match[str]]:
    return list(NUMBER_PATTERN.finditer(text))


def replace_digits_preserving_format(
    target_number: str,
    source_number: str,
) -> str:
    source_digits = "".join(
        character for character in source_number if character.isdigit()
    )
    target_digit_count = sum(
        character.isdigit() for character in target_number
    )

    if len(source_digits) == target_digit_count:
        digit_iterator = iter(source_digits)
        return "".join(
            next(digit_iterator) if character.isdigit() else character
            for character in target_number
        )

    target_decimal_separator = None
    for separator in (",", "."):
        if separator in target_number:
            target_decimal_separator = separator

    normalized_source = source_number.replace(" ", "")
    if target_decimal_separator == ",":
        normalized_source = normalized_source.replace(".", ",")
    elif target_decimal_separator == ".":
        normalized_source = normalized_source.replace(",", ".")

    return normalized_source


def select_numeric_paragraph(
    target: TargetCellInput,
    required_signature: Counter[str],
    context_signature: Counter[str],
    source_index: int,
    occupied_indices: set[int],
) -> CellParagraphInput | None:
    candidates: list[tuple[tuple[int, int], CellParagraphInput]] = []

    for paragraph in target.paragraphs:
        if paragraph.index in occupied_indices:
            continue

        signature = numeric_signature(paragraph.text)
        if any(
            signature[value] < count
            for value, count in required_signature.items()
        ):
            continue

        context_score = sum(
            min(signature[value], count)
            for value, count in context_signature.items()
        )
        candidates.append((
            (context_score, -abs(paragraph.index - source_index)),
            paragraph,
        ))

    if not candidates:
        return None

    candidates.sort(key=lambda candidate: candidate[0], reverse=True)
    if len(candidates) > 1 and candidates[0][0] == candidates[1][0]:
        return None

    return candidates[0][1]


def deterministic_numeric_edit(
    change: ChangedSourceParagraphInput,
    target: TargetCellInput,
    occupied_indices: set[int],
) -> CellEditOutput | None:
    if not is_numeric_only_change(change.original_text, change.current_text):
        return None

    before_matches = numeric_matches(change.original_text)
    after_matches = numeric_matches(change.current_text)
    if len(before_matches) != len(after_matches):
        return None

    changed_pairs = [
        (before.group(), after.group())
        for before, after in zip(before_matches, after_matches, strict=True)
        if numeric_signature(before.group()) != numeric_signature(after.group())
    ]
    if not changed_pairs:
        return None

    old_signature: Counter[str] = Counter()
    new_signature: Counter[str] = Counter()
    for old_number, new_number in changed_pairs:
        old_signature += numeric_signature(old_number)
        new_signature += numeric_signature(new_number)

    before_signature = numeric_signature(change.original_text)
    after_signature = numeric_signature(change.current_text)
    context_signature = before_signature & after_signature

    paragraph = select_numeric_paragraph(
        target,
        old_signature,
        context_signature,
        change.index,
        occupied_indices,
    )
    if paragraph is None:
        already_updated = select_numeric_paragraph(
            target,
            new_signature,
            context_signature,
            change.index,
            occupied_indices,
        )
        if already_updated is None:
            return None

        occupied_indices.add(already_updated.index)
        return CellEditOutput(
            source_paragraph_indices=[change.index],
            operation="none",
            target_paragraph_index=already_updated.index,
            original_text=already_updated.text,
            translated_text=already_updated.text,
        )

    replacements: list[tuple[int, int, str]] = []
    available_matches = numeric_matches(paragraph.text)
    used_match_indices: set[int] = set()

    for old_number, new_number in changed_pairs:
        old_number_signature = numeric_signature(old_number)
        match_index = next((
            index
            for index, match in enumerate(available_matches)
            if index not in used_match_indices
            and numeric_signature(match.group()) == old_number_signature
        ), None)
        if match_index is None:
            return None

        used_match_indices.add(match_index)
        match = available_matches[match_index]
        replacements.append((
            match.start(),
            match.end(),
            replace_digits_preserving_format(match.group(), new_number),
        ))

    translated_text = paragraph.text
    for start, end, replacement in reversed(replacements):
        translated_text = (
            translated_text[:start] + replacement + translated_text[end:]
        )

    occupied_indices.add(paragraph.index)
    return CellEditOutput(
        source_paragraph_indices=[change.index],
        operation="replace",
        target_paragraph_index=paragraph.index,
        original_text=paragraph.text,
        translated_text=translated_text,
    )


def build_deterministic_numeric_edits(
    request: TranslateCellChangesRequest,
) -> tuple[dict[int, list[CellEditOutput]], set[int]]:
    edits_by_column = {target.column: [] for target in request.targets}
    occupied_by_column = {target.column: set() for target in request.targets}
    handled_source_indices: set[int] = set()

    for change in request.changed_source_paragraphs:
        trial_occupied = {
            column: indices.copy()
            for column, indices in occupied_by_column.items()
        }
        proposed_edits = [
            deterministic_numeric_edit(
                change,
                target,
                trial_occupied[target.column],
            )
            for target in request.targets
        ]

        if any(edit is None for edit in proposed_edits):
            continue

        occupied_by_column = trial_occupied
        handled_source_indices.add(change.index)
        for target, edit in zip(
            request.targets,
            proposed_edits,
            strict=True,
        ):
            edits_by_column[target.column].append(edit)

    return edits_by_column, handled_source_indices


def validate_target_plan(
    target: TargetCellInput,
    plan: ClaudeTargetCellPlan,
    source_changes: dict[int, ChangedSourceParagraphInput],
) -> CellTranslationOutput:
    if plan.language != target.expected_language:
        raise PlanValidationError(
            f"Claude returned {plan.language} content for column "
            f"{target.column}, which is configured as "
            f"{target.expected_language}."
        )

    expected_source_order = list(source_changes)
    expected_source_indices = set(source_changes)
    replaced_target_indices: set[int] = set()
    edits: list[CellEditOutput] = []
    deletion_only_plan = all(
        any(change.type == "Deleted" and change.text.strip()
            for change in source_change.changes)
        and not any(change.type == "Added" and change.text.strip()
                    for change in source_change.changes)
        for source_change in source_changes.values()
    )

    if not plan.edits:
        raise PlanValidationError(
            f"Claude returned no target edits for column {target.column}."
        )

    for edit_index, candidate in enumerate(plan.edits):
        normalized_source_indices = [
            index
            for index in candidate.source_paragraph_indices
            if index in expected_source_indices
        ]
        if not normalized_source_indices:
            normalized_source_indices = [
                expected_source_order[min(
                    edit_index,
                    len(expected_source_order) - 1,
                )]
            ]
            logger.warning(
                "rebound advisory source references for column %d edit %d",
                target.column,
                edit_index,
            )

        if candidate.operation == "replace" \
                and candidate.target_paragraph_index in replaced_target_indices:
            raise PlanValidationError(
                f"Claude returned duplicate target edits for column {target.column}."
            )
        if candidate.operation == "replace":
            replaced_target_indices.add(candidate.target_paragraph_index)

        normalized_target_index = candidate.target_paragraph_index
        normalized_operation = candidate.operation

        if candidate.operation in ("replace", "none"):
            if candidate.target_paragraph_index >= len(target.paragraphs):
                raise PlanValidationError(
                    f"Claude returned an invalid paragraph index for column {target.column}."
                )

            actual_text = target.paragraphs[candidate.target_paragraph_index].text
            translated_text = (
                actual_text
                if candidate.operation == "none"
                else candidate.translated_text
            )
            if deletion_only_plan \
                    and candidate.operation == "replace" \
                    and len(translated_text) > len(actual_text):
                logger.warning(
                    "converted expanding deletion edit to no-op for column %d",
                    target.column,
                )
                normalized_operation = "none"
                translated_text = actual_text
        else:
            normalized_target_index = min(
                candidate.target_paragraph_index,
                len(target.paragraphs),
            )
            if normalized_target_index != candidate.target_paragraph_index:
                logger.warning(
                    "normalized insertion index %d to %d for column %d",
                    candidate.target_paragraph_index,
                    normalized_target_index,
                    target.column,
                )
            actual_text = ""
            translated_text = candidate.translated_text

        edits.append(CellEditOutput(
            **candidate.model_dump(exclude={
                "source_paragraph_indices",
                "operation",
                "target_paragraph_index",
                "translated_text",
            }),
            source_paragraph_indices=normalized_source_indices,
            operation=normalized_operation,
            target_paragraph_index=normalized_target_index,
            original_text=actual_text,
            translated_text=translated_text,
        ))

    covered_source_indices = {
        source_index
        for edit in edits
        for source_index in edit.source_paragraph_indices
    }
    missing_source_indices = expected_source_indices - covered_source_indices
    if missing_source_indices:
        raise PlanValidationError(
            f"Claude omitted source paragraphs "
            f"{sorted(missing_source_indices)} for column {target.column}."
        )

    return CellTranslationOutput(
        column=target.column,
        language=plan.language,
        edits=edits,
    )


def get_numeric_warnings(
    translations: list[CellTranslationOutput],
    source_changes: dict[int, ChangedSourceParagraphInput],
) -> list[str]:
    warnings: list[str] = []
    source_before = "\n".join(
        change.original_text for change in source_changes.values()
    )
    source_after = "\n".join(
        change.current_text for change in source_changes.values()
    )
    source_removed, source_added = numeric_delta(source_before, source_after)
    all_changes_are_numeric_only = all(
        is_numeric_only_change(change.original_text, change.current_text)
        for change in source_changes.values()
    )

    for translation in translations:
        target_before = "\n".join(
            edit.original_text for edit in translation.edits
        )
        target_after = "\n".join(
            edit.translated_text for edit in translation.edits
        )
        target_removed, target_added = numeric_delta(
            target_before,
            target_after,
        )
        target_signature = numeric_signature(target_after)
        numeric_change_matches = (
            (source_removed, source_added) == (target_removed, target_added)
            or (
                bool(source_added)
                and all(
                    target_signature[value] >= count
                    for value, count in source_added.items()
                )
            )
        )

        if not numeric_change_matches:
            warnings.append(
                f"Column {translation.column}: the aggregate numeric changes "
                "do not match the reference-language revisions."
            )
            continue

        if all_changes_are_numeric_only:
            for edit in translation.edits:
                if edit.operation == "insert":
                    continue
                if (
                NUMBER_PATTERN.sub("<NUMBER>", edit.original_text)
                != NUMBER_PATTERN.sub("<NUMBER>", edit.translated_text)
                ):
                    warnings.append(
                        f"Column {translation.column}, paragraph "
                        f"{edit.target_paragraph_index}: a number-only revision "
                        "changed nonnumeric target wording."
                    )

    return warnings
