import json
import logging
import time

from anthropic import APITimeoutError

from app.domain.models import (
    CellEditOutput,
    CellTranslationOutput,
    ClaudeCellTranslationPlan,
    ClaudeNewParagraphPlan,
    TranslateCellChangesRequest,
    TranslateCellChangesResponse,
)
from app.domain.prompts import EXISTING_PARAGRAPH_PROMPT, NEW_PARAGRAPH_PROMPT
from app.domain.translation_rules import (
    PlanValidationError,
    build_deterministic_numeric_edits,
    get_numeric_warnings,
    is_effectively_new_paragraph,
    validate_target_plan,
)
from app.ports.claude import ClaudePlanningPort


logger = logging.getLogger("uvicorn.error")


class ApplicationRequestError(ValueError):
    def __init__(self, *, status_code: int, detail: str):
        super().__init__(detail)
        self.status_code = status_code
        self.detail = detail


async def propagate_cell_changes(
    request: TranslateCellChangesRequest,
    request_id: str,
    planner: ClaudePlanningPort,
) -> TranslateCellChangesResponse:
    started_at = time.monotonic()

    expected_target_columns = {1, 2, 3} - {request.source_column}
    if {target.column for target in request.targets} != expected_target_columns:
        raise ApplicationRequestError(
            status_code=422,
            detail="Targets must be the two non-source columns.",
        )

    if {target.expected_language for target in request.targets} \
            != {"French", "German"}:
        raise ApplicationRequestError(
            status_code=422,
            detail="Targets must contain French and German exactly once.",
        )

    source_indices = {paragraph.index for paragraph in request.source_cell}
    changed_indices = [
        paragraph.index for paragraph in request.changed_source_paragraphs
    ]
    if len(changed_indices) != len(set(changed_indices)) \
            or not set(changed_indices) <= source_indices:
        raise ApplicationRequestError(
            status_code=422,
            detail="Changed source paragraph indices must be unique and present in the source cell.",
        )

    deterministic_edits, deterministic_source_indices = (
        build_deterministic_numeric_edits(request)
    )
    new_source_paragraphs = [
        paragraph
        for paragraph in request.changed_source_paragraphs
        if is_effectively_new_paragraph(paragraph)
    ]
    new_source_indices = {
        paragraph.index for paragraph in new_source_paragraphs
    }
    remaining_changes = [
        paragraph
        for paragraph in request.changed_source_paragraphs
        if paragraph.index not in deterministic_source_indices
        and paragraph.index not in new_source_indices
    ]

    logger.info(
        "translation request %s: %d changed source paragraphs",
        request_id,
        len(request.changed_source_paragraphs),
    )

    all_source_changes = {
        paragraph.index: paragraph
        for paragraph in request.changed_source_paragraphs
    }
    translations = [
        CellTranslationOutput(
            column=target.column,
            language=target.expected_language,
            edits=[],
        )
        for target in request.targets
    ]
    partial_errors: list[str] = []
    failed_source_paragraph_indices: set[int] = set()

    for new_paragraph in new_source_paragraphs:
        new_paragraph_input = {
            "source_cell": [
                paragraph.model_dump() for paragraph in request.source_cell
            ],
            "new_source_paragraph": new_paragraph.model_dump(),
            "target_1": request.targets[0].model_dump(),
            "target_2": request.targets[1].model_dump(),
        }

        def validate_new_paragraph_response(candidate_response) -> None:
            candidate_plan = candidate_response.parsed_output
            if candidate_plan is None:
                raise ValueError("Claude returned no translation")

            candidate_targets = (
                candidate_plan.target_1,
                candidate_plan.target_2,
            )
            mismatches = [
                f"column {target.column} expected {target.expected_language}, "
                f"received {target_plan.language}"
                for target, target_plan in zip(
                    request.targets,
                    candidate_targets,
                    strict=True,
                )
                if target_plan.language != target.expected_language
            ]
            if mismatches:
                raise ValueError("; ".join(mismatches))

        try:
            new_response = await planner.parse(
                request_id=request_id,
                purpose=f"new paragraph {new_paragraph.index}",
                response_validator=validate_new_paragraph_response,
                model=planner.model,
                max_tokens=2048,
                system=NEW_PARAGRAPH_PROMPT,
                messages=[
                    {
                        "role": "user",
                        "content": json.dumps(
                            new_paragraph_input,
                            ensure_ascii=False,
                            indent=2,
                        ),
                    }
                ],
                output_format=ClaudeNewParagraphPlan,
            )
        except APITimeoutError as error:
            logger.warning("new paragraph request %s timed out", request_id)
            partial_errors.append(
                f"New paragraph {new_paragraph.index}: Claude timed out."
            )
            failed_source_paragraph_indices.add(new_paragraph.index)
            continue
        except Exception as error:
            logger.exception(
                "new paragraph request %s failed",
                request_id,
            )
            partial_errors.append(
                f"New paragraph {new_paragraph.index}: {error}"
            )
            failed_source_paragraph_indices.add(new_paragraph.index)
            continue

        new_plan = new_response.parsed_output
        if new_plan is None:
            partial_errors.append(
                f"New paragraph {new_paragraph.index}: Claude returned no translation."
            )
            failed_source_paragraph_indices.add(new_paragraph.index)
            continue

        target_plans = (new_plan.target_1, new_plan.target_2)
        language_mismatches = [
            f"column {target.column} expected {target.expected_language}, "
            f"received {target_plan.language}"
            for target, target_plan in zip(
                request.targets,
                target_plans,
                strict=True,
            )
            if target_plan.language != target.expected_language
        ]
        if language_mismatches:
            partial_errors.append(
                f"New paragraph {new_paragraph.index}: "
                + "; ".join(language_mismatches)
            )
            failed_source_paragraph_indices.add(new_paragraph.index)
            continue

        for translation, target, target_plan in zip(
            translations,
            request.targets,
            target_plans,
            strict=True,
        ):
            if target_plan.operation == "none":
                if target.paragraphs:
                    target_index = min(
                        target_plan.insertion_index,
                        len(target.paragraphs) - 1,
                    )
                    original_text = target.paragraphs[target_index].text
                else:
                    target_index = 0
                    original_text = ""
                translation.edits.append(CellEditOutput(
                    source_paragraph_indices=[new_paragraph.index],
                    operation="none",
                    target_paragraph_index=target_index,
                    original_text=original_text,
                    translated_text=original_text,
                ))
            else:
                translation.edits.append(CellEditOutput(
                    source_paragraph_indices=[new_paragraph.index],
                    operation="insert",
                    target_paragraph_index=min(
                        target_plan.insertion_index,
                        len(target.paragraphs),
                    ),
                    original_text="",
                    translated_text=target_plan.translated_text,
                ))

    if remaining_changes:
        remaining_source_changes = {
            paragraph.index: paragraph
            for paragraph in remaining_changes
        }
        claude_input = {
            "source_cell": [
                paragraph.model_dump() for paragraph in request.source_cell
            ],
            "changed_source_paragraphs": [
                paragraph.model_dump() for paragraph in remaining_changes
            ],
            "target_1": request.targets[0].model_dump(),
            "target_2": request.targets[1].model_dump(),
        }

        def validate_existing_response(candidate_response) -> None:
            candidate_plan = candidate_response.parsed_output
            if candidate_plan is None:
                raise ValueError("Claude returned no edit plan")

            for target, target_plan in zip(
                request.targets,
                (candidate_plan.target_1, candidate_plan.target_2),
                strict=True,
            ):
                validate_target_plan(
                    target,
                    target_plan,
                    remaining_source_changes,
                )

        response = None
        try:
            response = await planner.parse(
                request_id=request_id,
                purpose="existing paragraph translation",
                response_validator=validate_existing_response,
                model=planner.model,
                max_tokens=4096,
                system=EXISTING_PARAGRAPH_PROMPT,
                messages=[
                    {
                        "role": "user",
                        "content": json.dumps(
                            claude_input,
                            ensure_ascii=False,
                            indent=2,
                        ),
                    }
                ],
                output_format=ClaudeCellTranslationPlan,
            )
        except APITimeoutError as error:
            logger.warning("translation request %s timed out", request_id)
            partial_errors.append("Existing paragraph translation timed out.")
            failed_source_paragraph_indices.update(
                paragraph.index for paragraph in remaining_changes
            )
        except Exception as error:
            logger.exception("translation request %s failed", request_id)
            partial_errors.append(
                f"Existing paragraph translation failed: {error}"
            )
            failed_source_paragraph_indices.update(
                paragraph.index for paragraph in remaining_changes
            )

        if response is not None:
            plan = response.parsed_output
            if plan is None:
                partial_errors.append(
                    "Existing paragraph translation returned no edit plan."
                )
                failed_source_paragraph_indices.update(
                    paragraph.index for paragraph in remaining_changes
                )
            else:
                try:
                    planned_translations = [
                        validate_target_plan(
                            target,
                            target_plan,
                            remaining_source_changes,
                        )
                        for target, target_plan in zip(
                            request.targets,
                            (plan.target_1, plan.target_2),
                            strict=True,
                        )
                    ]
                except PlanValidationError as error:
                    partial_errors.append(
                        f"Existing paragraph edit plan rejected: {error}"
                    )
                    failed_source_paragraph_indices.update(
                        paragraph.index for paragraph in remaining_changes
                    )
                else:
                    for translation, planned_translation in zip(
                        translations,
                        planned_translations,
                        strict=True,
                    ):
                        translation.edits.extend(planned_translation.edits)

    for translation in translations:
        translation.edits.extend(deterministic_edits[translation.column])

    warnings = get_numeric_warnings(translations, all_source_changes)

    logger.info(
        "translation request %s completed in %.2fs with %d edits",
        request_id,
        time.monotonic() - started_at,
        sum(len(translation.edits) for translation in translations),
    )

    return TranslateCellChangesResponse(
        translations=translations,
        numeric_consistent=not warnings,
        numeric_warnings=warnings,
        partial_errors=partial_errors,
        failed_source_paragraph_indices=sorted(
            failed_source_paragraph_indices
        ),
    )
