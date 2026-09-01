from typing import Literal

from pydantic import BaseModel, Field


Language = Literal["English", "French", "German"]
TargetLanguage = Literal["French", "German"]


class TranslationConfig(BaseModel):
    source_column: int = Field(ge=1, le=3)
    column_1_language: Language
    column_2_language: Language
    column_3_language: Language


class SourceColumnConfig(BaseModel):
    source_column: int = Field(ge=1, le=3)


class RevisionInput(BaseModel):
    type: Literal["Added", "Deleted", "Formatted", "None"]
    text: str


class CellParagraphInput(BaseModel):
    index: int = Field(ge=0)
    text: str


class ChangedSourceParagraphInput(BaseModel):
    index: int = Field(ge=0)
    original_text: str
    current_text: str
    changes: list[RevisionInput] = Field(min_length=1)


class TargetCellInput(BaseModel):
    column: int = Field(ge=1, le=3)
    expected_language: TargetLanguage
    paragraphs: list[CellParagraphInput]


class TranslateCellChangesRequest(BaseModel):
    source_column: int = Field(ge=1, le=3)
    source_cell: list[CellParagraphInput]
    changed_source_paragraphs: list[ChangedSourceParagraphInput] = Field(
        min_length=1
    )
    targets: list[TargetCellInput] = Field(min_length=2, max_length=2)


class ClaudeCellEdit(BaseModel):
    source_paragraph_indices: list[int] = Field(min_length=1)
    operation: Literal["replace", "insert", "none"]
    target_paragraph_index: int = Field(ge=0)
    translated_text: str


class ClaudeTargetCellPlan(BaseModel):
    language: TargetLanguage
    edits: list[ClaudeCellEdit]


class ClaudeCellTranslationPlan(BaseModel):
    target_1: ClaudeTargetCellPlan
    target_2: ClaudeTargetCellPlan


class ClaudeNewParagraphTarget(BaseModel):
    language: TargetLanguage
    operation: Literal["insert", "none"] = "insert"
    insertion_index: int = Field(ge=0)
    translated_text: str


class ClaudeNewParagraphPlan(BaseModel):
    target_1: ClaudeNewParagraphTarget
    target_2: ClaudeNewParagraphTarget


class CellEditOutput(BaseModel):
    source_paragraph_indices: list[int]
    operation: Literal["replace", "insert", "none"]
    target_paragraph_index: int
    original_text: str
    translated_text: str


class CellTranslationOutput(BaseModel):
    column: int = Field(ge=1, le=3)
    language: TargetLanguage
    edits: list[CellEditOutput]


class TranslateCellChangesResponse(BaseModel):
    translations: list[CellTranslationOutput]
    numeric_consistent: bool
    numeric_warnings: list[str]
    partial_errors: list[str]
    failed_source_paragraph_indices: list[int]
