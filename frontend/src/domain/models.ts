export type Language = "English" | "French" | "German";

export type TranslationConfig = {
    source_column: number;
    column_1_language: Language;
    column_2_language: Language;
    column_3_language: Language;
};

export type RevisionInput = {
    type: "Added" | "Deleted" | "Formatted" | "None";
    text: string;
};

export type CellParagraphInput = {
    index: number;
    text: string;
};

export type ChangedSourceParagraphInput = {
    index: number;
    original_text: string;
    current_text: string;
    changes: RevisionInput[];
};

export type TargetCellInput = {
    column: number;
    expected_language: Language;
    paragraphs: CellParagraphInput[];
};

export type TranslateCellChangesRequest = {
    source_column: number;
    source_language: Language;
    source_cell: CellParagraphInput[];
    changed_source_paragraphs: ChangedSourceParagraphInput[];
    targets: TargetCellInput[];
};

export type CellEditOutput = {
    source_paragraph_indices: number[];
    operation: "replace" | "insert" | "none";
    target_paragraph_index: number;
    original_text: string;
    translated_text: string;
};

export type CellTranslationOutput = {
    column: number;
    language: Language;
    edits: CellEditOutput[];
};

export type TranslateCellChangesResponse = {
    translations: CellTranslationOutput[];
    numeric_consistent: boolean;
    numeric_warnings: string[];
    partial_errors: string[];
    failed_source_paragraph_indices: number[];
};

export type RevisionCounts = Record<string, number>;

export type RevisionBaseline = {
    version: 1;
    fingerprints: RevisionCounts;
};

export type SingleSpanDiff = {
    prefixLength: number;
    suffixLength: number;
    removedText: string;
    insertedText: string;
};
