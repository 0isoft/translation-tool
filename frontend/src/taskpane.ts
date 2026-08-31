const output =
    document.getElementById("output") as HTMLPreElement;

const translateChangedParagraphsButton =
    document.getElementById(
        "translate-changed-paragraphs"
    ) as HTMLButtonElement;

const markRevisionsSeenButton =
    document.getElementById("mark-revisions-seen") as HTMLButtonElement;

const sourceColumnSelect =
    document.getElementById("source-column") as HTMLSelectElement;

const saveSourceColumnButton =
    document.getElementById("save-source-column") as HTMLButtonElement;

const columnLanguageSelects = [1, 2, 3].map((column) =>
    document.getElementById(
        `column-${column}-language`
    ) as HTMLSelectElement
);

const REVISION_BASELINE_KEY =
    "translationTool.revisionBaseline.v1";

// The document's language columns use zero-based table indices 1, 2, and 3.
// Column 0 and every column after 3 contain unrelated document metadata.
const LANGUAGE_COLUMN_INDICES = [1, 2, 3] as const;
const INSPECTION_ROW_BATCH_SIZE = 10;
const TRANSLATION_REQUEST_TIMEOUT_MS = 180_000;


Office.onReady((info) => {
    if (info.host !== Office.HostType.Word) {
        output.textContent = "This add-in must run inside Word.";
        return;
    }

    output.textContent = "Connected to Word.";

    translateChangedParagraphsButton.addEventListener(
        "click",
        translateChangedTableParagraphs
    );

    markRevisionsSeenButton.addEventListener(
        "click",
        markCurrentRevisionsAsSeen
    );

    saveSourceColumnButton.addEventListener(
        "click",
        saveSourceColumn
    );

    sourceColumnSelect.addEventListener(
        "change",
        alignEnglishLanguageWithSourceColumn
    );

    void loadSourceColumn();
});


type Language = "English" | "French" | "German";


type TranslationConfig = {
    source_column: number;
    column_1_language: Language;
    column_2_language: Language;
    column_3_language: Language;
};


async function getErrorDetail(response: Response): Promise<string> {
    try {
        const body = await response.json() as { detail?: unknown };
        return typeof body.detail === "string"
            ? body.detail
            : JSON.stringify(body.detail ?? body);
    } catch {
        return await response.text();
    }
}


async function getTranslationConfig(): Promise<TranslationConfig> {
    const response = await fetch("/api/config");

    if (!response.ok) {
        throw new Error(
            `Could not load configuration: ${await getErrorDetail(response)}`
        );
    }

    return await response.json() as TranslationConfig;
}


async function loadSourceColumn(): Promise<void> {
    try {
        const config = await getTranslationConfig();
        sourceColumnSelect.value = String(config.source_column);
        columnLanguageSelects[0].value = config.column_1_language;
        columnLanguageSelects[1].value = config.column_2_language;
        columnLanguageSelects[2].value = config.column_3_language;
    } catch (error) {
        console.error(error);
        output.textContent = `Error: ${String(error)}`;
    }
}


function alignEnglishLanguageWithSourceColumn(): void {
    const sourceCellIndex = Number(sourceColumnSelect.value) - 1;
    const previousEnglishIndex = columnLanguageSelects.findIndex(
        (select) => select.value === "English"
    );
    const selectedPreviousLanguage =
        columnLanguageSelects[sourceCellIndex].value;

    columnLanguageSelects[sourceCellIndex].value = "English";
    if (previousEnglishIndex >= 0
        && previousEnglishIndex !== sourceCellIndex) {
        columnLanguageSelects[previousEnglishIndex].value =
            selectedPreviousLanguage;
    }
}


async function saveSourceColumn(): Promise<void> {
    saveSourceColumnButton.disabled = true;

    try {
        const sourceColumn = Number(sourceColumnSelect.value);
        alignEnglishLanguageWithSourceColumn();
        const languages = columnLanguageSelects.map(
            (select) => select.value as Language
        );
        if (new Set(languages).size !== 3) {
            throw new Error(
                "Columns 1-3 must use English, French, and German exactly once."
            );
        }

        const response = await fetch("/api/config", {
            method: "PUT",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                source_column: sourceColumn,
                column_1_language: languages[0],
                column_2_language: languages[1],
                column_3_language: languages[2]
            })
        });

        if (!response.ok) {
            throw new Error(await getErrorDetail(response));
        }

        output.textContent =
            `Language columns saved: ${languages.join(" / ")}.`;
    } catch (error) {
        console.error(error);
        output.textContent = `Error: ${String(error)}`;
    } finally {
        saveSourceColumnButton.disabled = false;
    }
}


async function requestCellTranslation(
    request: TranslateCellChangesRequest,
    requestId: string
): Promise<TranslateCellChangesResponse> {
    const controller = new AbortController();
    const timeout = window.setTimeout(
        () => controller.abort(),
        TRANSLATION_REQUEST_TIMEOUT_MS
    );

    try {
        const response = await fetch("/api/translate-cell-changes", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "X-Request-ID": requestId
            },
            body: JSON.stringify(request),
            signal: controller.signal
        });

        if (!response.ok) {
            throw new Error(
                `Translation failed: ${await getErrorDetail(response)}`
            );
        }

        return await response.json() as TranslateCellChangesResponse;
    } catch (error) {
        if (error instanceof DOMException && error.name === "AbortError") {
            throw new Error(
                `Translation request ${requestId} timed out after 180 seconds.`
            );
        }

        throw error;
    } finally {
        window.clearTimeout(timeout);
    }
}


type PropagationResult = {
    tables: number;
    eligibleRows: number;
    sourceParagraphs: number;
    translatedParagraphs: number;
    updatedTargets: number;
    unchangedTargets: number;
    insertedTargets: number;
    minimallyTrackedTargets: number;
    fullParagraphFallbacks: number;
    ignoredWhitespaceOnlyChanges: number;
    numericRejected: number;
    failedCellPlans: number;
    partialCellPlans: number;
    translationErrors: string[];
    translationWarnings: string[];
    invalidRows: number;
    unequalParagraphCountRows: number;
    structuralSources: number;
    baselineCreated: boolean;
    baselineDeferred: boolean;
};


type InspectedParagraph = {
    paragraph: Word.Paragraph;
    trackedChanges: Word.TrackedChangeCollection;
    currentText: OfficeExtension.ClientResult<string>;
    originalText: OfficeExtension.ClientResult<string>;
    originalPosition?: number;
    currentPosition?: number;
    newRevisions: RevisionInput[];
    newRevisionFingerprints: string[];
};


type RevisionInput = {
    type: "Added" | "Deleted" | "Formatted" | "None";
    text: string;
};


type CellParagraphInput = {
    index: number;
    text: string;
};


type ChangedSourceParagraphInput = {
    index: number;
    original_text: string;
    current_text: string;
    changes: RevisionInput[];
};


type TargetCellInput = {
    column: number;
    expected_language: "French" | "German";
    paragraphs: CellParagraphInput[];
};


type TranslateCellChangesRequest = {
    source_column: number;
    source_cell: CellParagraphInput[];
    changed_source_paragraphs: ChangedSourceParagraphInput[];
    targets: TargetCellInput[];
};


type CellEditOutput = {
    source_paragraph_indices: number[];
    operation: "replace" | "insert" | "none";
    target_paragraph_index: number;
    original_text: string;
    translated_text: string;
};


type CellTranslationOutput = {
    column: number;
    language: "French" | "German";
    edits: CellEditOutput[];
};


type TranslateCellChangesResponse = {
    translations: CellTranslationOutput[];
    numeric_consistent: boolean;
    numeric_warnings: string[];
    partial_errors: string[];
    failed_source_paragraph_indices: number[];
};


type ReplaceMutation = {
    operation: "replace";
    paragraph: Word.Paragraph;
    originalText: string;
    text: string;
};


type InsertMutation = {
    operation: "insert";
    cellParagraphs: InspectedParagraph[];
    meaningfulParagraphs: InspectedParagraph[];
    insertionIndex: number;
    text: string;
};


type PlannedMutation = ReplaceMutation | InsertMutation;


type SearchCandidate = {
    ranges: Word.RangeCollection;
    expectedText: string;
    insertLocation: Word.InsertLocation;
};


type PreparedReplacement = {
    replacementText: string;
    candidates: SearchCandidate[];
};


type RevisionCounts = Record<string, number>;


type RevisionBaseline = {
    version: 1;
    fingerprints: RevisionCounts;
};


function hasMeaningfulParagraphText(value: string): boolean {
    // Word can expose paragraph, line, and table-cell markers as characters.
    // None of them make a blank visual line a logical paragraph for matching.
    return value.replace(/[\s\u0007]/gu, "").length > 0;
}


function normalizeWhitespace(value: string): string {
    return value.replace(/[\s\u0007]+/gu, " ").trim();
}


function hasSubstantiveNewRevision(
    revisions: RevisionInput[]
): boolean {
    return revisions.some((revision) =>
        (revision.type === "Added" || revision.type === "Deleted")
        && revision.text.replace(/[\s\u0007]+/gu, "").length > 0
    );
}


function isWholeParagraphAddition(
    revisions: RevisionInput[],
    currentText: string
): boolean {
    const normalizedCurrent = normalizeWhitespace(currentText);
    if (!normalizedCurrent) {
        return false;
    }

    const addedText = revisions
        .filter((revision) => revision.type === "Added")
        .map((revision) => revision.text)
        .join("");
    const hasSubstantiveDeletion = revisions.some((revision) =>
        revision.type === "Deleted"
        && revision.text.replace(/[\s\u0007]+/gu, "").length > 0
    );

    return !hasSubstantiveDeletion
        && normalizeWhitespace(addedText) === normalizedCurrent;
}


function getSingleSpanDiff(before: string, after: string): {
    prefixLength: number;
    suffixLength: number;
    removedText: string;
    insertedText: string;
} {
    let prefixLength = 0;
    const commonLength = Math.min(before.length, after.length);

    while (prefixLength < commonLength
        && before[prefixLength] === after[prefixLength]) {
        prefixLength += 1;
    }

    let suffixLength = 0;
    const remainingBefore = before.length - prefixLength;
    const remainingAfter = after.length - prefixLength;
    const maximumSuffix = Math.min(remainingBefore, remainingAfter);

    while (suffixLength < maximumSuffix
        && before[before.length - suffixLength - 1]
            === after[after.length - suffixLength - 1]) {
        suffixLength += 1;
    }

    return {
        prefixLength,
        suffixLength,
        removedText: before.slice(
            prefixLength,
            before.length - suffixLength
        ),
        insertedText: after.slice(
            prefixLength,
            after.length - suffixLength
        )
    };
}


function formatElapsed(startedAt: number): string {
    const totalSeconds = Math.floor((performance.now() - startedAt) / 1000);
    const minutes = Math.floor(totalSeconds / 60);
    const seconds = totalSeconds % 60;

    return minutes > 0
        ? `${minutes}m ${seconds}s`
        : `${seconds}s`;
}


async function showProgress(
    startedAt: number,
    stage: string,
    details: string[] = []
): Promise<void> {
    const lines = [
        stage,
        ...details,
        `Elapsed: ${formatElapsed(startedAt)}`
    ];

    output.textContent = lines.join("\n");
    console.info("[tracked-change propagation]", ...lines);

    // Give Word's task pane a chance to paint before a potentially expensive
    // Office.js synchronization starts.
    await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
}


async function syncWithProgress(
    context: Word.RequestContext,
    startedAt: number,
    stage: string,
    details: string[] = []
): Promise<void> {
    await showProgress(startedAt, stage, details);
    const syncStartedAt = performance.now();

    await context.sync();

    console.info(
        "[tracked-change propagation] Word sync completed",
        stage,
        `${Math.round(performance.now() - syncStartedAt)}ms`
    );
}


function parseRevisionBaseline(value: unknown): RevisionCounts {
    if (typeof value !== "string") {
        return {};
    }

    try {
        const parsed = JSON.parse(value) as Partial<RevisionBaseline>;

        if (parsed.version !== 1
            || !parsed.fingerprints
            || typeof parsed.fingerprints !== "object") {
            return {};
        }

        return parsed.fingerprints;
    } catch {
        return {};
    }
}


async function sha256(value: string): Promise<string> {
    const bytes = new TextEncoder().encode(value);
    const digest = await crypto.subtle.digest("SHA-256", bytes);

    return Array.from(new Uint8Array(digest))
        .map((byte) => byte.toString(16).padStart(2, "0"))
        .join("");
}


async function getRevisionFingerprint(
    change: Word.TrackedChange
): Promise<string> {
    return sha256([
        change.author,
        change.date.toISOString(),
        change.type,
        change.text
    ].join("\u0000"));
}


function addFingerprint(
    counts: RevisionCounts,
    fingerprint: string
): void {
    counts[fingerprint] = (counts[fingerprint] ?? 0) + 1;
}


async function getCurrentRevisionCounts(
    context: Word.RequestContext
): Promise<RevisionCounts> {
    const revisions = context.document.body.getTrackedChanges();
    revisions.load("items/author,items/date,items/text,items/type");
    await context.sync();

    const fingerprints = await Promise.all(
        revisions.items.map(getRevisionFingerprint)
    );
    const counts: RevisionCounts = {};

    for (const fingerprint of fingerprints) {
        addFingerprint(counts, fingerprint);
    }

    return counts;
}


function saveRevisionBaseline(
    context: Word.RequestContext,
    fingerprints: RevisionCounts
): void {
    const baseline: RevisionBaseline = {
        version: 1,
        fingerprints
    };

    context.document.settings.add(
        REVISION_BASELINE_KEY,
        JSON.stringify(baseline)
    );
}


async function markCurrentRevisionsAsSeen(): Promise<void> {
    markRevisionsSeenButton.disabled = true;
    translateChangedParagraphsButton.disabled = true;
    output.textContent = "Recording the current revision baseline...";

    try {
        const revisionCount = await Word.run(async (context) => {
            const fingerprints = await getCurrentRevisionCounts(context);
            saveRevisionBaseline(context, fingerprints);
            await context.sync();

            return Object.values(fingerprints)
                .reduce((total, count) => total + count, 0);
        });

        output.textContent =
            `Recorded ${revisionCount} current revisions as seen.\n`
            + "Make a new tracked edit, then run propagation.";
    } catch (error) {
        console.error(error);
        output.textContent = `Error: ${String(error)}`;
    } finally {
        markRevisionsSeenButton.disabled = false;
        translateChangedParagraphsButton.disabled = false;
    }
}


async function translateChangedTableParagraphs(): Promise<void> {
    if (!Office.context.requirements.isSetSupported("WordApi", "1.6")) {
        output.textContent =
            "This version of Word does not support tracked-change inspection (WordApi 1.6).";
        return;
    }

    translateChangedParagraphsButton.disabled = true;
    saveSourceColumnButton.disabled = true;
    const startedAt = performance.now();
    await showProgress(startedAt, "Loading translation configuration...");

    try {
        const translationConfig = await getTranslationConfig();
        const sourceColumn = translationConfig.source_column;
        sourceColumnSelect.value = String(sourceColumn);
        const configuredLanguages: Language[] = [
            translationConfig.column_1_language,
            translationConfig.column_2_language,
            translationConfig.column_3_language
        ];
        configuredLanguages.forEach((language, index) => {
            columnLanguageSelects[index].value = language;
        });
        const sourceCellIndex = sourceColumn - 1;
        const result = await Word.run(async (context): Promise<PropagationResult> => {
            const wordDocument = context.document;
            const tables = wordDocument.body.tables;
            const baselineSetting = wordDocument.settings
                .getItemOrNullObject(REVISION_BASELINE_KEY);

            wordDocument.load("changeTrackingMode");
            tables.load("items");
            baselineSetting.load("value");
            await syncWithProgress(
                context,
                startedAt,
                "Finding document tables and the saved revision baseline..."
            );

            const baselineCreated = baselineSetting.isNullObject;
            const savedRevisionCounts = baselineCreated
                ? {}
                : parseRevisionBaseline(baselineSetting.value);
            const remainingKnownRevisions: RevisionCounts = {
                ...savedRevisionCounts
            };
            const deferredRevisionCounts: RevisionCounts = {};

            const deferParagraphRevisions = (
                paragraph: InspectedParagraph
            ): void => {
                for (const fingerprint of paragraph.newRevisionFingerprints) {
                    addFingerprint(deferredRevisionCounts, fingerprint);
                }
            };

            const tableRows = tables.items.map((table) => {
                table.rows.load("items/cellCount");
                return table.rows;
            });

            await syncWithProgress(
                context,
                startedAt,
                "Reading table row metadata...",
                [`Tables found: ${tables.items.length}`]
            );

            const result: PropagationResult = {
                tables: tables.items.length,
                eligibleRows: 0,
                sourceParagraphs: 0,
                translatedParagraphs: 0,
                updatedTargets: 0,
                unchangedTargets: 0,
                insertedTargets: 0,
                minimallyTrackedTargets: 0,
                fullParagraphFallbacks: 0,
                ignoredWhitespaceOnlyChanges: 0,
                numericRejected: 0,
                failedCellPlans: 0,
                partialCellPlans: 0,
                translationErrors: [],
                translationWarnings: [],
                invalidRows: 0,
                unequalParagraphCountRows: 0,
                structuralSources: 0,
                baselineCreated,
                baselineDeferred: false
            };

            const eligibleRows: Word.TableRow[] = [];

            for (const rows of tableRows) {
                for (const row of rows.items) {
                    if (row.cellCount <= LANGUAGE_COLUMN_INDICES[2]) {
                        result.invalidRows += 1;
                        continue;
                    }

                    eligibleRows.push(row);
                }
            }

            result.eligibleRows = eligibleRows.length;
            const inspectedRows: InspectedParagraph[][][] = [];
            let inspectedParagraphCount = 0;

            for (let batchStart = 0;
                batchStart < eligibleRows.length;
                batchStart += INSPECTION_ROW_BATCH_SIZE) {
                const batchEnd = Math.min(
                    batchStart + INSPECTION_ROW_BATCH_SIZE,
                    eligibleRows.length
                );
                const rowBatch = eligibleRows.slice(batchStart, batchEnd);
                const batchLabel = `Rows ${batchStart + 1}-${batchEnd} of ${eligibleRows.length}`;

                for (const row of rowBatch) {
                    row.cells.load("items");
                }

                await syncWithProgress(
                    context,
                    startedAt,
                    "Loading language cells...",
                    [batchLabel, `Paragraphs inspected: ${inspectedParagraphCount}`]
                );

                const rowParagraphCollections = rowBatch.map((row) => {
                    const collections = LANGUAGE_COLUMN_INDICES.map((column) => {
                        const cell = row.cells.items[column];
                        const paragraphs = cell.body.paragraphs;
                        paragraphs.load("items");
                        return paragraphs;
                    });

                    return collections;
                });

                await syncWithProgress(
                    context,
                    startedAt,
                    "Loading paragraph lists...",
                    [batchLabel, `Paragraphs inspected: ${inspectedParagraphCount}`]
                );

                const inspectedBatch = rowParagraphCollections.map(
                    (collections) => collections.map((collection) =>
                        collection.items.map((paragraph): InspectedParagraph => {
                            const trackedChanges = paragraph.getTrackedChanges();
                            const currentText = paragraph.getReviewedText(
                                Word.ChangeTrackingVersion.current
                            );
                            const originalText = paragraph.getReviewedText(
                                Word.ChangeTrackingVersion.original
                            );

                            trackedChanges.load(
                                "items/author,items/date,items/text,items/type"
                            );

                            return {
                                paragraph,
                                trackedChanges,
                                currentText,
                                originalText,
                                newRevisions: [],
                                newRevisionFingerprints: []
                            };
                        })
                    )
                );

                const batchParagraphCount = inspectedBatch.reduce(
                    (rowTotal, cells) => rowTotal + cells.reduce(
                        (cellTotal, paragraphs) => cellTotal + paragraphs.length,
                        0
                    ),
                    0
                );

                await syncWithProgress(
                    context,
                    startedAt,
                    "Inspecting tracked changes and paragraph text...",
                    [
                        batchLabel,
                        `This batch: ${batchParagraphCount} paragraphs`,
                        `Previously inspected: ${inspectedParagraphCount}`
                    ]
                );

                inspectedParagraphCount += batchParagraphCount;

                for (const cells of inspectedBatch) {
                    const meaningfulParagraphCounts = cells.map((cell) =>
                        cell.filter((paragraph) => hasMeaningfulParagraphText(
                            paragraph.originalText.value
                        )).length
                    );

                    if (!meaningfulParagraphCounts.every(
                        (count) => count === meaningfulParagraphCounts[0]
                    )) {
                        result.unequalParagraphCountRows += 1;
                    }

                    for (const cell of cells) {
                        let originalPosition = 0;
                        let currentPosition = 0;

                        for (const inspected of cell) {
                            if (hasMeaningfulParagraphText(
                                inspected.originalText.value
                            )) {
                                inspected.originalPosition = originalPosition;
                                originalPosition += 1;
                            }

                            if (hasMeaningfulParagraphText(
                                inspected.currentText.value
                            )) {
                                inspected.currentPosition = currentPosition;
                                currentPosition += 1;
                            }

                            const fingerprints = await Promise.all(
                                inspected.trackedChanges.items.map(
                                    getRevisionFingerprint
                                )
                            );

                            for (let revisionIndex = 0;
                                revisionIndex < fingerprints.length;
                                revisionIndex += 1) {
                                const fingerprint = fingerprints[revisionIndex];
                                const knownCount =
                                    remainingKnownRevisions[fingerprint] ?? 0;

                                if (knownCount > 0) {
                                    remainingKnownRevisions[fingerprint] =
                                        knownCount - 1;
                                } else {
                                    const revision = inspected
                                        .trackedChanges.items[revisionIndex];

                                    inspected.newRevisions.push({
                                        type: revision.type,
                                        text: revision.text
                                    });
                                    inspected.newRevisionFingerprints.push(
                                        fingerprint
                                    );
                                }
                            }
                        }
                    }
                }

                inspectedRows.push(...inspectedBatch);
            }

            await showProgress(
                startedAt,
                "Tracked-change inspection complete.",
                [
                    `Rows inspected: ${result.eligibleRows}`,
                    `Paragraphs inspected: ${inspectedParagraphCount}`
                ]
            );

            const mutations: PlannedMutation[] = [];

            for (const cells of inspectedRows) {
                const currentParagraphs = cells.map((cell) =>
                    cell.filter((paragraph) =>
                        paragraph.currentPosition !== undefined
                    ).sort((left, right) =>
                        left.currentPosition! - right.currentPosition!
                    )
                );
                const sourceParagraphs = currentParagraphs[sourceCellIndex];
                const changedSourceParagraphs: ChangedSourceParagraphInput[] = [];

                for (const source of cells[sourceCellIndex]) {
                    if (source.newRevisions.length === 0) {
                        continue;
                    }

                    const isNewParagraph = !hasMeaningfulParagraphText(
                        source.originalText.value
                    ) && hasMeaningfulParagraphText(
                        source.currentText.value
                    ) || isWholeParagraphAddition(
                        source.newRevisions,
                        source.currentText.value
                    );
                    if (!isNewParagraph
                        && (!hasSubstantiveNewRevision(source.newRevisions)
                            || normalizeWhitespace(
                                source.originalText.value
                            ) === normalizeWhitespace(
                                source.currentText.value
                            ))) {
                        result.ignoredWhitespaceOnlyChanges += 1;
                        continue;
                    }

                    if (source.currentPosition === undefined
                        || !hasMeaningfulParagraphText(
                            source.currentText.value
                        )) {
                        // Fully deleted source paragraphs remain intentionally
                        // untouched, matching the existing product behavior.
                        result.structuralSources += 1;
                        deferParagraphRevisions(source);
                        continue;
                    }

                    changedSourceParagraphs.push({
                        index: source.currentPosition,
                        original_text: source.originalText.value,
                        current_text: source.currentText.value,
                        changes: source.newRevisions
                    });
                }

                if (changedSourceParagraphs.length === 0) {
                    continue;
                }

                const targets: TargetCellInput[] = [];
                for (let targetCellIndex = 0;
                    targetCellIndex < cells.length;
                    targetCellIndex += 1) {
                    if (targetCellIndex === sourceCellIndex) {
                        continue;
                    }

                    targets.push({
                        column: targetCellIndex + 1,
                        expected_language: configuredLanguages[
                            targetCellIndex
                        ] as "French" | "German",
                        paragraphs: currentParagraphs[targetCellIndex].map(
                            (paragraph) => ({
                                index: paragraph.currentPosition!,
                                text: paragraph.currentText.value
                            })
                        )
                    });
                }

                result.sourceParagraphs += changedSourceParagraphs.length;
                const requestId = crypto.randomUUID();
                await showProgress(
                    startedAt,
                    "Waiting for Claude to map changed paragraphs across complete cells...",
                    [
                        `Changed source paragraphs found: ${result.sourceParagraphs}`,
                        `This cell: ${changedSourceParagraphs.length}`,
                        `Completed translations: ${result.translatedParagraphs}`,
                        `Request: ${requestId}`
                    ]
                );
                let translation: TranslateCellChangesResponse;
                try {
                    translation = await requestCellTranslation({
                        source_column: sourceColumn,
                        source_cell: sourceParagraphs.map((paragraph) => ({
                            index: paragraph.currentPosition!,
                            text: paragraph.currentText.value
                        })),
                        changed_source_paragraphs: changedSourceParagraphs,
                        targets
                    }, requestId);
                } catch (error) {
                    const message = String(error);
                    result.failedCellPlans += 1;
                    result.translationErrors.push(message);
                    const failedIndices = new Set(
                        changedSourceParagraphs.map((paragraph) =>
                            paragraph.index
                        )
                    );
                    for (const source of cells[sourceCellIndex]) {
                        if (source.currentPosition !== undefined
                            && failedIndices.has(source.currentPosition)) {
                            deferParagraphRevisions(source);
                        }
                    }
                    console.error(
                        `[tracked-change propagation] ${requestId}`,
                        error
                    );
                    continue;
                }

                if (translation.partial_errors.length > 0) {
                    result.partialCellPlans += 1;
                    result.translationWarnings.push(
                        ...translation.partial_errors
                    );
                }

                const failedSourceIndices = new Set(
                    translation.failed_source_paragraph_indices ?? []
                );
                for (const source of cells[sourceCellIndex]) {
                    if (source.currentPosition !== undefined
                        && failedSourceIndices.has(source.currentPosition)) {
                        deferParagraphRevisions(source);
                    }
                }

                if (!translation.numeric_consistent) {
                    result.numericRejected += 1;
                    result.translationWarnings.push(
                        ...translation.numeric_warnings
                    );
                    console.warn(...translation.numeric_warnings);
                }

                result.translatedParagraphs += changedSourceParagraphs.length;

                for (const translatedCell of translation.translations) {
                    const targetCellIndex = translatedCell.column - 1;
                    const targetParagraphs = currentParagraphs[targetCellIndex];

                    if (targetCellIndex === sourceCellIndex
                        || !cells[targetCellIndex]) {
                        throw new Error(
                            "Backend returned an unexpected target column."
                        );
                    }

                    for (const edit of translatedCell.edits) {
                        if (edit.operation === "none") {
                            result.unchangedTargets += 1;
                            continue;
                        }

                        if (edit.operation === "replace") {
                            const paragraph =
                                targetParagraphs[edit.target_paragraph_index];

                            if (!paragraph
                                || paragraph.currentText.value
                                    !== edit.original_text) {
                                throw new Error(
                                    "Backend returned a stale target paragraph."
                                );
                            }

                            if (edit.translated_text === edit.original_text) {
                                result.unchangedTargets += 1;
                                continue;
                            }

                            mutations.push({
                                operation: "replace",
                                paragraph: paragraph.paragraph,
                                originalText: edit.original_text,
                                text: edit.translated_text
                            });
                            continue;
                        }

                        if (edit.target_paragraph_index
                            > targetParagraphs.length) {
                            throw new Error(
                                "Backend returned an invalid insertion position."
                            );
                        }

                        mutations.push({
                            operation: "insert",
                            cellParagraphs: cells[targetCellIndex],
                            meaningfulParagraphs: targetParagraphs,
                            insertionIndex: edit.target_paragraph_index,
                            text: edit.translated_text
                        });
                        result.insertedTargets += 1;
                    }
                }
            }

            const preparedReplacements = new Map<
                Word.Paragraph,
                PreparedReplacement
            >();
            let narrowSearchCount = 0;

            for (const mutation of mutations) {
                if (mutation.operation !== "replace") {
                    continue;
                }

                const diff = getSingleSpanDiff(
                    mutation.originalText,
                    mutation.text
                );
                const candidates: SearchCandidate[] = [];

                // No common context means Claude effectively changed the
                // entire paragraph, so a whole-paragraph tracked replacement
                // accurately represents the result.
                if (diff.prefixLength === 0 && diff.suffixLength === 0) {
                    continue;
                }

                if (diff.removedText.length > 0
                    && diff.removedText.length <= 200) {
                    const ranges = mutation.paragraph.search(
                        diff.removedText,
                        {
                            ignorePunct: false,
                            ignoreSpace: false,
                            matchCase: true,
                            matchWholeWord: false,
                            matchWildcards: false
                        }
                    );
                    ranges.load("items/text");
                    candidates.push({
                        ranges,
                        expectedText: diff.removedText,
                        insertLocation: Word.InsertLocation.replace
                    });
                } else if (diff.removedText.length === 0) {
                    if (diff.suffixLength > 0) {
                        const suffixAnchor = mutation.originalText.slice(
                            diff.prefixLength,
                            diff.prefixLength + Math.min(
                                diff.suffixLength,
                                80
                            )
                        );
                        const ranges = mutation.paragraph.search(
                            suffixAnchor,
                            {
                                ignorePunct: false,
                                ignoreSpace: false,
                                matchCase: true,
                                matchWholeWord: false,
                                matchWildcards: false
                            }
                        );
                        ranges.load("items/text");
                        candidates.push({
                            ranges,
                            expectedText: suffixAnchor,
                            insertLocation: Word.InsertLocation.before
                        });
                    }

                    if (diff.prefixLength > 0) {
                        const prefixAnchor = mutation.originalText.slice(
                            Math.max(0, diff.prefixLength - 80),
                            diff.prefixLength
                        );
                        const ranges = mutation.paragraph.search(
                            prefixAnchor,
                            {
                                ignorePunct: false,
                                ignoreSpace: false,
                                matchCase: true,
                                matchWholeWord: false,
                                matchWildcards: false
                            }
                        );
                        ranges.load("items/text");
                        candidates.push({
                            ranges,
                            expectedText: prefixAnchor,
                            insertLocation: Word.InsertLocation.after
                        });
                    }
                }

                if (candidates.length > 0) {
                    narrowSearchCount += candidates.length;
                    preparedReplacements.set(mutation.paragraph, {
                        replacementText: diff.insertedText,
                        candidates
                    });
                }
            }

            if (narrowSearchCount > 0) {
                await syncWithProgress(
                    context,
                    startedAt,
                    "Locating minimal changed spans inside target paragraphs...",
                    [`Candidate ranges: ${narrowSearchCount}`]
                );
            }

            const originalTrackingMode = wordDocument.changeTrackingMode;
            const mustRestoreTrackingMode =
                originalTrackingMode === Word.ChangeTrackingMode.off;

            if (mutations.length > 0 && mustRestoreTrackingMode) {
                wordDocument.changeTrackingMode =
                    Word.ChangeTrackingMode.trackAll;
                await syncWithProgress(
                    context,
                    startedAt,
                    "Enabling Track Changes for translated replacements..."
                );
            }

            try {
                // Work backwards so an earlier insertion cannot move a later
                // paragraph before all mutations have been queued.
                for (const mutation of mutations.reverse()) {
                    if (mutation.operation === "replace") {
                        const prepared = preparedReplacements.get(
                            mutation.paragraph
                        );
                        const exactCandidate = prepared?.candidates.find(
                            (candidate) => candidate.ranges.items.filter(
                                (range) => range.text
                                    === candidate.expectedText
                            ).length === 1
                        );

                        if (prepared && exactCandidate) {
                            const exactRange = exactCandidate.ranges.items.find(
                                (range) => range.text
                                    === exactCandidate.expectedText
                            )!;
                            exactRange.insertText(
                                prepared.replacementText,
                                exactCandidate.insertLocation
                            );
                            result.minimallyTrackedTargets += 1;
                        } else {
                            mutation.paragraph.insertText(
                                mutation.text,
                                Word.InsertLocation.replace
                            );
                            result.fullParagraphFallbacks += 1;
                        }
                        continue;
                    }

                    if (mutation.meaningfulParagraphs.length === 0) {
                        const emptyParagraph = mutation.cellParagraphs[0];
                        if (!emptyParagraph) {
                            throw new Error(
                                "Word returned a table cell without a paragraph."
                            );
                        }

                        emptyParagraph.paragraph.insertText(
                            mutation.text,
                            Word.InsertLocation.replace
                        );
                    } else if (mutation.insertionIndex
                        < mutation.meaningfulParagraphs.length) {
                        mutation.meaningfulParagraphs[
                            mutation.insertionIndex
                        ].paragraph.insertParagraph(
                            mutation.text,
                            Word.InsertLocation.before
                        );
                    } else {
                        mutation.meaningfulParagraphs[
                            mutation.meaningfulParagraphs.length - 1
                        ].paragraph.insertParagraph(
                            mutation.text,
                            Word.InsertLocation.after
                        );
                    }
                }

                await syncWithProgress(
                    context,
                    startedAt,
                    "Writing translated paragraphs...",
                    [`Target paragraph edits: ${mutations.length}`]
                );
                result.updatedTargets = mutations.length;
            } finally {
                if (mutations.length > 0 && mustRestoreTrackingMode) {
                    wordDocument.changeTrackingMode = originalTrackingMode;
                    await syncWithProgress(
                        context,
                        startedAt,
                        "Restoring the document's Track Changes setting..."
                    );
                }
            }

            result.baselineDeferred = Object.keys(
                deferredRevisionCounts
            ).length > 0;

            const currentRevisionCounts =
                await getCurrentRevisionCounts(context);
            for (const [fingerprint, deferredCount] of Object.entries(
                deferredRevisionCounts
            )) {
                const completedCount = (
                    currentRevisionCounts[fingerprint] ?? 0
                ) - deferredCount;
                if (completedCount > 0) {
                    currentRevisionCounts[fingerprint] = completedCount;
                } else {
                    delete currentRevisionCounts[fingerprint];
                }
            }
            saveRevisionBaseline(context, currentRevisionCounts);
            await syncWithProgress(
                context,
                startedAt,
                result.baselineDeferred
                    ? "Saving completed revisions and retaining failed revisions for retry..."
                    : "Saving the completed revision baseline..."
            );

            return result;
        });

        output.textContent = [
            `Tables inspected: ${result.tables}`,
            `Rows containing language columns 1-3: ${result.eligibleRows}`,
            `New English source paragraphs: ${result.sourceParagraphs}`,
            `Paragraphs translated by Claude: ${result.translatedParagraphs}`,
            `Target paragraphs updated with Track Changes: ${result.updatedTargets}`,
            `Target paragraphs edited only at changed spans: ${result.minimallyTrackedTargets}`,
            `Whole-paragraph safety fallbacks: ${result.fullParagraphFallbacks}`,
            `Whitespace-only source changes ignored: ${result.ignoredWhitespaceOnlyChanges}`,
            `New target paragraphs inserted: ${result.insertedTargets}`,
            `Unchanged translations skipped: ${result.unchangedTargets}`,
            `Cell plans applied with numeric warnings: ${result.numericRejected}`,
            `Cell plans that failed: ${result.failedCellPlans}`,
            `Partially completed cell plans: ${result.partialCellPlans}`,
            `Rows missing language columns 1-3: ${result.invalidRows}`,
            `Rows with unequal paragraph counts (processed): ${result.unequalParagraphCountRows}`,
            `Fully deleted source paragraphs skipped: ${result.structuralSources}`,
            result.baselineDeferred
                ? "Completed revisions recorded; failed/skipped revisions retained for retry."
                : "Revision baseline advanced through this run.",
            result.baselineCreated
                ? "No prior baseline: all existing revisions were treated as new."
                : "Only revisions newer than the saved baseline were treated as sources.",
            ...result.translationErrors.slice(0, 3).map(
                (error, index) => `Failed plan ${index + 1}: ${error}`
            ),
            ...result.translationWarnings.slice(0, 5).map(
                (warning, index) => `Warning ${index + 1}: ${warning}`
            ),
            `Completed in ${formatElapsed(startedAt)}.`
        ].join("\n");
    } catch (error) {
        console.error(error);
        output.textContent = `Error: ${String(error)}`;
    } finally {
        translateChangedParagraphsButton.disabled = false;
        saveSourceColumnButton.disabled = false;
    }
}
