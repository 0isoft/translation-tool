import {
    getTranslationConfig,
    saveTranslationConfig
} from "./adapters/configurationStore";
import {
    assertTranslationApiAvailable,
    isNetworkLoadFailure,
    requestCellTranslation
} from "./adapters/translationApi";
import {
    REVISION_BASELINE_KEY,
    addFingerprint,
    currentRevisionCounts as getCurrentRevisionCounts,
    extractOoxmlParagraphInsertions,
    parseRevisionBaseline,
    revisionFingerprint as getRevisionFingerprint,
    saveRevisionBaseline
} from "./adapters/wordRevisionStore";
import {
    applyTrackedMutations,
    type InspectedParagraph,
    type PlannedMutation
} from "./adapters/wordMutationWriter";
import type {
    ChangedSourceParagraphInput,
    Language,
    RevisionCounts,
    RevisionInput,
    TargetCellInput,
    TranslateCellChangesResponse,
    TranslationConfig
} from "./domain/models";
import {
    getSingleSpanDiff,
    hasMeaningfulParagraphText,
    hasSubstantiveRevision,
    isWholeParagraphAddition,
    normalizeWhitespace
} from "./domain/textRules";

const output =
    document.getElementById("output") as HTMLPreElement;

const translateChangedParagraphsButton =
    document.getElementById(
        "translate-changed-paragraphs"
    ) as HTMLButtonElement;

const markRevisionsSeenButton =
    document.getElementById("mark-revisions-seen") as HTMLButtonElement;

const retryAllRevisionsButton =
    document.getElementById("retry-all-revisions") as HTMLButtonElement;

const sourceColumnSelect =
    document.getElementById("source-column") as HTMLSelectElement;

const saveSourceColumnButton =
    document.getElementById("save-source-column") as HTMLButtonElement;

const configurationStatus =
    document.getElementById("configuration-status") as HTMLParagraphElement;

const columnLanguageSelects = [1, 2, 3].map((column) =>
    document.getElementById(
        `column-${column}-language`
    ) as HTMLSelectElement
);

// The document's language columns use zero-based table indices 1, 2, and 3.
// Column 0 and every column after 3 contain unrelated document metadata.
const STANDARD_LANGUAGE_COLUMN_INDICES = [1, 2, 3] as const;
const COMPACT_LANGUAGE_COLUMN_INDICES = [0, 1, 2] as const;
const INSPECTION_ROW_BATCH_SIZE = 10;
const MAX_CONSECUTIVE_NETWORK_FAILURES = 3;


Office.onReady((info) => {
    if (info.host !== Office.HostType.Word) {
        output.textContent = "This add-in must run inside Word.";
        return;
    }

    output.textContent = "Connected to Word.";

    translateChangedParagraphsButton.addEventListener(
        "click",
        () => void translateChangedTableParagraphs(false)
    );

    markRevisionsSeenButton.addEventListener(
        "click",
        markCurrentRevisionsAsSeen
    );

    retryAllRevisionsButton.addEventListener(
        "click",
        () => void translateChangedTableParagraphs(true)
    );

    saveSourceColumnButton.addEventListener(
        "click",
        saveSourceColumn
    );

    sourceColumnSelect.addEventListener(
        "change",
        markConfigurationUnsaved
    );

    columnLanguageSelects.forEach((select) => {
        select.addEventListener("change", markConfigurationUnsaved);
    });

    void loadSourceColumn();
});


async function loadSourceColumn(): Promise<void> {
    try {
        const config = await getTranslationConfig();
        sourceColumnSelect.value = String(config.source_column);
        columnLanguageSelects[0].value = config.column_1_language;
        columnLanguageSelects[1].value = config.column_2_language;
        columnLanguageSelects[2].value = config.column_3_language;
        updateReferenceColumnLabels();
        showConfigurationSaved(config);
    } catch (error) {
        console.error(error);
        configurationStatus.classList.remove("saved");
        configurationStatus.textContent = `Not saved: ${String(error)}`;
        saveSourceColumnButton.textContent =
            "Save reference and language columns";
        output.textContent = `Error: ${String(error)}`;
    }
}


function updateReferenceColumnLabels(): void {
    for (let index = 0; index < sourceColumnSelect.options.length; index += 1) {
        const language = columnLanguageSelects[index].value;
        sourceColumnSelect.options[index].textContent =
            `Column ${index + 1} — ${language}`;
    }
}


function markConfigurationUnsaved(): void {
    updateReferenceColumnLabels();
    configurationStatus.classList.remove("saved");
    configurationStatus.textContent =
        "Unsaved changes — Translate will save and use these selections automatically.";
    saveSourceColumnButton.textContent =
        "Save reference and language columns";
}


function showConfigurationSaved(config: TranslationConfig): void {
    const languages = [
        config.column_1_language,
        config.column_2_language,
        config.column_3_language
    ];
    configurationStatus.classList.add("saved");
    configurationStatus.textContent =
        `Saved: column ${config.source_column} `
        + `(${languages[config.source_column - 1]}) is the reference.`;
    saveSourceColumnButton.textContent = "Saved ✓";
}


function configurationFromControls(): TranslationConfig {
    const sourceColumn = Number(sourceColumnSelect.value);
    const languages = columnLanguageSelects.map(
        (select) => select.value as Language
    );
    if (!Number.isInteger(sourceColumn)
        || sourceColumn < 1
        || sourceColumn > 3) {
        throw new Error("Reference column must be column 1, 2, or 3.");
    }
    if (new Set(languages).size !== 3) {
        throw new Error(
            "Columns 1-3 must use English, French, and German exactly once."
        );
    }

    return {
        source_column: sourceColumn,
        column_1_language: languages[0],
        column_2_language: languages[1],
        column_3_language: languages[2]
    };
}


async function saveSourceColumn(): Promise<void> {
    saveSourceColumnButton.disabled = true;

    try {
        const config = configurationFromControls();
        await saveTranslationConfig(config);
        const languages = [
            config.column_1_language,
            config.column_2_language,
            config.column_3_language
        ];
        showConfigurationSaved(config);

        output.textContent =
            `Reference: column ${config.source_column} `
            + `(${languages[config.source_column - 1]}).\n`
            + `Language columns saved: ${languages.join(" / ")}.`;
    } catch (error) {
        console.error(error);
        configurationStatus.classList.remove("saved");
        configurationStatus.textContent = `Not saved: ${String(error)}`;
        saveSourceColumnButton.textContent =
            "Save reference and language columns";
        output.textContent = `Error: ${String(error)}`;
    } finally {
        saveSourceColumnButton.disabled = false;
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
    knownSourceRevisions: number;
    newSourceRevisions: number;
    diagnosticLines: string[];
};


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


async function translateChangedTableParagraphs(
    forceAllRevisions = false
): Promise<void> {
    if (!Office.context.requirements.isSetSupported("WordApi", "1.6")) {
        output.textContent =
            "This version of Word does not support tracked-change inspection (WordApi 1.6).";
        return;
    }

    translateChangedParagraphsButton.disabled = true;
    retryAllRevisionsButton.disabled = true;
    saveSourceColumnButton.disabled = true;
    const startedAt = performance.now();
    await showProgress(startedAt, "Loading translation configuration...");

    try {
        // The visible controls are authoritative. Persist them automatically
        // so Translate can never silently fall back to an older saved source.
        const translationConfig = configurationFromControls();
        await saveTranslationConfig(translationConfig);
        showConfigurationSaved(translationConfig);
        const sourceColumn = translationConfig.source_column;
        const configuredLanguages: Language[] = [
            translationConfig.column_1_language,
            translationConfig.column_2_language,
            translationConfig.column_3_language
        ];
        const sourceCellIndex = sourceColumn - 1;
        await showProgress(
            startedAt,
            "Using the selected reference configuration...",
            [
                `Reference: column ${sourceColumn} `
                + `(${configuredLanguages[sourceCellIndex]})`,
                `Targets: ${configuredLanguages
                    .filter((_, index) => index !== sourceCellIndex)
                    .join(" and ")}`
            ]
        );
        await showProgress(
            startedAt,
            "Checking the translation service..."
        );
        await assertTranslationApiAvailable();
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

            const baselineCreated = forceAllRevisions
                || baselineSetting.isNullObject;
            const savedRevisionCounts = baselineCreated
                ? {}
                : parseRevisionBaseline(baselineSetting.value);
            const remainingKnownRevisions: RevisionCounts = {
                ...savedRevisionCounts
            };
            const deferredRevisionCounts: RevisionCounts = {};
            const syntheticRevisionCounts: RevisionCounts = {};

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
                baselineDeferred: false,
                knownSourceRevisions: 0,
                newSourceRevisions: 0,
                diagnosticLines: []
            };

            const eligibleRows: Array<{
                row: Word.TableRow;
                languageColumnIndices: readonly [number, number, number];
            }> = [];

            for (const rows of tableRows) {
                for (const row of rows.items) {
                    if (row.cellCount < 3) {
                        result.invalidRows += 1;
                        continue;
                    }

                    eligibleRows.push({
                        row,
                        languageColumnIndices: row.cellCount >= 7
                            ? STANDARD_LANGUAGE_COLUMN_INDICES
                            : COMPACT_LANGUAGE_COLUMN_INDICES
                    });
                }
            }

            result.eligibleRows = eligibleRows.length;
            const inspectedRows: InspectedParagraph[][][] = [];
            const cellTrackedChangeCountRows: number[][] = [];
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

                for (const candidate of rowBatch) {
                    candidate.row.cells.load("items");
                }

                await syncWithProgress(
                    context,
                    startedAt,
                    "Loading language cells...",
                    [batchLabel, `Paragraphs inspected: ${inspectedParagraphCount}`]
                );

                const rowParagraphCollections = rowBatch.map((candidate) => {
                    const collections = candidate.languageColumnIndices.map((column) => {
                        const cell = candidate.row.cells.items[column];
                        const paragraphs = cell.body.paragraphs;
                        const trackedChanges = cell.body.getTrackedChanges();
                        const ooxml = cell.body.getOoxml();
                        paragraphs.load("items");
                        trackedChanges.load(
                            "items/author,items/date,items/text,items/type"
                        );
                        return { paragraphs, trackedChanges, ooxml };
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
                    (collections) => collections.map(({ paragraphs }) =>
                        paragraphs.items.map((paragraph): InspectedParagraph => {
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
                const cellTrackedChangesBatch = rowParagraphCollections.map(
                    (collections) => collections.map(
                        ({ trackedChanges }) => trackedChanges
                    )
                );
                const cellOoxmlBatch = rowParagraphCollections.map(
                    (collections) => collections.map(({ ooxml }) =>
                        ooxml.value
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

                for (let batchRowIndex = 0;
                    batchRowIndex < inspectedBatch.length;
                    batchRowIndex += 1) {
                    const cells = inspectedBatch[batchRowIndex];
                    const cellTrackedChanges =
                        cellTrackedChangesBatch[batchRowIndex];
                    const cellOoxml = cellOoxmlBatch[batchRowIndex];
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

                    // Word for Mac can omit a whole-paragraph insertion from
                    // Paragraph.getTrackedChanges(), even though the same
                    // revision is returned by the containing cell body. Add
                    // only those cell-level revisions not already observed at
                    // paragraph scope, then map them by their complete text.
                    for (let cellIndex = 0;
                        cellIndex < cells.length;
                        cellIndex += 1) {
                        const paragraphFingerprintCounts: RevisionCounts = {};
                        for (const inspected of cells[cellIndex]) {
                            const fingerprints = await Promise.all(
                                inspected.trackedChanges.items.map(
                                    getRevisionFingerprint
                                )
                            );
                            for (const fingerprint of fingerprints) {
                                addFingerprint(
                                    paragraphFingerprintCounts,
                                    fingerprint
                                );
                            }
                        }

                        for (const revision of cellTrackedChanges[cellIndex]
                            .items) {
                            const fingerprint = await getRevisionFingerprint(
                                revision
                            );
                            const paragraphCount =
                                paragraphFingerprintCounts[fingerprint] ?? 0;
                            if (paragraphCount > 0) {
                                paragraphFingerprintCounts[fingerprint] =
                                    paragraphCount - 1;
                                continue;
                            }

                            const knownCount =
                                remainingKnownRevisions[fingerprint] ?? 0;
                            if (knownCount > 0) {
                                remainingKnownRevisions[fingerprint] =
                                    knownCount - 1;
                                continue;
                            }

                            const revisionText = normalizeWhitespace(
                                revision.text
                            );
                            if (!revisionText) {
                                continue;
                            }
                            const candidateText = revision.type === "Deleted"
                                ? (paragraph: InspectedParagraph) =>
                                    paragraph.originalText.value
                                : (paragraph: InspectedParagraph) =>
                                    paragraph.currentText.value;
                            const matchingParagraphs = cells[cellIndex].filter(
                                (paragraph) => normalizeWhitespace(
                                    candidateText(paragraph)
                                ) === revisionText
                            );
                            if (matchingParagraphs.length !== 1) {
                                continue;
                            }

                            matchingParagraphs[0].newRevisions.push({
                                type: revision.type,
                                text: revision.text
                            });
                            matchingParagraphs[0]
                                .newRevisionFingerprints.push(fingerprint);
                        }

                        for (const insertion of
                            extractOoxmlParagraphInsertions(
                                cellOoxml[cellIndex]
                            )) {
                            addFingerprint(
                                syntheticRevisionCounts,
                                insertion.fingerprint
                            );
                            const knownCount = remainingKnownRevisions[
                                insertion.fingerprint
                            ] ?? 0;
                            if (knownCount > 0) {
                                remainingKnownRevisions[
                                    insertion.fingerprint
                                ] = knownCount - 1;
                                continue;
                            }

                            const insertionText = normalizeWhitespace(
                                insertion.text
                            );
                            const matchingParagraphs = cells[cellIndex].filter(
                                (paragraph) => normalizeWhitespace(
                                    paragraph.currentText.value
                                ) === insertionText
                            );
                            if (matchingParagraphs.length !== 1) {
                                continue;
                            }
                            const matchingParagraph = matchingParagraphs[0];
                            const alreadyObserved =
                                matchingParagraph.newRevisions.some(
                                    (revision) => revision.type === "Added"
                                        && normalizeWhitespace(revision.text)
                                            === insertionText
                                );
                            if (alreadyObserved) {
                                continue;
                            }

                            matchingParagraph.newRevisions.push({
                                type: "Added",
                                text: insertion.text
                            });
                            matchingParagraph.newRevisionFingerprints.push(
                                insertion.fingerprint
                            );
                        }
                    }
                }

                inspectedRows.push(...inspectedBatch);
                cellTrackedChangeCountRows.push(
                    ...cellTrackedChangesBatch.map((collections) =>
                        collections.map((collection) =>
                            collection.items.length
                        )
                    )
                );
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
            let consecutiveNetworkFailures = 0;

            for (let inspectedRowIndex = 0;
                inspectedRowIndex < inspectedRows.length;
                inspectedRowIndex += 1) {
                const cells = inspectedRows[inspectedRowIndex];
                const currentParagraphs = cells.map((cell) =>
                    cell.filter((paragraph) =>
                        paragraph.currentPosition !== undefined
                    ).sort((left, right) =>
                        left.currentPosition! - right.currentPosition!
                    )
                );
                const sourceParagraphs = currentParagraphs[sourceCellIndex];
                const changedSourceParagraphs: ChangedSourceParagraphInput[] = [];
                const sourceTrackedRevisionCount =
                    cellTrackedChangeCountRows[inspectedRowIndex][
                        sourceCellIndex
                    ];
                const sourceNewRevisionCount = cells[sourceCellIndex]
                    .reduce(
                        (total, paragraph) => total
                            + paragraph.newRevisions.length,
                        0
                    );
                result.knownSourceRevisions += Math.max(
                    0,
                    sourceTrackedRevisionCount - sourceNewRevisionCount
                );
                result.newSourceRevisions += sourceNewRevisionCount;

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
                        && (!hasSubstantiveRevision(source.newRevisions)
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

                result.diagnosticLines.push(
                    `Row ${inspectedRowIndex + 1}: paragraphs `
                    + currentParagraphs.map((paragraphs) =>
                        paragraphs.length
                    ).join("/")
                    + `; source tracked ${sourceTrackedRevisionCount}`
                    + `; new after baseline ${sourceNewRevisionCount}`
                    + `; eligible changes ${changedSourceParagraphs.length}.`
                );

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
                        ],
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
                        source_language: configuredLanguages[sourceCellIndex],
                        source_cell: sourceParagraphs.map((paragraph) => ({
                            index: paragraph.currentPosition!,
                            text: paragraph.currentText.value
                        })),
                        changed_source_paragraphs: changedSourceParagraphs,
                        targets
                    }, requestId);
                    consecutiveNetworkFailures = 0;
                } catch (error) {
                    const message = String(error);
                    result.failedCellPlans += 1;
                    result.translationErrors.push(message);
                    if (isNetworkLoadFailure(error)) {
                        consecutiveNetworkFailures += 1;
                    } else {
                        consecutiveNetworkFailures = 0;
                    }
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
                    if (consecutiveNetworkFailures
                        >= MAX_CONSECUTIVE_NETWORK_FAILURES) {
                        throw new Error(
                            "Translation service became unreachable during "
                            + "processing. Stopped after three consecutive "
                            + "network failures; no document changes or "
                            + "revision baseline were committed. Restart the "
                            + "services and use Retry all tracked changes."
                        );
                    }
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
                    if (translatedCell.language
                        !== configuredLanguages[targetCellIndex]) {
                        throw new Error(
                            `Backend returned ${translatedCell.language} for `
                            + `column ${translatedCell.column}, configured as `
                            + `${configuredLanguages[targetCellIndex]}.`
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

            const writeResult = await applyTrackedMutations(
                context,
                wordDocument,
                mutations,
                (stage, details = []) => syncWithProgress(
                    context,
                    startedAt,
                    stage,
                    details
                )
            );
            result.updatedTargets = writeResult.updatedTargets;
            result.minimallyTrackedTargets =
                writeResult.minimallyTrackedTargets;
            result.fullParagraphFallbacks =
                writeResult.fullParagraphFallbacks;

            result.baselineDeferred = Object.keys(
                deferredRevisionCounts
            ).length > 0;

            const currentRevisionCounts =
                await getCurrentRevisionCounts(context);
            for (const [fingerprint, count] of Object.entries(
                syntheticRevisionCounts
            )) {
                currentRevisionCounts[fingerprint] = count;
            }
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
            `New reference-language paragraphs: ${result.sourceParagraphs}`,
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
            `Source revisions already in baseline: ${result.knownSourceRevisions}`,
            `Source revisions considered in this run: ${result.newSourceRevisions}`,
            `Fully deleted source paragraphs skipped: ${result.structuralSources}`,
            result.baselineDeferred
                ? "Completed revisions recorded; failed/skipped revisions retained for retry."
                : "Revision baseline advanced through this run.",
            result.baselineCreated
                ? forceAllRevisions
                    ? "Forced retry: saved baseline was ignored for this run."
                    : "No prior baseline: all existing revisions were treated as new."
                : "Only revisions newer than the saved baseline were treated as sources.",
            ...result.diagnosticLines.slice(0, 10),
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
        retryAllRevisionsButton.disabled = false;
        translateChangedParagraphsButton.disabled = false;
        saveSourceColumnButton.disabled = false;
    }
}
