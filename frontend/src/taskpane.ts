const output =
    document.getElementById("output") as HTMLPreElement;

const readButton =
    document.getElementById("read-selection") as HTMLButtonElement;

const analyzeButton =
    document.getElementById("analyze-selection") as HTMLButtonElement;

const uppercaseChangedParagraphsButton =
    document.getElementById(
        "uppercase-changed-paragraphs"
    ) as HTMLButtonElement;

const markRevisionsSeenButton =
    document.getElementById("mark-revisions-seen") as HTMLButtonElement;

const REVISION_BASELINE_KEY =
    "translationTool.revisionBaseline.v1";


Office.onReady((info) => {
    if (info.host !== Office.HostType.Word) {
        output.textContent = "This add-in must run inside Word.";
        return;
    }

    output.textContent = "Connected to Word.";

    readButton.addEventListener(
        "click",
        readSelection
    );

    analyzeButton.addEventListener(
        "click",
        analyzeSelection
    );

    uppercaseChangedParagraphsButton.addEventListener(
        "click",
        uppercaseCorrespondingTableParagraphs
    );

    markRevisionsSeenButton.addEventListener(
        "click",
        markCurrentRevisionsAsSeen
    );
});


async function getSelectedText(): Promise<string> {
    return Word.run(async (context) => {
        const selection =
            context.document.getSelection();

        selection.load("text");

        await context.sync();

        return selection.text;
    });
}


async function readSelection(): Promise<void> {
    try {
        const text = await getSelectedText();

        output.textContent =
            `Selected text:\n\n${text}`;
    } catch (error) {
        console.error(error);

        output.textContent =
            `Error: ${String(error)}`;
    }
}


async function analyzeSelection(): Promise<void> {
    try {
        const text = await getSelectedText();

        const response = await fetch(
            "/api/analyze",
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    text
                })
            }
        );

        if (!response.ok) {
            throw new Error(
                `Backend returned ${response.status}`
            );
        }

        const result = await response.json();

        output.textContent =
            `Original:\n${result.original}\n\n`
            + `Suggestion:\n${result.suggestion}`;
    } catch (error) {
        console.error(error);

        output.textContent =
            `Error: ${String(error)}`;
    }
}


type UppercaseResult = {
    tables: number;
    eligibleRows: number;
    sourceParagraphs: number;
    updated: number;
    emptyTargets: number;
    alreadyUppercaseTargets: number;
    ambiguousParagraphs: number;
    invalidRows: number;
    unequalParagraphCountRows: number;
    missingTargets: number;
    structuralSources: number;
    baselineCreated: boolean;
};


type InspectedParagraph = {
    paragraph: Word.Paragraph;
    trackedChanges: Word.TrackedChangeCollection;
    currentText: OfficeExtension.ClientResult<string>;
    originalText: OfficeExtension.ClientResult<string>;
    originalPosition?: number;
    newRevisionCount: number;
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
    uppercaseChangedParagraphsButton.disabled = true;
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
        uppercaseChangedParagraphsButton.disabled = false;
    }
}


async function uppercaseCorrespondingTableParagraphs(): Promise<void> {
    if (!Office.context.requirements.isSetSupported("WordApi", "1.6")) {
        output.textContent =
            "This version of Word does not support tracked-change inspection (WordApi 1.6).";
        return;
    }

    uppercaseChangedParagraphsButton.disabled = true;
    output.textContent = "Inspecting tracked changes inside tables...";

    try {
        const result = await Word.run(async (context): Promise<UppercaseResult> => {
            const wordDocument = context.document;
            const tables = wordDocument.body.tables;
            const baselineSetting = wordDocument.settings
                .getItemOrNullObject(REVISION_BASELINE_KEY);

            wordDocument.load("changeTrackingMode");
            tables.load("items");
            baselineSetting.load("value");
            await context.sync();

            const baselineCreated = baselineSetting.isNullObject;
            const savedRevisionCounts = baselineCreated
                ? {}
                : parseRevisionBaseline(baselineSetting.value);
            const remainingKnownRevisions: RevisionCounts = {
                ...savedRevisionCounts
            };

            const tableRows = tables.items.map((table) => {
                table.rows.load("items/cellCount");
                return table.rows;
            });

            await context.sync();

            const result: UppercaseResult = {
                tables: tables.items.length,
                eligibleRows: 0,
                sourceParagraphs: 0,
                updated: 0,
                emptyTargets: 0,
                alreadyUppercaseTargets: 0,
                ambiguousParagraphs: 0,
                invalidRows: 0,
                unequalParagraphCountRows: 0,
                missingTargets: 0,
                structuralSources: 0,
                baselineCreated
            };

            const eligibleRows: Word.TableRow[] = [];

            for (const rows of tableRows) {
                for (const row of rows.items) {
                    if (row.cellCount !== 3) {
                        result.invalidRows += 1;
                        continue;
                    }

                    row.cells.load("items");
                    eligibleRows.push(row);
                }
            }

            await context.sync();

            const rowParagraphCollections = eligibleRows.map((row) => {
                const collections = row.cells.items.map((cell) => {
                    const paragraphs = cell.body.paragraphs;
                    paragraphs.load("items");
                    return paragraphs;
                });

                return {
                    row,
                    collections
                };
            });

            await context.sync();

            const inspectedRows: InspectedParagraph[][][] = [];

            for (const row of rowParagraphCollections) {
                const paragraphCounts = row.collections.map(
                    (collection) => collection.items.length
                );

                if (!paragraphCounts.every(
                    (count) => count === paragraphCounts[0]
                )) {
                    result.unequalParagraphCountRows += 1;
                }

                result.eligibleRows += 1;

                const inspectedCells = row.collections.map((collection) =>
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
                            newRevisionCount: 0
                        };
                    })
                );

                inspectedRows.push(inspectedCells);
            }

            await context.sync();

            for (const cells of inspectedRows) {
                for (const cell of cells) {
                    let originalPosition = 0;

                    for (const inspected of cell) {
                        if (hasMeaningfulParagraphText(
                            inspected.originalText.value
                        )) {
                            inspected.originalPosition = originalPosition;
                            originalPosition += 1;
                        }

                        const fingerprints = await Promise.all(
                            inspected.trackedChanges.items.map(
                                getRevisionFingerprint
                            )
                        );

                        for (const fingerprint of fingerprints) {
                            const knownCount =
                                remainingKnownRevisions[fingerprint] ?? 0;

                            if (knownCount > 0) {
                                remainingKnownRevisions[fingerprint] =
                                    knownCount - 1;
                            } else {
                                inspected.newRevisionCount += 1;
                            }
                        }
                    }
                }
            }

            const replacements: Array<{
                paragraph: Word.Paragraph;
                text: string;
            }> = [];

            for (const cells of inspectedRows) {
                const paragraphsByOriginalPosition = cells.map((cell) => {
                    const positions = new Map<number, InspectedParagraph>();

                    for (const paragraph of cell) {
                        if (paragraph.originalPosition === undefined) {
                            if (paragraph.newRevisionCount > 0
                                && hasMeaningfulParagraphText(
                                    paragraph.currentText.value
                                )) {
                                result.structuralSources += 1;
                            }

                            continue;
                        }

                        positions.set(paragraph.originalPosition, paragraph);
                    }

                    return positions;
                });
                const originalPositions = new Set<number>();

                for (const positions of paragraphsByOriginalPosition) {
                    for (const position of positions.keys()) {
                        originalPositions.add(position);
                    }
                }

                for (const originalPosition of originalPositions) {
                    const changedCellIndexes = paragraphsByOriginalPosition
                        .map((positions, cellIndex) => ({
                            cellIndex,
                            hasChanges: (() => {
                                const paragraph =
                                    positions.get(originalPosition);

                                return Boolean(
                                    paragraph
                                    && paragraph.newRevisionCount > 0
                                    && hasMeaningfulParagraphText(
                                        paragraph.currentText.value
                                    )
                                );
                            })()
                        }))
                        .filter((cell) => cell.hasChanges)
                        .map((cell) => cell.cellIndex);

                    if (changedCellIndexes.length === 0) {
                        continue;
                    }

                    if (changedCellIndexes.length !== 1) {
                        result.ambiguousParagraphs += 1;
                        continue;
                    }

                    result.sourceParagraphs += 1;
                    const sourceCellIndex = changedCellIndexes[0];

                    for (let targetCellIndex = 0;
                        targetCellIndex < cells.length;
                        targetCellIndex += 1) {
                        if (targetCellIndex === sourceCellIndex) {
                            continue;
                        }

                        const target = paragraphsByOriginalPosition
                            [targetCellIndex]
                            .get(originalPosition);

                        if (!target) {
                            result.missingTargets += 1;
                            continue;
                        }

                        const currentText = target.currentText.value;

                        if (!hasMeaningfulParagraphText(currentText)) {
                            result.emptyTargets += 1;
                            continue;
                        }

                        const uppercaseText = currentText.toLocaleUpperCase();

                        if (uppercaseText === currentText) {
                            result.alreadyUppercaseTargets += 1;
                            continue;
                        }

                        replacements.push({
                            paragraph: target.paragraph,
                            text: uppercaseText
                        });
                    }
                }
            }

            const originalTrackingMode = wordDocument.changeTrackingMode;
            const mustRestoreTrackingMode =
                originalTrackingMode === Word.ChangeTrackingMode.off;

            if (replacements.length > 0 && mustRestoreTrackingMode) {
                wordDocument.changeTrackingMode =
                    Word.ChangeTrackingMode.trackAll;
                await context.sync();
            }

            try {
                // Work backwards so changing an earlier paragraph can't move a
                // later paragraph's position before its replacement is queued.
                for (const replacement of replacements.reverse()) {
                    replacement.paragraph.insertText(
                        replacement.text,
                        Word.InsertLocation.replace
                    );
                }

                await context.sync();
                result.updated = replacements.length;
            } finally {
                if (replacements.length > 0 && mustRestoreTrackingMode) {
                    wordDocument.changeTrackingMode = originalTrackingMode;
                    await context.sync();
                }
            }

            const currentRevisionCounts =
                await getCurrentRevisionCounts(context);
            saveRevisionBaseline(context, currentRevisionCounts);
            await context.sync();

            return result;
        });

        output.textContent = [
            `Tables inspected: ${result.tables}`,
            `Eligible three-cell rows: ${result.eligibleRows}`,
            `Changed source paragraphs: ${result.sourceParagraphs}`,
            `Sibling paragraphs uppercased with Track Changes: ${result.updated}`,
            `Empty sibling paragraphs skipped: ${result.emptyTargets}`,
            `Already-uppercase siblings skipped: ${result.alreadyUppercaseTargets}`,
            `Ambiguous positions skipped: ${result.ambiguousParagraphs}`,
            `Non-three-cell rows skipped: ${result.invalidRows}`,
            `Rows with unequal paragraph counts (processed): ${result.unequalParagraphCountRows}`,
            `Missing sibling paragraph positions skipped: ${result.missingTargets}`,
            `Structural paragraph changes skipped: ${result.structuralSources}`,
            result.baselineCreated
                ? "No prior baseline: all existing revisions were treated as new."
                : "Only revisions newer than the saved baseline were treated as sources."
        ].join("\n");
    } catch (error) {
        console.error(error);
        output.textContent = `Error: ${String(error)}`;
    } finally {
        uppercaseChangedParagraphsButton.disabled = false;
    }
}
