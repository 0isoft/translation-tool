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
};


type InspectedParagraph = {
    paragraph: Word.Paragraph;
    trackedChanges: Word.TrackedChangeCollection;
    currentText: OfficeExtension.ClientResult<string>;
};


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

            wordDocument.load("changeTrackingMode");
            tables.load("items");
            await context.sync();

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
                missingTargets: 0
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

                        trackedChanges.load("items");

                        return {
                            paragraph,
                            trackedChanges,
                            currentText
                        };
                    })
                );

                inspectedRows.push(inspectedCells);
            }

            await context.sync();

            const replacements: Array<{
                paragraph: Word.Paragraph;
                text: string;
            }> = [];

            for (const cells of inspectedRows) {
                const paragraphCount = Math.max(
                    ...cells.map((cell) => cell.length)
                );

                for (let paragraphIndex = 0;
                    paragraphIndex < paragraphCount;
                    paragraphIndex += 1) {
                    const changedCellIndexes = cells
                        .map((cell, cellIndex) => ({
                            cellIndex,
                            hasChanges: Boolean(
                                cell[paragraphIndex]
                                && cell[paragraphIndex]
                                    .trackedChanges.items.length > 0
                            )
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

                        const target = cells[targetCellIndex][paragraphIndex];

                        if (!target) {
                            result.missingTargets += 1;
                            continue;
                        }

                        const currentText = target.currentText.value;

                        if (currentText.length === 0) {
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

            if (replacements.length === 0) {
                return result;
            }

            const originalTrackingMode = wordDocument.changeTrackingMode;
            const mustRestoreTrackingMode =
                originalTrackingMode === Word.ChangeTrackingMode.off;

            if (mustRestoreTrackingMode) {
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
                if (mustRestoreTrackingMode) {
                    wordDocument.changeTrackingMode = originalTrackingMode;
                    await context.sync();
                }
            }

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
            `Missing sibling paragraph positions skipped: ${result.missingTargets}`
        ].join("\n");
    } catch (error) {
        console.error(error);
        output.textContent = `Error: ${String(error)}`;
    } finally {
        uppercaseChangedParagraphsButton.disabled = false;
    }
}
