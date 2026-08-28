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
        uppercaseChangedParagraphs
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
    paragraphsWithChanges: number;
    updated: number;
    deletionOnly: number;
    alreadyUppercase: number;
};


async function uppercaseChangedParagraphs(): Promise<void> {
    if (!Office.context.requirements.isSetSupported("WordApi", "1.6")) {
        output.textContent =
            "This version of Word does not support tracked-change inspection (WordApi 1.6).";
        return;
    }

    uppercaseChangedParagraphsButton.disabled = true;
    output.textContent = "Inspecting tracked changes...";

    try {
        const result = await Word.run(async (context): Promise<UppercaseResult> => {
            const wordDocument = context.document;
            const paragraphs = wordDocument.body.paragraphs;

            wordDocument.load("changeTrackingMode");
            paragraphs.load("items");
            await context.sync();

            const inspectedParagraphs = paragraphs.items.map((paragraph) => {
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
            });

            await context.sync();

            const result: UppercaseResult = {
                paragraphsWithChanges: 0,
                updated: 0,
                deletionOnly: 0,
                alreadyUppercase: 0
            };

            const replacements: Array<{
                paragraph: Word.Paragraph;
                text: string;
            }> = [];

            for (const inspected of inspectedParagraphs) {
                if (inspected.trackedChanges.items.length === 0) {
                    continue;
                }

                result.paragraphsWithChanges += 1;

                const currentText = inspected.currentText.value;

                if (currentText.length === 0) {
                    result.deletionOnly += 1;
                    continue;
                }

                const uppercaseText = currentText.toLocaleUpperCase();

                if (uppercaseText === currentText) {
                    result.alreadyUppercase += 1;
                    continue;
                }

                replacements.push({
                    paragraph: inspected.paragraph,
                    text: uppercaseText
                });
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
            `Paragraphs containing tracked changes: ${result.paragraphsWithChanges}`,
            `Uppercased with Track Changes: ${result.updated}`,
            `Skipped because only deleted text remains: ${result.deletionOnly}`,
            `Skipped because already uppercase: ${result.alreadyUppercase}`
        ].join("\n");
    } catch (error) {
        console.error(error);
        output.textContent = `Error: ${String(error)}`;
    } finally {
        uppercaseChangedParagraphsButton.disabled = false;
    }
}
