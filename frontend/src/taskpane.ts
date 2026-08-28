const output =
    document.getElementById("output") as HTMLPreElement;

const readButton =
    document.getElementById("read-selection") as HTMLButtonElement;

const analyzeButton =
    document.getElementById("analyze-selection") as HTMLButtonElement;


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
