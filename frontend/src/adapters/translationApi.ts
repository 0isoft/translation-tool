import type {
    TranslateCellChangesRequest,
    TranslateCellChangesResponse,
    TranslationConfig
} from "../domain/models";


const TRANSLATION_REQUEST_TIMEOUT_MS = 180_000;


async function errorDetail(response: Response): Promise<string> {
    try {
        const body = await response.json() as { detail?: unknown };
        return typeof body.detail === "string"
            ? body.detail
            : JSON.stringify(body.detail ?? body);
    } catch {
        return await response.text();
    }
}


export async function getTranslationConfig(): Promise<TranslationConfig> {
    const response = await fetch("/api/config");
    if (!response.ok) {
        throw new Error(
            `Could not load configuration: ${await errorDetail(response)}`
        );
    }
    return await response.json() as TranslationConfig;
}


export async function saveTranslationConfig(
    config: TranslationConfig
): Promise<TranslationConfig> {
    const response = await fetch("/api/config", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(config)
    });
    if (!response.ok) {
        throw new Error(await errorDetail(response));
    }
    return await response.json() as TranslationConfig;
}


export async function requestCellTranslation(
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
                `Translation failed: ${await errorDetail(response)}`
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
