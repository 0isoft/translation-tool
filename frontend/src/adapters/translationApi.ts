import type {
    TranslateCellChangesRequest,
    TranslateCellChangesResponse
} from "../domain/models";


const TRANSLATION_REQUEST_TIMEOUT_MS = 180_000;
const HEALTH_REQUEST_TIMEOUT_MS = 10_000;


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


export async function assertTranslationApiAvailable(): Promise<void> {
    const controller = new AbortController();
    const timeout = window.setTimeout(
        () => controller.abort(),
        HEALTH_REQUEST_TIMEOUT_MS
    );

    try {
        const response = await fetch("/api/health", {
            cache: "no-store",
            signal: controller.signal
        });
        if (!response.ok) {
            throw new Error(
                `Translation service health check returned ${response.status}.`
            );
        }
    } catch (error) {
        throw new Error(
            "The translation service is unreachable. Confirm that the "
            + "Docker services are running, then retry.",
            { cause: error }
        );
    } finally {
        window.clearTimeout(timeout);
    }
}


export function isNetworkLoadFailure(error: unknown): boolean {
    return error instanceof TypeError
        && /load failed|failed to fetch|network/i.test(error.message);
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
