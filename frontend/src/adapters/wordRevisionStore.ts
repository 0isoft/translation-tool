import type {
    RevisionBaseline,
    RevisionCounts
} from "../domain/models";


export const REVISION_BASELINE_KEY =
    "translationTool.revisionBaseline.v1";

const WORDPROCESSINGML_NAMESPACE =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export type OoxmlParagraphInsertion = {
    text: string;
    fingerprint: string;
};


export function extractOoxmlParagraphInsertions(
    ooxml: string
): OoxmlParagraphInsertion[] {
    const xml = new DOMParser().parseFromString(ooxml, "application/xml");
    if (xml.getElementsByTagName("parsererror").length > 0) {
        return [];
    }

    const insertions: OoxmlParagraphInsertion[] = [];
    const paragraphs = Array.from(
        xml.getElementsByTagNameNS(WORDPROCESSINGML_NAMESPACE, "p")
    );
    for (const paragraph of paragraphs) {
        const insertedRuns = Array.from(
            paragraph.getElementsByTagNameNS(
                WORDPROCESSINGML_NAMESPACE,
                "ins"
            )
        );
        const text = insertedRuns.map((insertion) => Array.from(
            insertion.getElementsByTagNameNS(
                WORDPROCESSINGML_NAMESPACE,
                "t"
            )
        ).map((node) => node.textContent ?? "").join(""))
            .join("");
        if (text.replace(/[\s\u0007]+/gu, "").length === 0) {
            continue;
        }

        const identities = insertedRuns.map((insertion) => [
            insertion.getAttributeNS(WORDPROCESSINGML_NAMESPACE, "id") ?? "",
            insertion.getAttributeNS(
                WORDPROCESSINGML_NAMESPACE,
                "author"
            ) ?? ""
        ].join(":"));
        insertions.push({
            text,
            fingerprint: [
                "ooxml-paragraph-insertion-v1",
                identities.join(","),
                text
            ].join("\u0000")
        });
    }

    return insertions;
}


export function parseRevisionBaseline(value: unknown): RevisionCounts {
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


export async function revisionFingerprint(
    change: Word.TrackedChange
): Promise<string> {
    return sha256([
        change.author,
        change.date.toISOString(),
        change.type,
        change.text
    ].join("\u0000"));
}


export function addFingerprint(
    counts: RevisionCounts,
    fingerprint: string
): void {
    counts[fingerprint] = (counts[fingerprint] ?? 0) + 1;
}


export async function currentRevisionCounts(
    context: Word.RequestContext
): Promise<RevisionCounts> {
    const revisions = context.document.body.getTrackedChanges();
    revisions.load("items/author,items/date,items/text,items/type");
    await context.sync();

    const fingerprints = await Promise.all(
        revisions.items.map(revisionFingerprint)
    );
    const counts: RevisionCounts = {};
    for (const fingerprint of fingerprints) {
        addFingerprint(counts, fingerprint);
    }
    return counts;
}


export function saveRevisionBaseline(
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


export async function deleteRevisionBaseline(
    context: Word.RequestContext
): Promise<void> {
    const setting = context.document.settings.getItemOrNullObject(
        REVISION_BASELINE_KEY
    );
    setting.load("value");
    await context.sync();
    if (!setting.isNullObject) {
        setting.delete();
        await context.sync();
    }
}
