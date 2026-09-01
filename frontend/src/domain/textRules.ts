import type { RevisionInput, SingleSpanDiff } from "./models";


export function hasMeaningfulParagraphText(value: string): boolean {
    return value.replace(/[\s\u0007]/gu, "").length > 0;
}


export function normalizeWhitespace(value: string): string {
    return value.replace(/[\s\u0007]+/gu, " ").trim();
}


export function hasSubstantiveRevision(revisions: RevisionInput[]): boolean {
    return revisions.some((revision) =>
        (revision.type === "Added" || revision.type === "Deleted")
        && revision.text.replace(/[\s\u0007]+/gu, "").length > 0
    );
}


export function isWholeParagraphAddition(
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


export function getSingleSpanDiff(
    before: string,
    after: string
): SingleSpanDiff {
    let prefixLength = 0;
    const commonLength = Math.min(before.length, after.length);

    while (prefixLength < commonLength
        && before[prefixLength] === after[prefixLength]) {
        prefixLength += 1;
    }

    let suffixLength = 0;
    while (suffixLength < commonLength - prefixLength
        && before[before.length - 1 - suffixLength]
            === after[after.length - 1 - suffixLength]) {
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
