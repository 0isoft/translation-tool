import type { RevisionInput } from "../domain/models";
import { getSingleSpanDiff } from "../domain/textRules";


export type InspectedParagraph = {
    paragraph: Word.Paragraph;
    trackedChanges: Word.TrackedChangeCollection;
    currentText: OfficeExtension.ClientResult<string>;
    originalText: OfficeExtension.ClientResult<string>;
    originalPosition?: number;
    currentPosition?: number;
    newRevisions: RevisionInput[];
    newRevisionFingerprints: string[];
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

export type PlannedMutation = ReplaceMutation | InsertMutation;

type SearchCandidate = {
    ranges: Word.RangeCollection;
    expectedText: string;
    insertLocation: Word.InsertLocation;
};

type PreparedReplacement = {
    replacementText: string;
    candidates: SearchCandidate[];
};

export type MutationWriteResult = {
    updatedTargets: number;
    minimallyTrackedTargets: number;
    fullParagraphFallbacks: number;
};

type SyncProgress = (
    stage: string,
    details?: string[]
) => Promise<void>;


export async function applyTrackedMutations(
    context: Word.RequestContext,
    wordDocument: Word.Document,
    mutations: PlannedMutation[],
    syncProgress: SyncProgress
): Promise<MutationWriteResult> {
    const result: MutationWriteResult = {
        updatedTargets: 0,
        minimallyTrackedTargets: 0,
        fullParagraphFallbacks: 0
    };
    const preparedReplacements = new Map<
        Word.Paragraph,
        PreparedReplacement
    >();
    let narrowSearchCount = 0;

    for (const mutation of mutations) {
        if (mutation.operation !== "replace") {
            continue;
        }
        const diff = getSingleSpanDiff(mutation.originalText, mutation.text);
        const candidates: SearchCandidate[] = [];
        if (diff.prefixLength === 0 && diff.suffixLength === 0) {
            continue;
        }

        if (diff.removedText.length > 0 && diff.removedText.length <= 200) {
            const ranges = mutation.paragraph.search(diff.removedText, {
                ignorePunct: false,
                ignoreSpace: false,
                matchCase: true,
                matchWholeWord: false,
                matchWildcards: false
            });
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
                    diff.prefixLength + Math.min(diff.suffixLength, 80)
                );
                const ranges = mutation.paragraph.search(suffixAnchor, {
                    ignorePunct: false,
                    ignoreSpace: false,
                    matchCase: true,
                    matchWholeWord: false,
                    matchWildcards: false
                });
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
                const ranges = mutation.paragraph.search(prefixAnchor, {
                    ignorePunct: false,
                    ignoreSpace: false,
                    matchCase: true,
                    matchWholeWord: false,
                    matchWildcards: false
                });
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
        await syncProgress(
            "Locating minimal changed spans inside target paragraphs...",
            [`Candidate ranges: ${narrowSearchCount}`]
        );
    }

    const originalTrackingMode = wordDocument.changeTrackingMode;
    const mustRestoreTrackingMode =
        originalTrackingMode === Word.ChangeTrackingMode.off;
    if (mutations.length > 0 && mustRestoreTrackingMode) {
        wordDocument.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await syncProgress(
            "Enabling Track Changes for translated replacements..."
        );
    }

    try {
        for (const mutation of mutations.slice().reverse()) {
            if (mutation.operation === "replace") {
                const prepared = preparedReplacements.get(mutation.paragraph);
                const exactCandidate = prepared?.candidates.find(
                    (candidate) => candidate.ranges.items.filter(
                        (range) => range.text === candidate.expectedText
                    ).length === 1
                );
                if (prepared && exactCandidate) {
                    const exactRange = exactCandidate.ranges.items.find(
                        (range) => range.text === exactCandidate.expectedText
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

        await syncProgress(
            "Writing translated paragraphs...",
            [`Target paragraph edits: ${mutations.length}`]
        );
        result.updatedTargets = mutations.length;
    } finally {
        if (mutations.length > 0 && mustRestoreTrackingMode) {
            wordDocument.changeTrackingMode = originalTrackingMode;
            await syncProgress(
                "Restoring the document's Track Changes setting..."
            );
        }
    }

    return result;
}
