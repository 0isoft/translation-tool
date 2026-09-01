import type { Language, TranslationConfig } from "../domain/models";


const STORAGE_KEY = "translationTool.languageConfig.v1";

const DEFAULT_CONFIG: TranslationConfig = {
    source_column: 3,
    column_1_language: "German",
    column_2_language: "French",
    column_3_language: "English"
};

const LANGUAGES = new Set<Language>(["English", "French", "German"]);


function isTranslationConfig(value: unknown): value is TranslationConfig {
    if (!value || typeof value !== "object") {
        return false;
    }

    const candidate = value as Partial<TranslationConfig>;
    const languages = [
        candidate.column_1_language,
        candidate.column_2_language,
        candidate.column_3_language
    ];

    return Number.isInteger(candidate.source_column)
        && (candidate.source_column ?? 0) >= 1
        && (candidate.source_column ?? 0) <= 3
        && languages.every(
            (language) => LANGUAGES.has(language as Language)
        )
        && new Set(languages).size === 3
        && languages[(candidate.source_column ?? 1) - 1] === "English";
}


export async function getTranslationConfig(): Promise<TranslationConfig> {
    const saved = window.localStorage.getItem(STORAGE_KEY);
    if (!saved) {
        return { ...DEFAULT_CONFIG };
    }

    try {
        const parsed: unknown = JSON.parse(saved);
        return isTranslationConfig(parsed)
            ? parsed
            : { ...DEFAULT_CONFIG };
    } catch {
        return { ...DEFAULT_CONFIG };
    }
}


export async function saveTranslationConfig(
    config: TranslationConfig
): Promise<TranslationConfig> {
    if (!isTranslationConfig(config)) {
        throw new Error(
            "Columns 1-3 must use English, French, and German exactly once, "
            + "with English selected as the source column."
        );
    }

    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
    return config;
}
