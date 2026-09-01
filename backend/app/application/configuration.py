from app.domain.models import Language, TranslationConfig
from app.settings import Settings


class ConfigurationError(ValueError):
    pass


class InMemoryTranslationConfiguration:
    """Process-local configuration; restarting the backend resets it."""

    def __init__(self, source_column: int, languages: dict[int, Language]):
        self._validate(source_column, languages)
        self._source_column = source_column
        self._languages = languages.copy()

    @classmethod
    def from_settings(
        cls,
        settings: Settings,
    ) -> "InMemoryTranslationConfiguration":
        languages: dict[int, Language] = {
            1: settings.column_1_language,
            2: settings.column_2_language,
            3: settings.column_3_language,
        }
        return cls(settings.source_column, languages)

    @staticmethod
    def _validate(source_column: int, languages: dict[int, str]) -> None:
        if source_column not in (1, 2, 3):
            raise ConfigurationError("SOURCE_COLUMN must be 1, 2, or 3")
        if set(languages.values()) != {"English", "French", "German"}:
            raise ConfigurationError(
                "Columns 1-3 must use English, French, and German exactly once."
            )
        if languages[source_column] != "English":
            raise ConfigurationError(
                "The configured source column must be English."
            )

    @property
    def source_column(self) -> int:
        return self._source_column

    def language_for(self, column: int) -> Language:
        return self._languages[column]

    def get(self) -> TranslationConfig:
        return TranslationConfig(
            source_column=self._source_column,
            column_1_language=self._languages[1],
            column_2_language=self._languages[2],
            column_3_language=self._languages[3],
        )

    def replace(self, config: TranslationConfig) -> TranslationConfig:
        languages = {
            1: config.column_1_language,
            2: config.column_2_language,
            3: config.column_3_language,
        }
        self._validate(config.source_column, languages)
        self._source_column = config.source_column
        self._languages = languages
        return self.get()

    def move_source(self, source_column: int) -> TranslationConfig:
        if source_column not in (1, 2, 3):
            raise ConfigurationError("Source column must be 1, 2, or 3.")
        previous_source = self._source_column
        selected_language = self._languages[source_column]
        self._languages[previous_source] = selected_language
        self._languages[source_column] = "English"
        self._source_column = source_column
        return self.get()
