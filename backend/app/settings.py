import os
from dataclasses import dataclass
from typing import cast

from app.domain.models import Language


class SettingsError(RuntimeError):
    pass


def _required(name: str) -> str:
    try:
        value = os.environ[name]
    except KeyError as error:
        raise SettingsError(
            f"Required environment variable {name} is missing."
        ) from error

    if not value.strip():
        raise SettingsError(
            f"Required environment variable {name} cannot be empty."
        )
    return value


def _integer(name: str) -> int:
    try:
        return int(_required(name))
    except ValueError as error:
        raise SettingsError(f"{name} must be an integer.") from error


def _number(name: str) -> float:
    try:
        return float(_required(name))
    except ValueError as error:
        raise SettingsError(f"{name} must be numeric.") from error


@dataclass(frozen=True)
class Settings:
    anthropic_api_key: str
    anthropic_model: str
    anthropic_timeout_seconds: float
    anthropic_max_retries: int
    claude_max_attempts: int
    claude_backoff_initial_seconds: float
    claude_backoff_max_seconds: float
    source_column: int
    column_1_language: Language
    column_2_language: Language
    column_3_language: Language

    @classmethod
    def from_environment(cls) -> "Settings":
        settings = cls(
            anthropic_api_key=_required("ANTHROPIC_API_KEY"),
            anthropic_model=_required("ANTHROPIC_MODEL"),
            anthropic_timeout_seconds=_number(
                "ANTHROPIC_TIMEOUT_SECONDS"
            ),
            anthropic_max_retries=_integer("ANTHROPIC_MAX_RETRIES"),
            claude_max_attempts=_integer("CLAUDE_MAX_ATTEMPTS"),
            claude_backoff_initial_seconds=_number(
                "CLAUDE_BACKOFF_INITIAL_SECONDS"
            ),
            claude_backoff_max_seconds=_number(
                "CLAUDE_BACKOFF_MAX_SECONDS"
            ),
            source_column=_integer("SOURCE_COLUMN"),
            column_1_language=cast(
                Language,
                _required("COLUMN_1_LANGUAGE"),
            ),
            column_2_language=cast(
                Language,
                _required("COLUMN_2_LANGUAGE"),
            ),
            column_3_language=cast(
                Language,
                _required("COLUMN_3_LANGUAGE"),
            ),
        )
        settings.validate()
        return settings

    def validate(self) -> None:
        if self.anthropic_timeout_seconds <= 0:
            raise SettingsError("ANTHROPIC_TIMEOUT_SECONDS must be positive.")
        if self.anthropic_max_retries < 0:
            raise SettingsError("ANTHROPIC_MAX_RETRIES cannot be negative.")
        if self.claude_max_attempts < 1:
            raise SettingsError("CLAUDE_MAX_ATTEMPTS must be at least 1.")
        if self.claude_backoff_initial_seconds < 0 \
                or self.claude_backoff_max_seconds < 0:
            raise SettingsError("Claude backoff delays cannot be negative.")
        if self.claude_backoff_max_seconds \
                < self.claude_backoff_initial_seconds:
            raise SettingsError(
                "CLAUDE_BACKOFF_MAX_SECONDS cannot be smaller than "
                "CLAUDE_BACKOFF_INITIAL_SECONDS."
            )
        if self.source_column not in (1, 2, 3):
            raise SettingsError("SOURCE_COLUMN must be 1, 2, or 3.")

        languages = {
            self.column_1_language,
            self.column_2_language,
            self.column_3_language,
        }
        if languages != {"English", "French", "German"}:
            raise SettingsError(
                "Columns 1-3 must use English, French, and German exactly once."
            )
        configured_languages = (
            self.column_1_language,
            self.column_2_language,
            self.column_3_language,
        )
        if configured_languages[self.source_column - 1] != "English":
            raise SettingsError(
                "The configured source column must be English."
            )
