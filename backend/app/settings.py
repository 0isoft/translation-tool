import os
from dataclasses import dataclass
from urllib.parse import urlparse


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
    public_base_url: str
    anthropic_api_key: str
    anthropic_model: str
    anthropic_timeout_seconds: float
    anthropic_max_retries: int
    claude_max_attempts: int
    claude_backoff_initial_seconds: float
    claude_backoff_max_seconds: float

    @classmethod
    def from_environment(cls) -> "Settings":
        settings = cls(
            public_base_url=_required("PUBLIC_BASE_URL").rstrip("/"),
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
        )
        settings.validate()
        return settings

    def validate(self) -> None:
        public_url = urlparse(self.public_base_url)
        if public_url.scheme != "https" or not public_url.netloc \
                or public_url.path not in ("", "/") \
                or public_url.params or public_url.query \
                or public_url.fragment:
            raise SettingsError(
                "PUBLIC_BASE_URL must be an HTTPS origin without a path, "
                "query, or fragment."
            )
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
