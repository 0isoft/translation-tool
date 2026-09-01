import os
import unittest
from unittest.mock import patch

from app.settings import Settings, SettingsError


VALID_ENVIRONMENT = {
    "ANTHROPIC_API_KEY": "test-key",
    "ANTHROPIC_MODEL": "test-model",
    "ANTHROPIC_TIMEOUT_SECONDS": "180",
    "ANTHROPIC_MAX_RETRIES": "1",
    "CLAUDE_MAX_ATTEMPTS": "3",
    "CLAUDE_BACKOFF_INITIAL_SECONDS": "0.75",
    "CLAUDE_BACKOFF_MAX_SECONDS": "8",
    "SOURCE_COLUMN": "3",
    "COLUMN_1_LANGUAGE": "German",
    "COLUMN_2_LANGUAGE": "French",
    "COLUMN_3_LANGUAGE": "English",
}


class SettingsTests(unittest.TestCase):
    def test_complete_environment_is_parsed(self):
        with patch.dict(os.environ, VALID_ENVIRONMENT, clear=True):
            settings = Settings.from_environment()

        self.assertEqual(settings.anthropic_timeout_seconds, 180)
        self.assertEqual(settings.claude_max_attempts, 3)
        self.assertEqual(settings.source_column, 3)

    def test_missing_variable_fails_with_its_name(self):
        incomplete = VALID_ENVIRONMENT.copy()
        del incomplete["CLAUDE_BACKOFF_MAX_SECONDS"]

        with patch.dict(os.environ, incomplete, clear=True):
            with self.assertRaisesRegex(
                SettingsError,
                "CLAUDE_BACKOFF_MAX_SECONDS",
            ):
                Settings.from_environment()

    def test_invalid_language_mapping_fails_at_startup(self):
        invalid = {
            **VALID_ENVIRONMENT,
            "COLUMN_1_LANGUAGE": "French",
            "COLUMN_2_LANGUAGE": "French",
        }

        with patch.dict(os.environ, invalid, clear=True):
            with self.assertRaises(SettingsError):
                Settings.from_environment()


if __name__ == "__main__":
    unittest.main()
