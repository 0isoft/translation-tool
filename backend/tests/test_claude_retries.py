import types
import unittest
from dataclasses import replace
from unittest.mock import AsyncMock, patch

from app.adapters.anthropic import parse_claude_with_backoff
from app.main import settings


class FlakyMessages:
    def __init__(self):
        self.calls = 0

    async def parse(self, **_kwargs):
        self.calls += 1
        if self.calls < 3:
            raise RuntimeError("simulated truncated response")
        return "parsed result"


class FlakyClient:
    def __init__(self):
        self.messages = FlakyMessages()


class ClaudeRetryTests(unittest.IsolatedAsyncioTestCase):
    async def test_retries_with_exponential_backoff_then_succeeds(self):
        client = FlakyClient()
        sleep = AsyncMock()

        retry_settings = replace(
            settings,
            claude_max_attempts=3,
            claude_backoff_initial_seconds=0.5,
            claude_backoff_max_seconds=8,
        )

        with patch(
            "app.adapters.anthropic.get_anthropic_client",
            return_value=client,
        ), patch(
            "app.adapters.anthropic.random.uniform",
            return_value=0,
        ), patch(
            "app.adapters.anthropic.asyncio.sleep",
            sleep,
        ):
            result = await parse_claude_with_backoff(
                request_id="retry-test",
                purpose="test translation",
                settings=retry_settings,
                model="fixture-model",
                messages=[],
                output_format=types.SimpleNamespace,
            )

        self.assertEqual(result, "parsed result")
        self.assertEqual(client.messages.calls, 3)
        self.assertEqual(
            [call.args[0] for call in sleep.await_args_list],
            [0.5, 1.0],
        )


if __name__ == "__main__":
    unittest.main()
