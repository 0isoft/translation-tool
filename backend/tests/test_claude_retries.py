import os
import types
import unittest
from unittest.mock import AsyncMock, patch

from app.main import parse_claude_with_backoff


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

        with patch.dict(
            os.environ,
            {
                "CLAUDE_MAX_ATTEMPTS": "3",
                "CLAUDE_BACKOFF_INITIAL_SECONDS": "0.5",
                "CLAUDE_BACKOFF_MAX_SECONDS": "8",
            },
        ), patch(
            "app.main.get_anthropic_client",
            return_value=client,
        ), patch(
            "app.main.random.uniform",
            return_value=0,
        ), patch(
            "app.main.asyncio.sleep",
            sleep,
        ):
            result = await parse_claude_with_backoff(
                request_id="retry-test",
                purpose="test translation",
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
