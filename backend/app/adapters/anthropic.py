import asyncio
import logging
import random

from anthropic import AsyncAnthropic

from app.ports.claude import ResponseValidator
from app.settings import Settings


logger = logging.getLogger("uvicorn.error")


def get_anthropic_client(settings: Settings) -> AsyncAnthropic:
    return AsyncAnthropic(
        api_key=settings.anthropic_api_key,
        timeout=settings.anthropic_timeout_seconds,
        max_retries=settings.anthropic_max_retries,
    )


def get_anthropic_model(settings: Settings) -> str:
    return settings.anthropic_model


def get_claude_retry_config(
    settings: Settings,
) -> tuple[int, float, float]:
    return (
        settings.claude_max_attempts,
        settings.claude_backoff_initial_seconds,
        settings.claude_backoff_max_seconds,
    )


async def parse_claude_with_backoff(
    *,
    request_id: str,
    purpose: str,
    settings: Settings,
    response_validator=None,
    **parse_arguments,
):
    max_attempts, initial_delay, maximum_delay = get_claude_retry_config(
        settings
    )
    client = get_anthropic_client(settings)

    for attempt in range(1, max_attempts + 1):
        try:
            response = await client.messages.parse(**parse_arguments)
            if response_validator is not None:
                response_validator(response)
            return response
        except Exception as error:
            if attempt == max_attempts:
                raise

            exponential_delay = min(
                initial_delay * (2 ** (attempt - 1)),
                maximum_delay,
            )
            jitter = random.uniform(0, exponential_delay * 0.25)
            delay = exponential_delay + jitter
            logger.warning(
                "%s request %s attempt %d/%d failed (%s); retrying in %.2fs",
                purpose,
                request_id,
                attempt,
                max_attempts,
                type(error).__name__,
                delay,
            )
            await asyncio.sleep(delay)

    raise AssertionError("Claude retry loop completed without a result")


class AnthropicClaudePlanner:
    """Claude SDK adapter implementing the application's outbound port."""

    def __init__(self, settings: Settings):
        self.settings = settings

    @property
    def model(self) -> str:
        return get_anthropic_model(self.settings)

    async def parse(
        self,
        *,
        request_id: str,
        purpose: str,
        response_validator: ResponseValidator | None = None,
        **parse_arguments,
    ):
        return await parse_claude_with_backoff(
            request_id=request_id,
            purpose=purpose,
            settings=self.settings,
            response_validator=response_validator,
            **parse_arguments,
        )
