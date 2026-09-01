import asyncio
import logging
import os
import random

from anthropic import AsyncAnthropic
from fastapi import HTTPException

from app.ports.claude import ResponseValidator


logger = logging.getLogger("uvicorn.error")


def get_anthropic_client() -> AsyncAnthropic:
    api_key = os.getenv("ANTHROPIC_API_KEY")

    if not api_key:
        raise HTTPException(
            status_code=503,
            detail="ANTHROPIC_API_KEY is not configured",
        )

    try:
        timeout_seconds = float(os.getenv("ANTHROPIC_TIMEOUT_SECONDS", "180"))
        max_retries = int(os.getenv("ANTHROPIC_MAX_RETRIES", "1"))
    except ValueError as error:
        raise HTTPException(
            status_code=503,
            detail=(
                "ANTHROPIC_TIMEOUT_SECONDS and ANTHROPIC_MAX_RETRIES must "
                "be numeric."
            ),
        ) from error

    return AsyncAnthropic(
        api_key=api_key,
        timeout=timeout_seconds,
        max_retries=max_retries,
    )


def get_anthropic_model() -> str:
    model = os.getenv("ANTHROPIC_MODEL")

    if not model:
        raise HTTPException(
            status_code=503,
            detail="ANTHROPIC_MODEL is not configured",
        )

    return model


def get_claude_retry_config() -> tuple[int, float, float]:
    try:
        max_attempts = int(os.getenv("CLAUDE_MAX_ATTEMPTS", "3"))
        initial_delay = float(
            os.getenv("CLAUDE_BACKOFF_INITIAL_SECONDS", "0.75")
        )
        maximum_delay = float(
            os.getenv("CLAUDE_BACKOFF_MAX_SECONDS", "8")
        )
    except ValueError as error:
        raise HTTPException(
            status_code=503,
            detail="Claude retry settings must be numeric.",
        ) from error

    if max_attempts < 1 or initial_delay < 0 or maximum_delay < 0:
        raise HTTPException(
            status_code=503,
            detail=(
                "CLAUDE_MAX_ATTEMPTS must be at least 1 and backoff delays "
                "cannot be negative."
            ),
        )

    return max_attempts, initial_delay, maximum_delay


async def parse_claude_with_backoff(
    *,
    request_id: str,
    purpose: str,
    response_validator=None,
    **parse_arguments,
):
    max_attempts, initial_delay, maximum_delay = get_claude_retry_config()
    client = get_anthropic_client()

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

    @property
    def model(self) -> str:
        return get_anthropic_model()

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
            response_validator=response_validator,
            **parse_arguments,
        )

