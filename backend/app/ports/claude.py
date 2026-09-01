from typing import Any, Callable, Protocol


ResponseValidator = Callable[[Any], None]


class ClaudePlanningPort(Protocol):
    @property
    def model(self) -> str:
        ...

    async def parse(
        self,
        *,
        request_id: str,
        purpose: str,
        response_validator: ResponseValidator | None = None,
        **parse_arguments: Any,
    ) -> Any:
        ...
