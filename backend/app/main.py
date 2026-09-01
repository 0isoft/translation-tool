import uuid
from fastapi import FastAPI, Header, HTTPException
from fastapi.middleware.cors import CORSMiddleware

from app.application.configuration import (
    ConfigurationError,
    InMemoryTranslationConfiguration,
)
from app.application.translation_service import (
    ApplicationRequestError,
    propagate_cell_changes,
)
from app.adapters.anthropic import (
    AnthropicClaudePlanner,
)
from app.domain.models import (
    SourceColumnConfig,
    TranslateCellChangesRequest,
    TranslateCellChangesResponse,
    TranslationConfig,
)


app = FastAPI()


app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


configuration = InMemoryTranslationConfiguration.from_environment()
claude_planner = AnthropicClaudePlanner()




@app.get("/health")
def health():
    return {"status": "ok"}


def current_config() -> TranslationConfig:
    return configuration.get()


@app.get("/config", response_model=TranslationConfig)
def get_config():
    return current_config()


@app.put("/config", response_model=TranslationConfig)
def set_config(config: TranslationConfig):
    try:
        return configuration.replace(config)
    except ConfigurationError as error:
        raise HTTPException(status_code=422, detail=str(error)) from error


@app.put("/config/source-column", response_model=TranslationConfig)
def set_source_column(config: SourceColumnConfig):
    try:
        return configuration.move_source(config.source_column)
    except ConfigurationError as error:
        raise HTTPException(status_code=422, detail=str(error)) from error


@app.post(
    "/translate-cell-changes",
    response_model=TranslateCellChangesResponse,
)
async def translate_cell_changes(
    request: TranslateCellChangesRequest,
    x_request_id: str | None = Header(default=None),
):
    try:
        return await propagate_cell_changes(
            request=request,
            request_id=x_request_id or str(uuid.uuid4()),
            configuration=configuration,
            planner=claude_planner,
        )
    except ApplicationRequestError as error:
        raise HTTPException(
            status_code=error.status_code,
            detail=error.detail,
        ) from error
