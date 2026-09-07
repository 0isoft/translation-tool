import uuid
from xml.sax.saxutils import escape

from fastapi import FastAPI, Header, HTTPException, Request
from fastapi.responses import Response

from app.application.translation_service import (
    ApplicationRequestError,
    propagate_cell_changes,
)
from app.adapters.anthropic import (
    AnthropicClaudePlanner,
)
from app.domain.models import (
    TranslateCellChangesRequest,
    TranslateCellChangesResponse,
)
from app.settings import Settings


app = FastAPI()


settings = Settings.from_environment()
claude_planner = AnthropicClaudePlanner(settings)


@app.get("/api/health")
def health():
    return {"status": "ok"}


def request_origin(request: Request) -> str:
    forwarded_proto = request.headers.get("x-forwarded-proto")
    forwarded_host = request.headers.get("x-forwarded-host")
    scheme = (
        forwarded_proto.split(",", 1)[0].strip()
        if forwarded_proto
        else request.url.scheme
    )
    host = (
        forwarded_host.split(",", 1)[0].strip()
        if forwarded_host
        else request.headers.get("host", request.url.netloc)
    )
    if scheme not in {"http", "https"} or not host:
        raise HTTPException(
            status_code=400,
            detail="Cannot determine the public request origin.",
        )
    return escape(f"{scheme}://{host}", {'"': "&quot;"})


@app.get("/manifest.xml", include_in_schema=False)
def manifest(request: Request):
    base_url = request_origin(request)
    content = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xsi:type="TaskPaneApp">
  <Id>39DA7B46-3E25-4FE8-86C0-77B1B75AF0AC</Id>
  <Version>1.0.2.0</Version>
  <ProviderName>Translation Tool</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Translation Assistant"/>
  <Description DefaultValue="Propagate tracked changes across English, French, and German table cells."/>
  <IconUrl DefaultValue="{base_url}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{base_url}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="{base_url}/"/>
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.6">
      <Set Name="WordApi" MinVersion="1.6"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="{base_url}/taskpane/"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
'''
    return Response(
        content=content,
        media_type="application/xml",
        headers={
            "Cache-Control": "no-store",
            "Content-Disposition": (
                'attachment; filename="translation-assistant.xml"'
            ),
        },
    )


@app.post(
    "/api/translate-cell-changes",
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
            planner=claude_planner,
        )
    except ApplicationRequestError as error:
        raise HTTPException(
            status_code=error.status_code,
            detail=error.detail,
        ) from error
