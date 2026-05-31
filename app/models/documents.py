from datetime import datetime
from pathlib import Path
from typing import Any
from uuid import UUID

from pydantic import BaseModel


class GenerateDocumentRequest(BaseModel):
    upload_to_backend: bool = False


class DocumentResponse(BaseModel):
    filename: str
    path: Path
    download_url: str
    generated_at: datetime
    document_id: UUID | None = None
    uploaded: bool = False
    upload: dict[str, Any] | None = None
