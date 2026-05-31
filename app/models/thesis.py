from datetime import datetime
from typing import Any
from uuid import UUID

from pydantic import BaseModel, Field

from app.models.references import CitationStyle


class ThesisCreate(BaseModel):
    estudiante_id: UUID | None = None
    universidad_id: UUID | None = None
    title: str = Field(min_length=1)
    author: str | None = None
    institution: str | None = None
    year: int | None = Field(default=None, ge=0, le=3000)
    style: CitationStyle = CitationStyle.APA7
    metadata: dict[str, Any] = Field(default_factory=dict)


class ThesisRead(BaseModel):
    id: UUID
    data: dict[str, Any]
    version: int
    created_at: datetime
    updated_at: datetime
    deleted_at: datetime | None = None

    @property
    def thesis_metadata(self) -> dict[str, Any]:
        metadata = self.data.get("metadata")
        if isinstance(metadata, dict):
            return metadata
        return self.data


class SectionBase(BaseModel):
    title: str = Field(min_length=1)
    subtitle: str | None = None
    level: int = Field(ge=1, le=3)
    content: str = ""
    order: int = Field(ge=0)


class SectionCreate(SectionBase):
    pass


class SectionUpdate(BaseModel):
    title: str | None = Field(default=None, min_length=1)
    subtitle: str | None = None
    level: int | None = Field(default=None, ge=1, le=3)
    content: str | None = None
    order: int | None = Field(default=None, ge=0)


class SectionRead(SectionBase):
    id: UUID
    tesis_id: UUID
    version: int
    created_at: datetime
    updated_at: datetime
    deleted_at: datetime | None = None
