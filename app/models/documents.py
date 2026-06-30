from datetime import datetime
from enum import StrEnum
from pathlib import Path
from typing import Any, Literal
from uuid import UUID

from pydantic import BaseModel, Field


class GenerateDocumentRequest(BaseModel):
    upload_to_backend: bool = False


class DocumentResponse(BaseModel):
    filename: str
    path: Path
    download_url: str
    generated_at: datetime
    format: str = "docx"
    mime_type: str = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    document_id: UUID | None = None
    uploaded: bool = False
    upload: dict[str, Any] | None = None


class RawDataUpdateRequest(BaseModel):
    raw_data: str


class RawDataResponse(BaseModel):
    document_id: UUID
    raw_data: str | None = None


class DocumentParagraph(BaseModel):
    paragraph_index: int
    text: str
    char_count: int
    style: str | None = None


class RawDocumentResponse(BaseModel):
    document_id: UUID
    raw_data: str
    paragraphs: list[DocumentParagraph]


class StructuredSectionBase(BaseModel):
    heading: str = Field(min_length=1)
    level: int = Field(ge=1, le=6)
    content: str = ""
    order_index: int = Field(ge=0)
    parent_section_id: UUID | None = None
    source_paragraphs: list[int] = Field(default_factory=list)
    manual_override: bool = False


class StructuredSectionUpdate(BaseModel):
    heading: str | None = Field(default=None, min_length=1)
    level: int | None = Field(default=None, ge=1, le=6)
    content: str | None = None
    order_index: int | None = Field(default=None, ge=0)
    parent_section_id: UUID | None = None
    source_paragraphs: list[int] | None = None
    manual_override: bool | None = None


class StructuredSectionRead(StructuredSectionBase):
    id: UUID
    document_id: UUID
    version: int
    created_at: datetime
    updated_at: datetime
    deleted_at: datetime | None = None


class StructuredReferenceBase(BaseModel):
    type: str = Field(min_length=1)
    authors: list[dict[str, Any]] = Field(default_factory=list)
    year: int | None = Field(default=None, ge=0, le=3000)
    title: str = Field(min_length=1)
    raw_text: str = Field(default="")
    style: str | None = None
    source: str = Field(default="text")
    confidence: float = Field(default=0.5, ge=0.0, le=1.0)


class StructuredReferenceUpdate(BaseModel):
    type: str | None = Field(default=None, min_length=1)
    authors: list[dict[str, Any]] | None = None
    year: int | None = Field(default=None, ge=0, le=3000)
    title: str | None = Field(default=None, min_length=1)
    raw_text: str | None = None
    style: str | None = None
    source: str | None = None
    confidence: float | None = Field(default=None, ge=0.0, le=1.0)


class StructuredReferenceRead(StructuredReferenceBase):
    id: UUID
    document_id: UUID
    version: int
    created_at: datetime
    updated_at: datetime
    deleted_at: datetime | None = None


class DocumentRawParagraph(BaseModel):
    paragraph_index: int
    text: str
    style: str | None = None
    heading_level: int | None = None
    section_order_index: int | None = None
    is_reference: bool = False


class DocumentRawDataItem(BaseModel):
    document_id: UUID
    title: str | None = None
    raw_data: str | None = None
    raw_data_json: dict[str, Any] | None = None
    processing_status: str | None = None
    processing_error: str | None = None
    sections: list[StructuredSectionRead] = Field(default_factory=list)
    references: list[StructuredReferenceRead] = Field(default_factory=list)
    paragraphs: list[DocumentRawParagraph] = Field(default_factory=list)


class DocumentPreviewBlock(BaseModel):
    kind: Literal["title", "heading", "paragraph", "reference"]
    text: str
    level: int | None = None
    section_id: UUID | None = None
    order_index: int | None = None


class DocumentPreviewResponse(BaseModel):
    document_id: UUID
    title: str | None = None
    preview_html: str
    blocks: list[DocumentPreviewBlock] = Field(default_factory=list)
    sections: list[StructuredSectionRead] = Field(default_factory=list)
    references: list[StructuredReferenceRead] = Field(default_factory=list)


class ExtractedReferenceSummary(BaseModel):
    id: UUID | None = None
    title: str
    year: int | None = None
    type: str
    source: str
    status: str
    reason: str | None = None


class ExtractedReferencesResponse(BaseModel):
    document_id: UUID
    tesis_id: UUID
    extracted_count: int = 0
    created_count: int = 0
    skipped_count: int = 0
    references: list[ExtractedReferenceSummary] = Field(default_factory=list)


class ExtractedOutlineItem(BaseModel):
    id: UUID | None = None
    title: str
    level: int
    order: int
    source: str = "heading"
    status: str
    reason: str | None = None
    subtitles: list[str] = Field(default_factory=list)


class ExtractedOutlineResponse(BaseModel):
    document_id: UUID
    tesis_id: UUID
    extracted_count: int = 0
    created_count: int = 0
    updated_count: int = 0
    skipped_count: int = 0
    sections: list[ExtractedOutlineItem] = Field(default_factory=list)


class InlineCitationRequest(BaseModel):
    reference_id: UUID
    paragraph_index: int = Field(ge=0)
    char_offset: int = Field(ge=0)


class ParagraphEditMode(StrEnum):
    INSERT = "insert"
    REPLACE = "replace"


class HeadingInsertRequest(BaseModel):
    text: str = Field(min_length=1)
    paragraph_index: int = Field(ge=0)
    char_offset: int = Field(default=0, ge=0)
    level: int = Field(default=1, ge=1, le=3)
    mode: ParagraphEditMode = ParagraphEditMode.INSERT


class SubtitleInsertRequest(BaseModel):
    text: str = Field(min_length=1)
    paragraph_index: int = Field(ge=0)
    char_offset: int = Field(default=0, ge=0)
    level: int = Field(default=2, ge=1, le=3)
    mode: ParagraphEditMode = ParagraphEditMode.INSERT


# ── AI Harness models ────────────────────────────────────────────────────────

class HarnessSection(BaseModel):
    heading: str
    level: int
    content: str
    order: int


class HarnessReference(BaseModel):
    title: str
    authors: list[dict[str, Any]]
    year: int | None
    type: str
    raw_text: str


class DocumentMetadataHarness(BaseModel):
    document_id: UUID
    title: str | None
    format: str | None
    word_count: int
    processing_status: str
    sections: list[HarnessSection]
    references: list[HarnessReference]
