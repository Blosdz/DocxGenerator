from datetime import date, datetime
from enum import StrEnum
from typing import Annotated
from uuid import UUID

from pydantic import BaseModel, ConfigDict, Field, model_validator


class CitationStyle(StrEnum):
    APA7 = "APA7"
    IEEE = "IEEE"
    VANCOUVER = "VANCOUVER"
    ISO690 = "ISO690"
    MLA = "MLA"


class ReferenceType(StrEnum):
    BOOK = "book"
    ARTICLE = "article"
    WEB = "web"


class Author(BaseModel):
    first_name: str | None = None
    last_name: str


class VerificationStatus(StrEnum):
    FOUND_IN_BODY = "found_in_body"
    FOUND_IN_REFERENCES_SECTION = "found_in_references_section"
    NOT_FOUND = "not_found"
    FETCH_FAILED = "fetch_failed"
    INVALID_CONTENT_TYPE = "invalid_content_type"
    NO_URL = "no_url"
    ERROR = "error"


class SourceVerification(BaseModel):
    quote: str
    status: VerificationStatus
    matched_sentence: str | None = None
    in_reference_section: bool | None = None
    verified_at: datetime
    http_status: int | None = None
    detail: str | None = None


class VerifySourceRequest(BaseModel):
    quote: str = Field(min_length=1, max_length=2000)


class ReferenceBase(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    authors: Annotated[list[Author], Field(min_length=1)]
    year: int | None = Field(default=None, ge=0, le=3000)
    title: str = Field(min_length=1)
    type: ReferenceType
    publisher: str | None = None
    journal: str | None = None
    volume: str | None = None
    issue: str | None = None
    pages: str | None = None
    doi: str | None = None
    url: str | None = None
    accessed_at: date | None = None
    style: CitationStyle = CitationStyle.APA7
    source_verification: SourceVerification | None = None

    @model_validator(mode="before")
    @classmethod
    def normalize_aliases(cls, value):
        if not isinstance(value, dict):
            return value

        normalized = value.copy()
        if "publisher" not in normalized and "editorial" in normalized:
            normalized["publisher"] = normalized["editorial"]
        if "journal" not in normalized and "revista" in normalized:
            normalized["journal"] = normalized["revista"]
        if "volume" not in normalized and "volumen" in normalized:
            normalized["volume"] = normalized["volumen"]
        if "issue" not in normalized and "numero" in normalized:
            normalized["issue"] = normalized["numero"]
        if "pages" not in normalized and "paginas" in normalized:
            normalized["pages"] = normalized["paginas"]
        if "accessed_at" not in normalized:
            if "access_date" in normalized:
                normalized["accessed_at"] = normalized["access_date"]
            elif "fecha_consulta" in normalized:
                normalized["accessed_at"] = normalized["fecha_consulta"]
        if "style" in normalized and normalized["style"] is not None:
            raw = str(normalized["style"]).strip().replace("-", "").replace("_", "").upper()
            _style_map = {"APA7": "APA7", "APA": "APA7", "VANCOUVER": "VANCOUVER", "IEEE": "IEEE", "ISO690": "ISO690", "ISO": "ISO690", "MLA": "MLA"}
            normalized["style"] = _style_map.get(raw, normalized["style"])
        return normalized


class ReferenceCreate(ReferenceBase):
    pass


class ReferenceUpdate(BaseModel):
    authors: list[Author] | None = None
    year: int | None = Field(default=None, ge=0, le=3000)
    title: str | None = Field(default=None, min_length=1)
    type: ReferenceType | None = None
    publisher: str | None = None
    journal: str | None = None
    volume: str | None = None
    issue: str | None = None
    pages: str | None = None
    doi: str | None = None
    url: str | None = None
    accessed_at: date | None = None
    style: CitationStyle | None = None
    source_verification: SourceVerification | None = None


class ReferenceRead(ReferenceBase):
    id: UUID
    tesis_id: UUID
    version: int
    created_at: datetime
    updated_at: datetime
    deleted_at: datetime | None = None


class ReferenceDeleteResponse(BaseModel):
    id: UUID
    deleted: bool
