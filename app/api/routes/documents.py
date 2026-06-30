from uuid import UUID

from fastapi import APIRouter, Body, HTTPException, Query, Request
from fastapi.responses import FileResponse

from app.core.config import get_settings
from app.models.documents import (
    DocumentMetadataHarness,
    DocumentPreviewResponse,
    DocumentResponse,
    DocumentRawDataItem,
    ExtractedOutlineItem,
    ExtractedOutlineResponse,
    ExtractedReferenceSummary,
    ExtractedReferencesResponse,
    GenerateDocumentRequest,
    HeadingInsertRequest,
    InlineCitationRequest,
    RawDataUpdateRequest,
    RawDataResponse,
    RawDocumentResponse,
    StructuredReferenceUpdate,
    StructuredSectionUpdate,
    SubtitleInsertRequest,
)
from app.repositories.documents_repository import DocumentsRepository
from app.services.outline_extraction_service import OutlineExtractionService
from app.services.reference_extraction_service import ReferenceExtractionService
from app.services.thesis_builder import ThesisBuilder

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
DOCM_MIME = "application/vnd.ms-word.document.macroEnabled.12"

router = APIRouter(tags=["documents"])


def get_builder() -> ThesisBuilder:
    return ThesisBuilder()


def get_repository() -> DocumentsRepository:
    return DocumentsRepository()


def get_reference_extraction_service() -> ReferenceExtractionService:
    return ReferenceExtractionService()


def get_outline_extraction_service() -> OutlineExtractionService:
    return OutlineExtractionService()


@router.post("/theses/{tesis_id}/documents/docx", response_model=DocumentResponse)
async def generate_docx(
    tesis_id: UUID,
    request: Request,
    payload: GenerateDocumentRequest | None = Body(default=None),
) -> DocumentResponse:
    payload = payload or GenerateDocumentRequest()
    backend_url = request.headers.get("x-backend-base-url") or get_settings().backend_url
    return get_builder().build_docx(
        tesis_id,
        upload_to_backend=payload.upload_to_backend,
        backend_url=backend_url,
        authorization=request.headers.get("authorization"),
    )


@router.post("/documents/{document_id}/process", response_model=DocumentRawDataItem)
async def process_document(document_id: UUID) -> DocumentRawDataItem:
    result = get_repository().process_document(document_id)
    # Sync extracted references into the thesis-level store so the reference
    # manager (Gestor de Referencias) reflects them. Deduped and best-effort:
    # a failure here must not break document processing.
    try:
        get_reference_extraction_service().extract_and_create(document_id)
    except Exception:
        pass
    return DocumentRawDataItem(**result)


@router.get("/documents/{document_id}/raw-data")
async def get_raw_data(document_id: UUID):
    repository = get_repository()
    data = repository.get_raw_data(document_id)
    return data


@router.get("/documents/{document_id}/metadata-harness", response_model=DocumentMetadataHarness)
async def get_metadata_harness(document_id: UUID) -> DocumentMetadataHarness:
    try:
        return DocumentMetadataHarness(**get_repository().get_metadata_harness(document_id))
    except Exception as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc


@router.get("/documents/{document_id}/raw-document", response_model=RawDocumentResponse)
async def get_raw_document(document_id: UUID) -> RawDocumentResponse:
    try:
        return RawDocumentResponse(**get_repository().get_raw_document(document_id))
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc


@router.patch("/documents/{document_id}/raw-data", response_model=RawDataResponse)
async def update_raw_data(
    document_id: UUID,
    payload: RawDataUpdateRequest,
) -> RawDataResponse:
    updated = get_repository().update_raw_data(document_id, payload.raw_data)
    return RawDataResponse(**updated)


@router.post("/documents/{document_id}/raw-data/extract")
async def extract_raw_data(document_id: UUID):
    try:
        repository = get_repository()
        process_document = getattr(repository, "process_document", None)
        if callable(process_document):
            result = process_document(document_id)
            try:
                get_reference_extraction_service().extract_and_create(document_id)
            except Exception:
                pass
            return result
        return repository.extract_raw_data(document_id)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc


@router.get("/documents/{document_id}/sections")
async def list_sections(document_id: UUID):
    return get_repository().list_sections(document_id)


@router.get("/documents/{document_id}/sections/{section_id}")
async def get_section(document_id: UUID, section_id: UUID):
    return get_repository().get_section(document_id, section_id)


@router.patch("/documents/{document_id}/sections/{section_id}")
async def update_section(
    document_id: UUID,
    section_id: UUID,
    payload: StructuredSectionUpdate,
):
    return get_repository().update_section(document_id, section_id, payload)


@router.get("/documents/{document_id}/references")
async def list_references(document_id: UUID):
    return get_repository().list_references(document_id)


@router.get("/documents/{document_id}/references/{reference_id}")
async def get_reference(document_id: UUID, reference_id: UUID):
    return get_repository().get_reference(document_id, reference_id)


@router.patch("/documents/{document_id}/references/{reference_id}")
async def update_reference(
    document_id: UUID,
    reference_id: UUID,
    payload: StructuredReferenceUpdate,
):
    return get_repository().update_reference(document_id, reference_id, payload)


@router.get("/documents/{document_id}/preview", response_model=DocumentPreviewResponse)
async def get_preview(document_id: UUID) -> DocumentPreviewResponse:
    return get_repository().get_preview(document_id)


@router.post("/documents/{document_id}/references/extract", response_model=ExtractedReferencesResponse)
async def extract_references(document_id: UUID) -> ExtractedReferencesResponse:
    try:
        return ExtractedReferencesResponse(
            **get_reference_extraction_service().extract_and_create(document_id)
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc


@router.post("/documents/{document_id}/outline/extract", response_model=ExtractedOutlineResponse)
async def extract_outline(
    document_id: UUID,
    replace: bool = Query(default=False),
) -> ExtractedOutlineResponse:
    try:
        return ExtractedOutlineResponse(
            **get_outline_extraction_service().extract_and_create(document_id, replace=replace)
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc


@router.post("/documents/{document_id}/citations", response_model=RawDocumentResponse)
async def insert_citation(
    document_id: UUID,
    payload: InlineCitationRequest,
) -> RawDocumentResponse:
    try:
        return RawDocumentResponse(
            **get_repository().insert_citation(
                document_id=document_id,
                reference_id=payload.reference_id,
                paragraph_index=payload.paragraph_index,
                char_offset=payload.char_offset,
            )
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc


@router.post("/documents/{document_id}/headings", response_model=RawDocumentResponse)
async def insert_heading(
    document_id: UUID,
    payload: HeadingInsertRequest,
) -> RawDocumentResponse:
    try:
        return RawDocumentResponse(
            **get_repository().insert_heading(
                document_id=document_id,
                text=payload.text,
                paragraph_index=payload.paragraph_index,
                char_offset=payload.char_offset,
                level=payload.level,
                mode=payload.mode,
            )
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc


@router.post("/documents/{document_id}/subtitles", response_model=RawDocumentResponse)
async def insert_subtitle(
    document_id: UUID,
    payload: SubtitleInsertRequest,
) -> RawDocumentResponse:
    try:
        return RawDocumentResponse(
            **get_repository().insert_subtitle(
                document_id=document_id,
                text=payload.text,
                paragraph_index=payload.paragraph_index,
                char_offset=payload.char_offset,
                level=payload.level,
                mode=payload.mode,
            )
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc


@router.get("/documents/{filename}")
async def download_document(filename: str) -> FileResponse:
    try:
        path = get_repository().get_path(filename)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail="Document was not found") from exc

    return FileResponse(
        path=path,
        media_type=DOCM_MIME if filename.lower().endswith(".docm") else DOCX_MIME,
        filename=filename,
    )
