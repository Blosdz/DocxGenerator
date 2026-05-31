from uuid import UUID

from fastapi import APIRouter, Body, HTTPException, Request
from fastapi.responses import FileResponse

from app.core.config import get_settings
from app.models.documents import DocumentResponse, GenerateDocumentRequest
from app.repositories.documents_repository import DocumentsRepository
from app.services.thesis_builder import ThesisBuilder


router = APIRouter(tags=["documents"])


def get_builder() -> ThesisBuilder:
    return ThesisBuilder()


def get_repository() -> DocumentsRepository:
    return DocumentsRepository()


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


@router.get("/documents/{filename}")
async def download_document(filename: str) -> FileResponse:
    try:
        path = get_repository().get_path(filename)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail="Document was not found") from exc

    return FileResponse(
        path=path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
    )
