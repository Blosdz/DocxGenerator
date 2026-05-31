from uuid import UUID

from app.models.documents import DocumentResponse
from app.models.references import CitationStyle
from app.repositories.documents_repository import DocumentsRepository
from app.repositories.references_repository import ReferencesRepository
from app.repositories.thesis_repository import ThesisRepository
from app.services.backend_upload_service import BackendUploadService
from app.services.docx_service import DocxService


class ThesisBuilder:
    def __init__(
        self,
        thesis_repository: ThesisRepository | None = None,
        references_repository: ReferencesRepository | None = None,
        documents_repository: DocumentsRepository | None = None,
        docx_service: DocxService | None = None,
        backend_upload_service: BackendUploadService | None = None,
    ) -> None:
        self.thesis_repository = thesis_repository or ThesisRepository()
        self.references_repository = references_repository or ReferencesRepository()
        self.documents_repository = documents_repository or DocumentsRepository()
        self.docx_service = docx_service or DocxService()
        self.backend_upload_service = backend_upload_service or BackendUploadService()

    def build_docx(
        self,
        tesis_id: UUID,
        upload_to_backend: bool = False,
        backend_url: str | None = None,
        authorization: str | None = None,
    ) -> DocumentResponse:
        thesis = self.thesis_repository.get(tesis_id)
        sections = self.thesis_repository.list_sections(tesis_id)
        references = self.references_repository.list_by_thesis(tesis_id)
        style = CitationStyle(thesis.thesis_metadata.get("style", CitationStyle.APA7.value))
        response = self.docx_service.generate(thesis, sections, references, style)

        if upload_to_backend:
            upload = self.backend_upload_service.upload_thesis_document(
                tesis_id=tesis_id,
                path=response.path,
                filename=self.docx_service.upload_filename(thesis),
                backend_url=backend_url or "",
                authorization=authorization or "",
            )
            return response.model_copy(update={"uploaded": True, "upload": upload})

        document_id = self.documents_repository.register_generated(
            tesis_id,
            response.path,
            response.filename,
        )
        return response.model_copy(update={"document_id": document_id})
