from pathlib import Path
from typing import Any
from uuid import UUID

import httpx


DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


class BackendUploadError(RuntimeError):
    def __init__(self, message: str, status_code: int | None = None) -> None:
        super().__init__(message)
        self.status_code = status_code


class BackendUploadService:
    def upload_thesis_document(
        self,
        tesis_id: UUID,
        path: Path,
        filename: str,
        backend_url: str,
        authorization: str,
    ) -> dict[str, Any]:
        if not backend_url:
            raise BackendUploadError("Backend URL is not configured")
        if not authorization:
            raise BackendUploadError("Authorization header is required")

        url = f"{backend_url.rstrip('/')}/documentos/tesis/{tesis_id}/archivo"
        try:
            with path.open("rb") as file_obj:
                response = httpx.post(
                    url,
                    headers={"Authorization": authorization},
                    data={"modo": "tesis"},
                    files={"file": (filename, file_obj, DOCX_MIME)},
                    timeout=90,
                )
        except httpx.HTTPError as exc:
            raise BackendUploadError(
                f"Could not upload generated document to backend: {exc}"
            ) from exc

        payload = self._parse_response(response)
        if response.is_error:
            detail = payload.get("message") if isinstance(payload, dict) else payload
            raise BackendUploadError(
                f"Backend upload failed: {detail or response.reason_phrase}",
                status_code=response.status_code,
            )

        if not isinstance(payload, dict):
            raise BackendUploadError("Backend upload returned an unexpected response")

        return payload

    def _parse_response(self, response: httpx.Response) -> Any:
        try:
            return response.json()
        except ValueError:
            return response.text
