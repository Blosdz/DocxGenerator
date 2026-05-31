import psycopg
from fastapi import FastAPI, Request, status
from fastapi.responses import JSONResponse

from app.api.routes import documents, references, thesis
from app.core.database import DatabaseConfigurationError
from app.repositories.errors import RepositoryNotFoundError
from app.services.backend_upload_service import BackendUploadError
from app.services.citation_service import UnsupportedCitationStyleError


app = FastAPI(
    title="Thesis DOCX Generator API",
    version="0.1.0",
    description="API para gestionar tesis, referencias y generar documentos DOCX.",
)

app.include_router(thesis.router)
app.include_router(references.router)
app.include_router(documents.router)


@app.get("/health", tags=["health"])
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.exception_handler(RepositoryNotFoundError)
async def repository_not_found_handler(
    request: Request,
    exc: RepositoryNotFoundError,
) -> JSONResponse:
    return JSONResponse(status_code=status.HTTP_404_NOT_FOUND, content={"detail": str(exc)})


@app.exception_handler(DatabaseConfigurationError)
async def database_configuration_handler(
    request: Request,
    exc: DatabaseConfigurationError,
) -> JSONResponse:
    return JSONResponse(
        status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
        content={"detail": str(exc)},
    )


@app.exception_handler(UnsupportedCitationStyleError)
async def unsupported_citation_style_handler(
    request: Request,
    exc: UnsupportedCitationStyleError,
) -> JSONResponse:
    return JSONResponse(
        status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
        content={"detail": str(exc)},
    )


@app.exception_handler(BackendUploadError)
async def backend_upload_error_handler(
    request: Request,
    exc: BackendUploadError,
) -> JSONResponse:
    status_code = (
        exc.status_code
        if exc.status_code and 400 <= exc.status_code < 500
        else status.HTTP_502_BAD_GATEWAY
    )
    return JSONResponse(
        status_code=status_code,
        content={"detail": str(exc)},
    )


@app.exception_handler(psycopg.Error)
async def database_error_handler(request: Request, exc: psycopg.Error) -> JSONResponse:
    return JSONResponse(
        status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
        content={"detail": "Database operation failed"},
    )
