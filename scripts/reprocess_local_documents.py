from app.core.database import get_connection
from app.repositories.documents_repository import DocumentsRepository


def main() -> int:
    with get_connection() as connection:
        documents = connection.execute(
            '''
            SELECT id, nombre_archivo, processing_status
            FROM "AT".documentos_tesis
            WHERE ruta_storage IS NOT NULL
              AND lower(coalesce(nombre_archivo, '')) ~ '\\.(docx|docm)$'
            ORDER BY creado_en
            '''
        ).fetchall()

    repository = DocumentsRepository()
    failures = 0
    print(f"Found {len(documents)} local Word document(s)")
    for document in documents:
        document_id = document["id"]
        try:
            result = repository.process_document(document_id)
            sections = len(result.get("sections") or [])
            references = len(result.get("references") or [])
            print(f"OK {document_id} sections={sections} references={references}")
        except Exception as error:
            failures += 1
            print(f"ERROR {document_id} {type(error).__name__}: {error}")

    print(f"Completed: {len(documents) - failures} ok, {failures} failed")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
