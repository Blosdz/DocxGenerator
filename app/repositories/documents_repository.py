from pathlib import Path
import re
from uuid import UUID

from app.core.config import get_settings
from app.core.database import get_connection


class DocumentsRepository:
    def __init__(self) -> None:
        self.generated_dir = get_settings().generated_dir
        self.schema = get_settings().db_schema

    @property
    def documents_table(self) -> str:
        return f'"{self.schema}".documentos_tesis'

    def get_path(self, filename: str) -> Path:
        if not re.fullmatch(r"[A-Za-z0-9_.-]+\.docx", filename):
            raise FileNotFoundError(filename)

        path = (self.generated_dir / filename).resolve()
        generated_dir = self.generated_dir.resolve()
        if generated_dir not in path.parents or not path.exists():
            raise FileNotFoundError(filename)

        return path

    def register_generated(self, tesis_id: UUID, path: Path, filename: str) -> UUID:
        query = f"""
            INSERT INTO {self.documents_table}
                (tesis_id, subido_por, nombre_archivo, url_archivo_drive,
                 documento_drive_id, version, estado_revision, comentario_revision,
                 ruta_storage, tipo_mime, tamano_bytes)
            VALUES (
                %s,
                NULL,
                %s,
                NULL,
                NULL,
                COALESCE(
                    (
                        SELECT max(version) + 1
                        FROM {self.documents_table}
                        WHERE tesis_id = %s
                    ),
                    1
                ),
                'pendiente',
                'Generado automaticamente por thesis-doc-generator',
                %s,
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                %s
            )
            RETURNING id
        """
        resolved_path = path.resolve()
        with get_connection() as connection:
            row = connection.execute(
                query,
                (
                    tesis_id,
                    filename,
                    tesis_id,
                    str(resolved_path),
                    resolved_path.stat().st_size,
                ),
            ).fetchone()

        return row["id"]
