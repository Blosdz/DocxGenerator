from uuid import UUID

from psycopg.types.json import Jsonb

from app.core.config import get_settings
from app.core.database import get_connection
from app.models.thesis import SectionCreate, SectionRead, SectionUpdate, ThesisCreate, ThesisRead
from app.repositories.errors import RepositoryNotFoundError


class ThesisRepository:
    def __init__(self) -> None:
        self.schema = get_settings().db_schema

    @property
    def thesis_table(self) -> str:
        return f'"{self.schema}".tesis'

    @property
    def contents_table(self) -> str:
        return f'"{self.schema}".tesis_contents'

    def create(self, thesis: ThesisCreate) -> ThesisRead:
        if thesis.estudiante_id is None:
            raise ValueError("estudiante_id is required when creating a thesis")

        metadata = thesis.metadata.copy()
        metadata.update(
            {
                "title": thesis.title,
                "titulo": thesis.title,
                "author": thesis.author,
                "institution": thesis.institution,
                "year": thesis.year,
                "style": thesis.style.value,
            }
        )

        query = f"""
            INSERT INTO {self.thesis_table}
                (estudiante_id, universidad_id, titulo, descripcion, estado)
            VALUES (%s, %s, %s, %s, 'borrador')
            RETURNING *
        """
        with get_connection() as connection:
            row = connection.execute(
                query,
                (
                    thesis.estudiante_id,
                    thesis.universidad_id,
                    thesis.title,
                    metadata.get("description") or metadata.get("descripcion"),
                ),
            ).fetchone()

        return self._thesis_from_row(row)

    def get(self, tesis_id: UUID) -> ThesisRead:
        query = f"""
            SELECT
                t.*,
                pe.nombres AS estudiante_nombres,
                pe.apellidos AS estudiante_apellidos,
                pe.carrera AS estudiante_carrera
            FROM {self.thesis_table} t
            LEFT JOIN "{self.schema}".perfil_estudiante pe
                ON pe.estudiante_id = t.estudiante_id
            WHERE t.id = %s AND t.eliminado_en IS NULL
        """
        with get_connection() as connection:
            row = connection.execute(query, (tesis_id,)).fetchone()

        if not row:
            raise RepositoryNotFoundError(f"Thesis {tesis_id} was not found")
        return self._thesis_from_row(row)

    def create_section(self, tesis_id: UUID, section: SectionCreate) -> SectionRead:
        self.get(tesis_id)
        data = section.model_dump(mode="json")
        query = f"""
            INSERT INTO {self.contents_table} (tesis_id, data)
            VALUES (%s, %s)
            RETURNING id, tesis_id, data, version, created_at, updated_at, deleted_at
        """
        with get_connection() as connection:
            row = connection.execute(query, (tesis_id, Jsonb(data))).fetchone()

        return self._section_from_row(row)

    def list_sections(self, tesis_id: UUID) -> list[SectionRead]:
        self.get(tesis_id)
        query = f"""
            SELECT id, tesis_id, data, version, created_at, updated_at, deleted_at
            FROM {self.contents_table}
            WHERE tesis_id = %s AND deleted_at IS NULL
            ORDER BY COALESCE((data->>'order')::int, 0), created_at
        """
        with get_connection() as connection:
            rows = connection.execute(query, (tesis_id,)).fetchall()

        return [self._section_from_row(row) for row in rows]

    def replace_sections(self, tesis_id: UUID, sections: list[SectionCreate]) -> list[SectionRead]:
        self.get(tesis_id)
        with get_connection() as connection:
            connection.execute(
                f"""
                    UPDATE {self.contents_table}
                    SET deleted_at = now(), updated_at = now(), version = version + 1
                    WHERE tesis_id = %s AND deleted_at IS NULL
                """,
                (tesis_id,),
            )

            rows = []
            for section in sections:
                row = connection.execute(
                    f"""
                        INSERT INTO {self.contents_table} (tesis_id, data)
                        VALUES (%s, %s)
                        RETURNING id, tesis_id, data, version, created_at, updated_at, deleted_at
                    """,
                    (tesis_id, Jsonb(section.model_dump(mode="json"))),
                ).fetchone()
                rows.append(row)

        return [self._section_from_row(row) for row in rows]

    def update_section(
        self,
        tesis_id: UUID,
        section_id: UUID,
        section: SectionUpdate,
    ) -> SectionRead:
        self.get(tesis_id)
        existing = self._get_section_row(tesis_id, section_id)
        data = existing["data"].copy()
        data.update(section.model_dump(mode="json", exclude_unset=True))

        query = f"""
            UPDATE {self.contents_table}
            SET data = %s, version = version + 1, updated_at = now()
            WHERE id = %s AND tesis_id = %s AND deleted_at IS NULL
            RETURNING id, tesis_id, data, version, created_at, updated_at, deleted_at
        """
        with get_connection() as connection:
            row = connection.execute(query, (Jsonb(data), section_id, tesis_id)).fetchone()

        if not row:
            raise RepositoryNotFoundError(f"Section {section_id} was not found")
        return self._section_from_row(row)

    def delete_section(self, tesis_id: UUID, section_id: UUID) -> dict[str, UUID | bool]:
        self.get(tesis_id)
        query = f"""
            UPDATE {self.contents_table}
            SET deleted_at = now(), version = version + 1, updated_at = now()
            WHERE id = %s AND tesis_id = %s AND deleted_at IS NULL
            RETURNING id
        """
        with get_connection() as connection:
            row = connection.execute(query, (section_id, tesis_id)).fetchone()

        if not row:
            raise RepositoryNotFoundError(f"Section {section_id} was not found")
        return {"id": row["id"], "deleted": True}

    def _thesis_from_row(self, row: dict) -> ThesisRead:
        author_parts = [
            row.get("estudiante_nombres"),
            row.get("estudiante_apellidos"),
        ]
        author = " ".join(part for part in author_parts if part)
        metadata = {
            "title": row["titulo"],
            "titulo": row["titulo"],
            "description": row.get("descripcion"),
            "descripcion": row.get("descripcion"),
            "estado": row.get("estado"),
            "author": author or None,
            "estudiante_id": str(row.get("estudiante_id")) if row.get("estudiante_id") else None,
            "universidad_id": str(row.get("universidad_id")) if row.get("universidad_id") else None,
            "tipo_tesis_id": str(row.get("tipo_tesis_id")) if row.get("tipo_tesis_id") else None,
            "plan_id": str(row.get("plan_id")) if row.get("plan_id") else None,
            "programa_id": str(row.get("programa_id")) if row.get("programa_id") else None,
            "nivel_academico": row.get("nivel_academico"),
            "style": "APA7",
        }
        return ThesisRead(
            id=row["id"],
            data={"metadata": metadata},
            version=1,
            created_at=row["creado_en"],
            updated_at=row["actualizado_en"],
            deleted_at=row["eliminado_en"],
        )

    def _get_section_row(self, tesis_id: UUID, section_id: UUID) -> dict:
        query = f"""
            SELECT id, tesis_id, data, version, created_at, updated_at, deleted_at
            FROM {self.contents_table}
            WHERE id = %s AND tesis_id = %s AND deleted_at IS NULL
        """
        with get_connection() as connection:
            row = connection.execute(query, (section_id, tesis_id)).fetchone()

        if not row:
            raise RepositoryNotFoundError(f"Section {section_id} was not found")
        return row

    def _section_from_row(self, row: dict) -> SectionRead:
        return SectionRead(
            id=row["id"],
            tesis_id=row["tesis_id"],
            version=row["version"],
            created_at=row["created_at"],
            updated_at=row["updated_at"],
            deleted_at=row["deleted_at"],
            **row["data"],
        )
