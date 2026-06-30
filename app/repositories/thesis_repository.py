from uuid import UUID

from app.core.config import get_settings
from app.core.database import get_connection
from app.models.thesis import (
    FromSkeletonRequest,
    SectionCreate,
    SectionRead,
    SectionUpdate,
    SkeletonSectionItem,
    ThesisCreate,
    ThesisRead,
)
from app.repositories.errors import RepositoryNotFoundError


class ThesisRepository:
    def __init__(self) -> None:
        self.schema = get_settings().db_schema

    @property
    def thesis_table(self) -> str:
        return f'"{self.schema}".tesis'

    @property
    def contents_table(self) -> str:
        return f'"{self.schema}".tesis_sections'

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
                dtf.uname AS doc_thesis_format,
                dtf.name AS doc_thesis_format_name,
                dtf.citation_type,
                dtf.skeleton_json,
                dtf.word_settings_json,
                pe.nombres AS estudiante_nombres,
                pe.apellidos AS estudiante_apellidos,
                pe.carrera AS estudiante_carrera
            FROM {self.thesis_table} t
            LEFT JOIN "{self.schema}".doc_thesis_formats dtf
                ON dtf.id = t.doc_thesis_format_id
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
        query = f"""
            INSERT INTO {self.contents_table}
                (tesis_id, parent_id, title, content, level, order_index)
            VALUES (%s, %s, %s, %s, %s, %s)
            RETURNING *
        """
        with get_connection() as connection:
            row = connection.execute(
                query,
                (
                    tesis_id,
                    section.parent_id,
                    section.title,
                    section.content or "",
                    section.level,
                    section.order,
                ),
            ).fetchone()

        return self._section_from_row(row)

    def list_sections(self, tesis_id: UUID) -> list[SectionRead]:
        self.get(tesis_id)
        query = f"""
            SELECT *
            FROM {self.contents_table}
            WHERE tesis_id = %s AND deleted_at IS NULL
            ORDER BY level, order_index, created_at
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
                        INSERT INTO {self.contents_table}
                            (tesis_id, parent_id, title, content, level, order_index)
                        VALUES (%s, %s, %s, %s, %s, %s)
                        RETURNING *
                    """,
                    (
                        tesis_id,
                        section.parent_id,
                        section.title,
                        section.content or "",
                        section.level,
                        section.order,
                    ),
                ).fetchone()
                rows.append(row)

        return [self._section_from_row(row) for row in rows]

    def replace_sections_from_skeleton(
        self,
        tesis_id: UUID,
        skeleton_sections: list[SkeletonSectionItem],
    ) -> list[SectionRead]:
        sections_to_create = [
            SectionCreate(
                title=item.title,
                level=item.level,
                order=item.order,
                content="",
                subsections=[],
            )
            for item in sorted(skeleton_sections, key=lambda s: s.order)
        ]
        return self.replace_sections(tesis_id, sections_to_create)

    def update_section(
        self,
        tesis_id: UUID,
        section_id: UUID,
        section: SectionUpdate,
    ) -> SectionRead:
        self.get(tesis_id)
        updates = section.model_dump(exclude_unset=True)
        if not updates:
            return self._section_from_row(self._get_section_row(tesis_id, section_id))

        set_clauses = []
        values = []

        field_map = {
            "title": "title",
            "content": "content",
            "level": "level",
            "order": "order_index",
            "parent_id": "parent_id",
        }

        for model_field, col in field_map.items():
            if model_field in updates:
                set_clauses.append(f"{col} = %s")
                values.append(updates[model_field])

        if not set_clauses:
            return self._section_from_row(self._get_section_row(tesis_id, section_id))

        values.extend([section_id, tesis_id])
        query = f"""
            UPDATE {self.contents_table}
            SET {', '.join(set_clauses)}, version = version + 1, updated_at = now()
            WHERE id = %s AND tesis_id = %s AND deleted_at IS NULL
            RETURNING *
        """
        with get_connection() as connection:
            row = connection.execute(query, values).fetchone()

        if not row:
            raise RepositoryNotFoundError(f"Section {section_id} was not found")
        return self._section_from_row(row)

    def append_section_text(self, tesis_id: UUID, section_id: UUID, text: str) -> SectionRead:
        self.get(tesis_id)
        existing = self._get_section_row(tesis_id, section_id)
        current_content = str(existing["content"] or "").strip()
        next_text = text.strip()
        new_content = f"{current_content}\n\n{next_text}" if current_content else next_text

        query = f"""
            UPDATE {self.contents_table}
            SET content = %s, version = version + 1, updated_at = now()
            WHERE id = %s AND tesis_id = %s AND deleted_at IS NULL
            RETURNING *
        """
        with get_connection() as connection:
            row = connection.execute(query, (new_content, section_id, tesis_id)).fetchone()

        if not row:
            raise RepositoryNotFoundError(f"Section {section_id} was not found")
        return self._section_from_row(row)

    def update_cover_metadata(self, tesis_id: UUID, storage_path: str) -> None:
        import json
        cover_data = {"cover_docx_storage_path": storage_path}
        query = f"""
            UPDATE {self.thesis_table}
            SET metadata = COALESCE(metadata, '{{}}'::jsonb) || %s::jsonb,
                actualizado_en = now()
            WHERE id = %s AND eliminado_en IS NULL
        """
        with get_connection() as connection:
            connection.execute(query, (json.dumps(cover_data), tesis_id))

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

    def delete_all_sections(self, tesis_id: UUID) -> int:
        query = f"""
            UPDATE {self.contents_table}
            SET deleted_at = now(), version = version + 1, updated_at = now()
            WHERE tesis_id = %s AND deleted_at IS NULL
            RETURNING id
        """
        with get_connection() as connection:
            rows = connection.execute(query, (tesis_id,)).fetchall()
        return len(rows)

    def _thesis_from_row(self, row: dict) -> ThesisRead:
        author_parts = [
            row.get("estudiante_nombres"),
            row.get("estudiante_apellidos"),
        ]
        author = " ".join(part for part in author_parts if part)
        stored_metadata = row.get("metadata") if isinstance(row.get("metadata"), dict) else {}
        format_uname = (
            row.get("doc_thesis_format")
            or stored_metadata.get("doc_thesis_format")
            or stored_metadata.get("citation_style")
            or "apa7"
        )
        metadata = {
            **stored_metadata,
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
            "doc_thesis_format_id": str(row.get("doc_thesis_format_id")) if row.get("doc_thesis_format_id") else None,
            "doc_thesis_format": format_uname,
            "doc_thesis_format_name": row.get("doc_thesis_format_name"),
            "citation_style": format_uname,
            "citation_mode": row.get("citation_type") or stored_metadata.get("citation_mode"),
            "skeleton_json": row.get("skeleton_json") or stored_metadata.get("skeleton_json"),
            "word_settings_json": row.get("word_settings_json") or stored_metadata.get("word_settings_json"),
            "style": str(format_uname).upper(),
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
            SELECT *
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
            parent_id=row["parent_id"],
            title=row["title"],
            content=row.get("content") or "",
            level=row["level"],
            order=row["order_index"],
            version=row["version"],
            created_at=row["created_at"],
            updated_at=row["updated_at"],
            deleted_at=row["deleted_at"],
            subsections=[],
        )
