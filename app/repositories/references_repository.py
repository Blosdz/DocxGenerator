from uuid import UUID

from psycopg.types.json import Jsonb

from app.core.config import get_settings
from app.core.database import get_connection
from app.models.references import ReferenceCreate, ReferenceDeleteResponse, ReferenceRead, ReferenceUpdate
from app.repositories.errors import RepositoryNotFoundError


class ReferencesRepository:
    def __init__(self) -> None:
        self.schema = get_settings().db_schema

    @property
    def thesis_table(self) -> str:
        return f'"{self.schema}".tesis'

    @property
    def references_table(self) -> str:
        return f'"{self.schema}".tesis_references'

    def create(self, tesis_id: UUID, reference: ReferenceCreate) -> ReferenceRead:
        self._ensure_thesis_exists(tesis_id)
        data = reference.model_dump(mode="json")
        query = f"""
            INSERT INTO {self.references_table} (tesis_id, data)
            VALUES (%s, %s)
            RETURNING id, tesis_id, data, version, created_at, updated_at, deleted_at
        """
        with get_connection() as connection:
            row = connection.execute(query, (tesis_id, Jsonb(data))).fetchone()

        return self._reference_from_row(row)

    def list_by_thesis(self, tesis_id: UUID) -> list[ReferenceRead]:
        self._ensure_thesis_exists(tesis_id)
        query = f"""
            SELECT id, tesis_id, data, version, created_at, updated_at, deleted_at
            FROM {self.references_table}
            WHERE tesis_id = %s AND deleted_at IS NULL
            ORDER BY lower(data->>'title'), created_at
        """
        with get_connection() as connection:
            rows = connection.execute(query, (tesis_id,)).fetchall()

        return [self._reference_from_row(row) for row in rows]

    def get(self, reference_id: UUID) -> ReferenceRead:
        query = f"""
            SELECT id, tesis_id, data, version, created_at, updated_at, deleted_at
            FROM {self.references_table}
            WHERE id = %s AND deleted_at IS NULL
        """
        with get_connection() as connection:
            row = connection.execute(query, (reference_id,)).fetchone()

        if not row:
            raise RepositoryNotFoundError(f"Reference {reference_id} was not found")
        return self._reference_from_row(row)

    def update(self, reference_id: UUID, reference: ReferenceUpdate) -> ReferenceRead:
        existing = self._get_reference_row(reference_id)
        data = existing["data"].copy()
        data.update(reference.model_dump(mode="json", exclude_unset=True))

        query = f"""
            UPDATE {self.references_table}
            SET data = %s, version = version + 1, updated_at = now()
            WHERE id = %s AND deleted_at IS NULL
            RETURNING id, tesis_id, data, version, created_at, updated_at, deleted_at
        """
        with get_connection() as connection:
            row = connection.execute(query, (Jsonb(data), reference_id)).fetchone()

        if not row:
            raise RepositoryNotFoundError(f"Reference {reference_id} was not found")
        return self._reference_from_row(row)

    def delete(self, reference_id: UUID) -> ReferenceDeleteResponse:
        query = f"""
            UPDATE {self.references_table}
            SET deleted_at = now(), version = version + 1
            WHERE id = %s AND deleted_at IS NULL
            RETURNING id
        """
        with get_connection() as connection:
            row = connection.execute(query, (reference_id,)).fetchone()

        if not row:
            raise RepositoryNotFoundError(f"Reference {reference_id} was not found")
        return ReferenceDeleteResponse(id=row["id"], deleted=True)

    def _ensure_thesis_exists(self, tesis_id: UUID) -> None:
        query = f"""
            SELECT id
            FROM {self.thesis_table}
            WHERE id = %s AND eliminado_en IS NULL
        """
        with get_connection() as connection:
            row = connection.execute(query, (tesis_id,)).fetchone()

        if not row:
            raise RepositoryNotFoundError(f"Thesis {tesis_id} was not found")

    def _get_reference_row(self, reference_id: UUID) -> dict:
        query = f"""
            SELECT id, tesis_id, data, version, created_at, updated_at, deleted_at
            FROM {self.references_table}
            WHERE id = %s AND deleted_at IS NULL
        """
        with get_connection() as connection:
            row = connection.execute(query, (reference_id,)).fetchone()

        if not row:
            raise RepositoryNotFoundError(f"Reference {reference_id} was not found")
        return row

    def _reference_from_row(self, row: dict) -> ReferenceRead:
        return ReferenceRead(
            id=row["id"],
            tesis_id=row["tesis_id"],
            version=row["version"],
            created_at=row["created_at"],
            updated_at=row["updated_at"],
            deleted_at=row["deleted_at"],
            **row["data"],
        )
