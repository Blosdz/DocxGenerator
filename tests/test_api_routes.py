import asyncio
from datetime import UTC, datetime
from pathlib import Path
from uuid import uuid4

import httpx

from app.api.routes import documents, references, thesis
from app.main import app
from app.models.documents import DocumentResponse
from app.models.references import ReferenceDeleteResponse
from app.models.thesis import SectionRead, ThesisRead


def run(coro):
    return asyncio.run(coro)


def make_client() -> httpx.AsyncClient:
    transport = httpx.ASGITransport(app=app)
    return httpx.AsyncClient(transport=transport, base_url="http://test")


class FakeThesisRepository:
    def create(self, payload):
        now = datetime.now(UTC)
        data = payload.model_dump(mode="json")
        metadata = data.pop("metadata")
        metadata.update(data)
        return ThesisRead(id=uuid4(), data={"metadata": metadata}, version=1, created_at=now, updated_at=now)

    def get(self, tesis_id):
        now = datetime.now(UTC)
        return ThesisRead(id=tesis_id, data={"metadata": {"title": "Demo"}}, version=1, created_at=now, updated_at=now)

    def create_section(self, tesis_id, payload):
        now = datetime.now(UTC)
        return SectionRead(
            id=uuid4(),
            tesis_id=tesis_id,
            version=1,
            created_at=now,
            updated_at=now,
            **payload.model_dump(),
        )

    def list_sections(self, tesis_id):
        return []

    def replace_sections(self, tesis_id, payload):
        return [self.create_section(tesis_id, section) for section in payload]

    def update_section(self, tesis_id, section_id, payload):
        now = datetime.now(UTC)
        data = {
            "title": "Capítulo I actualizado",
            "subtitle": None,
            "level": 1,
            "content": "Texto",
            "order": 1,
        }
        data.update(payload.model_dump(exclude_unset=True))
        return SectionRead(
            id=section_id,
            tesis_id=tesis_id,
            version=2,
            created_at=now,
            updated_at=now,
            **data,
        )

    def delete_section(self, tesis_id, section_id):
        return {"id": section_id, "deleted": True}


class FakeReferencesRepository:
    def create(self, tesis_id, payload):
        now = datetime.now(UTC)
        return {
            "id": uuid4(),
            "tesis_id": tesis_id,
            "version": 1,
            "created_at": now,
            "updated_at": now,
            **payload.model_dump(mode="json"),
        }

    def list_by_thesis(self, tesis_id):
        return []

    def get(self, reference_id):
        now = datetime.now(UTC)
        return {
            "id": reference_id,
            "tesis_id": uuid4(),
            "authors": [{"first_name": "Ada", "last_name": "Lovelace"}],
            "year": 1843,
            "title": "Notas",
            "type": "book",
            "publisher": "Demo",
            "version": 1,
            "created_at": now,
            "updated_at": now,
        }

    def update(self, reference_id, payload):
        data = self.get(reference_id)
        data.update(payload.model_dump(mode="json", exclude_unset=True))
        data["version"] = 2
        return data

    def delete(self, reference_id):
        return ReferenceDeleteResponse(id=reference_id, deleted=True)


class FakeBuilder:
    def build_docx(
        self,
        tesis_id,
        upload_to_backend=False,
        backend_url=None,
        authorization=None,
    ):
        return DocumentResponse(
            filename="demo.docx",
            path=Path("app/generated/demo.docx"),
            download_url="/documents/demo.docx",
            generated_at=datetime.now(UTC),
            uploaded=upload_to_backend,
            upload={"backend_url": backend_url, "authorization": authorization}
            if upload_to_backend
            else None,
        )


def test_thesis_and_document_routes_use_services(monkeypatch) -> None:
    monkeypatch.setattr(thesis, "get_repository", lambda: FakeThesisRepository())
    monkeypatch.setattr(documents, "get_builder", lambda: FakeBuilder())
    tesis_id = uuid4()

    async def scenario():
        async with make_client() as client:
            created_response = await client.post("/theses", json={"title": "Demo"})
            section_response = await client.post(
                f"/theses/{tesis_id}/sections",
                json={"title": "Capítulo I", "level": 1, "content": "Texto", "order": 1},
            )
            index_response = await client.put(
                f"/theses/{tesis_id}/sections",
                json=[
                    {"title": "Capítulo I", "level": 1, "content": "Texto", "order": 1},
                    {"title": "Subtítulo", "level": 2, "content": "Texto", "order": 2},
                ],
            )
            generated_response = await client.post(f"/theses/{tesis_id}/documents/docx")
        return (
            created_response.json(),
            section_response.json(),
            index_response.json(),
            generated_response.json(),
        )

    created, section, index, generated = run(scenario())

    assert created["data"]["metadata"]["title"] == "Demo"
    assert section["title"] == "Capítulo I"
    assert len(index) == 2
    assert generated["download_url"] == "/documents/demo.docx"


def test_generate_docx_route_can_request_backend_upload(monkeypatch) -> None:
    monkeypatch.setattr(documents, "get_builder", lambda: FakeBuilder())
    tesis_id = uuid4()

    async def scenario():
        async with make_client() as client:
            response = await client.post(
                f"/theses/{tesis_id}/documents/docx",
                json={"upload_to_backend": True},
                headers={
                    "Authorization": "Bearer demo-token",
                    "X-Backend-Base-Url": "http://backend.test",
                },
            )
        return response.json()

    generated = run(scenario())

    assert generated["uploaded"] is True
    assert generated["upload"]["backend_url"] == "http://backend.test"
    assert generated["upload"]["authorization"] == "Bearer demo-token"


def test_reference_routes_use_repository(monkeypatch) -> None:
    monkeypatch.setattr(references, "get_repository", lambda: FakeReferencesRepository())
    tesis_id = uuid4()
    reference_id = uuid4()

    async def scenario():
        async with make_client() as client:
            created_response = await client.post(
                f"/theses/{tesis_id}/references",
                json={
                    "authors": [{"first_name": "Ada", "last_name": "Lovelace"}],
                    "year": 1843,
                    "title": "Notas",
                    "type": "book",
                    "publisher": "Demo",
                },
            )
            deleted_response = await client.delete(f"/references/{reference_id}")
            updated_response = await client.patch(
                f"/references/{reference_id}",
                json={"title": "Notas actualizadas"},
            )
        return created_response.json(), updated_response.json(), deleted_response.json()

    created, updated, deleted = run(scenario())

    assert created["title"] == "Notas"
    assert updated["title"] == "Notas actualizadas"
    assert deleted == {"id": str(reference_id), "deleted": True}
