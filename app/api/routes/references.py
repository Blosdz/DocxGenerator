from uuid import UUID

from fastapi import APIRouter, status

from app.models.references import ReferenceCreate, ReferenceDeleteResponse, ReferenceRead, ReferenceUpdate
from app.repositories.references_repository import ReferencesRepository


router = APIRouter(tags=["references"])


def get_repository() -> ReferencesRepository:
    return ReferencesRepository()


@router.post(
    "/theses/{tesis_id}/references",
    response_model=ReferenceRead,
    status_code=status.HTTP_201_CREATED,
)
async def create_reference(tesis_id: UUID, payload: ReferenceCreate) -> ReferenceRead:
    return get_repository().create(tesis_id, payload)


@router.get("/theses/{tesis_id}/references", response_model=list[ReferenceRead])
async def list_references(tesis_id: UUID) -> list[ReferenceRead]:
    return get_repository().list_by_thesis(tesis_id)


@router.get("/references/{reference_id}", response_model=ReferenceRead)
async def get_reference(reference_id: UUID) -> ReferenceRead:
    return get_repository().get(reference_id)


@router.patch("/references/{reference_id}", response_model=ReferenceRead)
async def update_reference(reference_id: UUID, payload: ReferenceUpdate) -> ReferenceRead:
    return get_repository().update(reference_id, payload)


@router.delete("/references/{reference_id}", response_model=ReferenceDeleteResponse)
async def delete_reference(reference_id: UUID) -> ReferenceDeleteResponse:
    return get_repository().delete(reference_id)
