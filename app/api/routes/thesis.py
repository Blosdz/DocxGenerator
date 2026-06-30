from uuid import UUID

from fastapi import APIRouter, status

from app.models.thesis import (
    AppendSectionTextRequest,
    FromSkeletonRequest,
    SectionCreate,
    SectionRead,
    SectionUpdate,
    ThesisCreate,
    ThesisRead,
)
from app.repositories.thesis_repository import ThesisRepository


router = APIRouter(prefix="/theses", tags=["theses"])


def get_repository() -> ThesisRepository:
    return ThesisRepository()


@router.post("", response_model=ThesisRead, status_code=status.HTTP_201_CREATED)
async def create_thesis(payload: ThesisCreate) -> ThesisRead:
    return get_repository().create(payload)


@router.get("/{tesis_id}", response_model=ThesisRead)
async def get_thesis(tesis_id: UUID) -> ThesisRead:
    return get_repository().get(tesis_id)


@router.post("/{tesis_id}/sections", response_model=SectionRead, status_code=status.HTTP_201_CREATED)
async def create_section(tesis_id: UUID, payload: SectionCreate) -> SectionRead:
    return get_repository().create_section(tesis_id, payload)


@router.get("/{tesis_id}/sections", response_model=list[SectionRead])
async def list_sections(tesis_id: UUID) -> list[SectionRead]:
    return get_repository().list_sections(tesis_id)


@router.put("/{tesis_id}/sections", response_model=list[SectionRead])
async def replace_sections(tesis_id: UUID, payload: list[SectionCreate]) -> list[SectionRead]:
    return get_repository().replace_sections(tesis_id, payload)


@router.post("/{tesis_id}/sections/from-skeleton", response_model=list[SectionRead], status_code=status.HTTP_201_CREATED)
async def apply_skeleton_sections(tesis_id: UUID, payload: FromSkeletonRequest) -> list[SectionRead]:
    return get_repository().replace_sections_from_skeleton(tesis_id, payload.sections)


@router.patch("/{tesis_id}/sections/{section_id}", response_model=SectionRead)
async def update_section(
    tesis_id: UUID,
    section_id: UUID,
    payload: SectionUpdate,
) -> SectionRead:
    return get_repository().update_section(tesis_id, section_id, payload)


@router.post("/{tesis_id}/sections/{section_id}/append-text", response_model=SectionRead)
async def append_section_text(
    tesis_id: UUID,
    section_id: UUID,
    payload: AppendSectionTextRequest,
) -> SectionRead:
    return get_repository().append_section_text(tesis_id, section_id, payload.text)


@router.delete("/{tesis_id}/sections/{section_id}")
async def delete_section(tesis_id: UUID, section_id: UUID) -> dict[str, UUID | bool]:
    return get_repository().delete_section(tesis_id, section_id)
