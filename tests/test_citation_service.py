from datetime import UTC, date, datetime
from uuid import uuid4

from app.models.references import Author, ReferenceRead, ReferenceType
from app.services.citation_service import CitationService


def make_reference(reference_type: ReferenceType, **overrides) -> ReferenceRead:
    payload = {
        "id": uuid4(),
        "tesis_id": uuid4(),
        "authors": [Author(first_name="Roberto", last_name="Hernández")],
        "year": 2014,
        "title": "Metodología de la investigación",
        "type": reference_type,
        "publisher": None,
        "journal": None,
        "doi": None,
        "url": None,
        "accessed_at": None,
        "version": 1,
        "created_at": datetime.now(UTC),
        "updated_at": datetime.now(UTC),
    }
    payload.update(overrides)
    return ReferenceRead(**payload)


def test_formats_book_in_apa7() -> None:
    reference = make_reference(ReferenceType.BOOK, publisher="McGraw-Hill")

    formatted = CitationService().format_reference(reference)

    assert formatted == "Hernández, R. (2014). Metodología de la investigación. McGraw-Hill."


def test_formats_article_in_apa7_with_doi() -> None:
    reference = make_reference(
        ReferenceType.ARTICLE,
        title="Aprendizaje automático aplicado",
        journal="Revista de Ingeniería",
        doi="10.1234/example",
    )

    formatted = CitationService().format_reference(reference)

    assert "Revista de Ingeniería" in formatted
    assert "https://doi.org/10.1234/example" in formatted


def test_formats_web_reference_with_access_date() -> None:
    reference = make_reference(
        ReferenceType.WEB,
        title="Guía de investigación",
        url="https://example.com/guia",
        accessed_at=date(2026, 5, 26),
    )

    formatted = CitationService().format_reference(reference)

    assert "Recuperado el 2026-05-26" in formatted
    assert "https://example.com/guia" in formatted
