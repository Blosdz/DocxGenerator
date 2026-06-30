from datetime import UTC, date, datetime
from uuid import uuid4

from app.models.references import Author, CitationStyle, ReferenceRead, ReferenceType
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


def test_formats_vancouver_references_with_numbers() -> None:
    reference = make_reference(ReferenceType.BOOK, publisher="McGraw-Hill")

    formatted = CitationService().format_references([reference], CitationStyle.VANCOUVER)

    assert formatted == ["1. Hernández R. Metodología de la investigación. McGraw-Hill; 2014."]


def test_formats_ieee_references_with_numbers() -> None:
    reference = make_reference(ReferenceType.ARTICLE, journal="Revista de Ingeniería")

    formatted = CitationService().format_references([reference], CitationStyle.IEEE)

    assert formatted == ['[1] R. Hernández, "Metodología de la investigación," Revista de Ingeniería, 2014.']


def test_formats_iso690_book() -> None:
    reference = make_reference(ReferenceType.BOOK, publisher="McGraw-Hill")

    formatted = CitationService().format_reference(reference, CitationStyle.ISO690)

    assert formatted == "HERNÁNDEZ, R. Metodología de la investigación. McGraw-Hill, 2014."


def test_formats_book_in_mla() -> None:
    reference = make_reference(ReferenceType.BOOK, publisher="McGraw-Hill")

    formatted = CitationService().format_reference(reference, CitationStyle.MLA)

    assert formatted == "Hernández, Roberto. Metodología de la investigación. McGraw-Hill, 2014."


def test_formats_article_in_mla_with_volume_issue_pages() -> None:
    reference = make_reference(
        ReferenceType.ARTICLE,
        title="Aprendizaje automático aplicado",
        journal="Revista de Ingeniería",
        volume="12",
        issue="3",
        pages="45-67",
    )

    formatted = CitationService().format_reference(reference, CitationStyle.MLA)

    assert formatted == (
        'Hernández, Roberto. "Aprendizaje automático aplicado." '
        "Revista de Ingeniería, vol. 12, no. 3, 2014, pp. 45-67."
    )


def test_formats_mla_with_two_authors() -> None:
    reference = make_reference(
        ReferenceType.BOOK,
        publisher="McGraw-Hill",
        authors=[
            Author(first_name="Roberto", last_name="Hernández"),
            Author(first_name="María", last_name="López"),
        ],
    )

    formatted = CitationService().format_reference(reference, CitationStyle.MLA)

    assert formatted.startswith("Hernández, Roberto, and María López.")


def test_formats_web_in_mla_with_access_date() -> None:
    reference = make_reference(
        ReferenceType.WEB,
        title="Guía de investigación",
        url="https://example.com/guia",
        accessed_at=date(2026, 5, 26),
    )

    formatted = CitationService().format_reference(reference, CitationStyle.MLA)

    assert "Accessed 2026-05-26" in formatted
    assert "https://example.com/guia" in formatted


def test_apa7_article_includes_volume_issue_pages() -> None:
    reference = make_reference(
        ReferenceType.ARTICLE,
        title="Aprendizaje automático aplicado",
        journal="Revista de Ingeniería",
        volume="12",
        issue="3",
        pages="45-67",
    )

    formatted = CitationService().format_reference(reference)

    assert formatted == (
        "Hernández, R. (2014). Aprendizaje automático aplicado. "
        "Revista de Ingeniería, 12(3), 45-67."
    )


def test_ieee_article_includes_volume_issue_pages() -> None:
    reference = make_reference(
        ReferenceType.ARTICLE,
        journal="Revista de Ingeniería",
        volume="12",
        issue="3",
        pages="45-67",
    )

    formatted = CitationService().format_references([reference], CitationStyle.IEEE)

    assert formatted == [
        '[1] R. Hernández, "Metodología de la investigación," '
        "Revista de Ingeniería, vol. 12, no. 3, pp. 45-67, 2014."
    ]
