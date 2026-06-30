from dataclasses import dataclass
from datetime import date
import re
from zipfile import BadZipFile, ZipFile
import xml.etree.ElementTree as ET
from uuid import UUID

from docx import Document

from app.models.references import Author, CitationStyle, ReferenceCreate, ReferenceType
from app.repositories.documents_repository import DocumentsRepository
from app.repositories.references_repository import ReferencesRepository


REFERENCE_HEADINGS = {
    "referencias",
    "bibliografia",
    "bibliografía",
    "references",
    "bibliography",
}

WORD_SOURCE_TYPES = {
    "book": ReferenceType.BOOK,
    "booksection": ReferenceType.BOOK,
    "journalarticle": ReferenceType.ARTICLE,
    "articleinaperiodical": ReferenceType.ARTICLE,
    "internetsite": ReferenceType.WEB,
    "website": ReferenceType.WEB,
    "documentfrominternetsite": ReferenceType.WEB,
    "electronicsource": ReferenceType.WEB,
}


@dataclass
class ExtractedReference:
    payload: ReferenceCreate
    raw_text: str
    source: str


class ReferenceExtractionService:
    def __init__(
        self,
        documents_repository: DocumentsRepository | None = None,
        references_repository: ReferencesRepository | None = None,
    ) -> None:
        self.documents_repository = documents_repository or DocumentsRepository()
        self.references_repository = references_repository or ReferencesRepository()

    def extract_and_create(self, document_id: UUID) -> dict:
        context = self.documents_repository.get_editable_document_context(document_id)
        tesis_id = context["tesis_id"]
        existing_keys = {
            self._reference_key(reference.title, reference.year, reference.doi)
            for reference in self.references_repository.list_by_thesis(tesis_id)
        }

        extracted = self.extract_from_path(context["path"])
        created_count = 0
        skipped_count = 0
        references: list[dict] = []

        for item in extracted:
            key = self._reference_key(item.payload.title, item.payload.year, item.payload.doi)
            if key in existing_keys:
                skipped_count += 1
                references.append(self._reference_summary(item, "skipped", reason="already_exists"))
                continue

            created = self.references_repository.create(tesis_id, item.payload)
            existing_keys.add(key)
            created_count += 1
            references.append(self._reference_summary(item, "created", reference_id=getattr(created, "id", None)))

        return {
            "document_id": context["document_id"],
            "tesis_id": tesis_id,
            "extracted_count": len(extracted),
            "created_count": created_count,
            "skipped_count": skipped_count,
            "references": references,
        }

    def extract_from_path(self, path) -> list[ExtractedReference]:
        document = Document(path)
        paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs]
        reference_texts = self._reference_paragraphs(paragraphs)
        text_references = [
            parsed
            for parsed in (self._parse_reference(text) for text in reference_texts)
            if parsed is not None
        ]
        metadata_references = self._metadata_references(path)
        return self._deduplicate_extracted([*metadata_references, *text_references])

    def _reference_paragraphs(self, paragraphs: list[str]) -> list[str]:
        in_references = False
        entries: list[str] = []

        for text in paragraphs:
            clean = self._clean_spaces(text)
            if not clean:
                continue

            heading = self._normalize_heading(clean)
            if heading in REFERENCE_HEADINGS:
                in_references = True
                continue

            if not in_references:
                continue

            if self._looks_like_next_section(clean):
                break

            entry = self._strip_numbering(clean)
            if entry:
                entries.append(entry)

        return entries

    def _parse_reference(self, text: str) -> ExtractedReference | None:
        clean = self._clean_spaces(text)
        if len(clean) < 12:
            return None

        year = self._extract_year(clean)
        doi = self._extract_doi(clean)
        url = self._extract_url(clean)
        authors_text, remainder = self._split_authors(clean)
        authors = self._parse_authors(authors_text)
        title = self._extract_title(remainder)
        if not authors or not title:
            return None

        reference_type = ReferenceType.WEB if url else ReferenceType.ARTICLE
        journal = self._extract_container(remainder, title) if reference_type == ReferenceType.ARTICLE else None

        payload = ReferenceCreate(
            authors=authors,
            year=year,
            title=title,
            type=reference_type,
            journal=journal,
            doi=doi,
            url=url,
            style=CitationStyle.APA7,
        )
        return ExtractedReference(payload=payload, raw_text=text, source="text")

    def _metadata_references(self, path) -> list[ExtractedReference]:
        try:
            with ZipFile(path) as archive:
                names = [
                    name
                    for name in archive.namelist()
                    if name.startswith("customXml/")
                    and name.endswith(".xml")
                    and "/_rels/" not in name
                    and not name.endswith(".xml.rels")
                    and "itemProps" not in name
                ]
                references: list[ExtractedReference] = []
                for name in names:
                    try:
                        references.extend(self._metadata_references_from_xml(archive.read(name)))
                    except (KeyError, ET.ParseError):
                        continue
                return references
        except (BadZipFile, OSError):
            return []

    def _metadata_references_from_xml(self, xml: bytes) -> list[ExtractedReference]:
        root = ET.fromstring(xml)
        style = self._citation_style_from_sources(root)
        return [
            reference
            for source in root.iter()
            if self._local_name(source.tag) == "Source"
            for reference in [self._parse_word_source(source, style)]
            if reference is not None
        ]

    def _parse_word_source(
        self,
        source: ET.Element,
        style: CitationStyle,
    ) -> ExtractedReference | None:
        title = self._clean_spaces(self._child_text(source, "Title") or "")
        authors = self._word_authors(source)
        if not title or not authors:
            return None

        year = self._int_or_none(self._child_text(source, "Year"))
        doi = self._normalize_doi(self._child_text(source, "DOI"))
        url = self._clean_spaces(self._child_text(source, "URL") or "") or None
        reference_type = self._word_reference_type(source, url)
        accessed_at = self._word_accessed_at(source)

        payload = ReferenceCreate(
            authors=authors,
            year=year,
            title=title[:500],
            type=reference_type,
            publisher=self._clean_spaces(self._child_text(source, "Publisher") or "") or None,
            journal=self._clean_spaces(self._child_text(source, "JournalName") or "") or None,
            doi=doi,
            url=url,
            accessed_at=accessed_at,
            style=style,
        )
        return ExtractedReference(
            payload=payload,
            raw_text=self._metadata_raw_text(source),
            source="metadata",
        )

    def _word_authors(self, source: ET.Element) -> list[Author]:
        authors: list[Author] = []

        for person in source.iter():
            if self._local_name(person.tag) != "Person":
                continue
            last_name = self._clean_spaces(self._child_text(person, "Last") or "")
            first_name = self._clean_spaces(
                " ".join(
                    value
                    for value in [
                        self._child_text(person, "First"),
                        self._child_text(person, "Middle"),
                    ]
                    if value
                )
            )
            if last_name:
                authors.append(Author(last_name=last_name, first_name=first_name or None))

        if authors:
            return authors[:8]

        corporate_names = [
            self._clean_spaces(element.text or "")
            for element in source.iter()
            if self._local_name(element.tag) == "Corporate" and self._clean_spaces(element.text or "")
        ]
        if corporate_names:
            return [Author(last_name=name) for name in corporate_names[:8]]

        # Fallback: use ProductionCompany as institutional author
        production_company = self._clean_spaces(self._child_text(source, "ProductionCompany") or "")
        if production_company:
            return [Author(last_name=production_company)]

        return []

    def _word_reference_type(self, source: ET.Element, url: str | None) -> ReferenceType:
        source_type = self._clean_spaces(self._child_text(source, "SourceType") or "")
        normalized = re.sub(r"[^a-z0-9]+", "", source_type.lower())
        mapped = WORD_SOURCE_TYPES.get(source_type) or WORD_SOURCE_TYPES.get(normalized)
        if mapped:
            return mapped
        if url:
            return ReferenceType.WEB
        if self._child_text(source, "Publisher"):
            return ReferenceType.BOOK
        return ReferenceType.ARTICLE

    def _word_accessed_at(self, source: ET.Element) -> date | None:
        year = self._int_or_none(self._child_text(source, "YearAccessed"))
        month = self._int_or_none(self._child_text(source, "MonthAccessed"))
        day = self._int_or_none(self._child_text(source, "DayAccessed"))
        if not year:
            return None

        try:
            return date(year, month or 1, day or 1)
        except ValueError:
            return None

    def _metadata_raw_text(self, source: ET.Element) -> str:
        parts = [
            self._child_text(source, "Title"),
            self._child_text(source, "Year"),
            self._child_text(source, "DOI"),
            self._child_text(source, "URL"),
        ]
        return " | ".join(self._clean_spaces(part or "") for part in parts if part)

    def _citation_style_from_sources(self, root: ET.Element) -> CitationStyle:
        raw = " ".join(
            str(root.attrib.get(key, ""))
            for key in ("SelectedStyle", "StyleName")
        ).upper()
        if "VANCOUVER" in raw:
            return CitationStyle.VANCOUVER
        if "IEEE" in raw:
            return CitationStyle.IEEE
        if "ISO" in raw:
            return CitationStyle.ISO690
        return CitationStyle.APA7

    def _deduplicate_extracted(self, references: list[ExtractedReference]) -> list[ExtractedReference]:
        seen = set()
        result: list[ExtractedReference] = []
        for item in references:
            key = self._reference_key(item.payload.title, item.payload.year, item.payload.doi)
            if key in seen:
                continue
            seen.add(key)
            result.append(item)
        return result

    def _reference_summary(
        self,
        item: ExtractedReference,
        status: str,
        reference_id: UUID | None = None,
        reason: str | None = None,
    ) -> dict:
        return {
            "id": reference_id,
            "title": item.payload.title,
            "year": item.payload.year,
            "type": item.payload.type,
            "source": item.source,
            "status": status,
            "reason": reason,
        }

    def _split_authors(self, text: str) -> tuple[str, str]:
        apa_match = re.search(r"\(\s*(\d{4}|s\.?\s*f\.?|n\.?\s*d\.?)\s*\)", text, re.IGNORECASE)
        if apa_match:
            return text[: apa_match.start()].strip(" ."), text[apa_match.end() :].strip(" .")

        year_match = re.search(r"\b(1[5-9]|20)\d{2}\b", text)
        if year_match:
            before_year = text[: year_match.start()].strip(" .,")
            parts = before_year.split(".")
            if parts:
                return parts[0].strip(), text[year_match.end() :].strip(" .")

        parts = text.split(".", 1)
        if len(parts) == 2:
            return parts[0].strip(), parts[1].strip()
        return "", text

    def _parse_authors(self, text: str) -> list[Author]:
        text = re.sub(r"\s+&\s+", ", ", text)
        text = re.sub(r"\s+y\s+", ", ", text, flags=re.IGNORECASE)

        apa_pairs = re.findall(
            r"([^,.;]+),\s*((?:[A-ZÁÉÍÓÚÑ]\.?\s*){1,5})",
            text,
        )
        if apa_pairs:
            return [
                Author(
                    last_name=last_name.strip(),
                    first_name=self._normalize_initials(initials) or None,
                )
                for last_name, initials in apa_pairs[:8]
                if last_name.strip()
            ]

        chunks = [chunk.strip(" ,.") for chunk in text.split(";") if chunk.strip(" ,.")]
        if len(chunks) == 1:
            chunks = [chunk.strip(" ,.") for chunk in re.split(r",\s+(?=[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúñ'-]+(?:\s|,|$))", chunks[0]) if chunk.strip(" ,.")]

        authors = []
        for chunk in chunks[:8]:
            if not chunk or chunk.lower() in {"et al", "et al."}:
                continue
            authors.append(self._author_from_text(chunk))
        return [author for author in authors if author.last_name]

    def _author_from_text(self, text: str) -> Author:
        parts = [part.strip() for part in text.split(",", 1)]
        if len(parts) == 2 and parts[0]:
            return Author(last_name=parts[0], first_name=parts[1] or None)

        words = text.split()
        if len(words) > 1:
            return Author(last_name=words[-1], first_name=" ".join(words[:-1]))
        return Author(last_name=text)

    def _normalize_initials(self, value: str) -> str:
        letters = re.findall(r"[A-ZÁÉÍÓÚÑ]", value or "")
        return " ".join(f"{letter}." for letter in letters)

    def _extract_title(self, text: str) -> str | None:
        clean = re.sub(r"^\(?\s*(\d{4}|s\.?\s*f\.?|n\.?\s*d\.?)\s*\)?\.?", "", text, flags=re.IGNORECASE).strip()
        clean = re.sub(r"https?://\S+", "", clean)
        clean = re.sub(r"doi:\s*\S+|https?://doi\.org/\S+", "", clean, flags=re.IGNORECASE).strip(" .")
        if not clean:
            return None

        parts = [part.strip() for part in clean.split(".") if part.strip()]
        return parts[0][:500] if parts else clean[:500]

    def _extract_container(self, text: str, title: str) -> str | None:
        after_title = text.replace(title, "", 1).strip(" .")
        if not after_title:
            return None
        parts = [part.strip() for part in after_title.split(".") if part.strip()]
        return parts[0][:300] if parts else None

    def _extract_year(self, text: str) -> int | None:
        match = re.search(r"\b(1[5-9]|20)\d{2}\b", text)
        return int(match.group(0)) if match else None

    def _extract_doi(self, text: str) -> str | None:
        match = re.search(r"(?:doi:\s*|https?://doi\.org/)(10\.\S+)", text, re.IGNORECASE)
        return self._normalize_doi(match.group(1)) if match else None

    def _extract_url(self, text: str) -> str | None:
        match = re.search(r"https?://\S+", text)
        return match.group(0).rstrip(" .") if match else None

    def _reference_key(self, title: str, year: int | None, doi: str | None = None) -> tuple:
        normalized_doi = self._normalize_doi(doi)
        if normalized_doi:
            return ("doi", normalized_doi.lower())
        normalized = re.sub(r"[^a-z0-9]+", " ", title.lower()).strip()
        return ("title", normalized, year)

    def _strip_numbering(self, text: str) -> str:
        return re.sub(r"^\s*(\[\d+\]|\d+[.)])\s*", "", text).strip()

    def _looks_like_next_section(self, text: str) -> bool:
        return bool(re.match(r"^(anexos?|ap[eé]ndices?|appendix)\b", text, re.IGNORECASE))

    def _normalize_heading(self, text: str) -> str:
        return self._clean_spaces(text).lower().strip(" .:")

    def _clean_spaces(self, text: str) -> str:
        return re.sub(r"\s+", " ", str(text or "")).strip()

    def _child_text(self, element: ET.Element, child_name: str) -> str | None:
        for child in element:
            if self._local_name(child.tag) == child_name:
                return child.text
        return None

    def _local_name(self, tag: str) -> str:
        return str(tag).rsplit("}", 1)[-1]

    def _int_or_none(self, value: str | None) -> int | None:
        if value is None:
            return None
        match = re.search(r"\d{1,4}", str(value))
        return int(match.group(0)) if match else None

    def _normalize_doi(self, value: str | None) -> str | None:
        clean = self._clean_spaces(value or "")
        clean = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", clean, flags=re.IGNORECASE)
        clean = re.sub(r"^doi:\s*", "", clean, flags=re.IGNORECASE)
        return clean.rstrip(" .") or None
