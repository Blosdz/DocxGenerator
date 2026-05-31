from datetime import UTC, datetime
from pathlib import Path
import re
from tempfile import NamedTemporaryFile
import unicodedata
from uuid import UUID, uuid4
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from app.core.config import get_settings
from app.models.documents import DocumentResponse
from app.models.references import CitationStyle, ReferenceRead, ReferenceType
from app.models.thesis import SectionRead, ThesisRead
from app.services.citation_service import CitationService
from app.services.template_service import TemplateService

BIBLIOGRAPHY_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CUSTOM_XML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CUSTOM_XML_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
CUSTOM_XML_PROPS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
BIBLIOGRAPHY_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
BIBLIOGRAPHY_ITEM = "customXml/item1.xml"
BIBLIOGRAPHY_ITEM_PROPS = "customXml/itemProps1.xml"
BIBLIOGRAPHY_ITEM_RELS = "customXml/_rels/item1.xml.rels"
ROOT_RELS = "_rels/.rels"
CONTENT_TYPES = "[Content_Types].xml"
FIELD_LOCALE = "10250"
CITATION_TOKEN_PATTERN = re.compile(r"(\{\{cite:([A-Za-z0-9_-]+)\}\}|\[cite:([A-Za-z0-9_-]+)\])")

ET.register_namespace("b", BIBLIOGRAPHY_NS)
ET.register_namespace("ds", CUSTOM_XML_NS)


class DocxService:
    def __init__(
        self,
        template_service: TemplateService | None = None,
        citation_service: CitationService | None = None,
    ) -> None:
        settings = get_settings()
        self.generated_dir = settings.generated_dir
        self.template_service = template_service or TemplateService()
        self.citation_service = citation_service or CitationService()

    def generate(
        self,
        thesis: ThesisRead,
        sections: list[SectionRead],
        references: list[ReferenceRead],
        style: CitationStyle = CitationStyle.APA7,
    ) -> DocumentResponse:
        template_path = self.template_service.get_template_path(style)
        document = Document(template_path)
        self._apply_word_metadata(document, thesis)
        self._apply_document_styles(document)
        self._add_cover(document, thesis)
        self._add_table_of_contents(document)
        reference_tags = self._reference_tags(references)
        self._add_sections(document, sections, references, reference_tags)
        self._add_references(document, references, style)
        self._enable_auto_field_update(document)

        self.generated_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now(UTC).strftime("%Y%m%d%H%M%S")
        filename = f"{self._filename_stem(thesis)}-{timestamp}.docx"
        path = self.generated_dir / filename
        document.save(path)
        self._write_bibliography_metadata(path, references, style, reference_tags)

        return DocumentResponse(
            filename=filename,
            path=path,
            download_url=f"/documents/{filename}",
            generated_at=datetime.now(UTC),
        )

    def upload_filename(self, thesis: ThesisRead) -> str:
        return f"{self._filename_stem(thesis)}.docx"

    def _filename_stem(self, thesis: ThesisRead) -> str:
        title = thesis.thesis_metadata.get("title") or thesis.thesis_metadata.get("titulo")
        value = str(title or thesis.id)
        ascii_value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode()
        stem = re.sub(r"[^A-Za-z0-9._-]+", "-", ascii_value).strip(".-_")
        return (stem or "tesis")[:120]

    def _apply_word_metadata(self, document: Document, thesis: ThesisRead) -> None:
        metadata = thesis.thesis_metadata
        properties = document.core_properties
        title = metadata.get("title") or metadata.get("titulo")
        author = metadata.get("author") or metadata.get("autor")

        if title:
            properties.title = str(title)
            properties.subject = str(title)
        if author:
            properties.author = str(author)
            properties.last_modified_by = str(author)
        properties.category = "Tesis"
        properties.keywords = "tesis, APA, bibliografia, indice"
        properties.modified = datetime.now(UTC)

    def _apply_document_styles(self, document: Document) -> None:
        normal = document.styles["Normal"]
        self._set_style_font(normal, bold=False)
        self._set_style_paragraph_format(normal)

        for style_name in ("Title", "Heading 1", "Heading 2", "Heading 3"):
            style = document.styles[style_name]
            self._set_style_font(style, bold=True)
            self._set_style_paragraph_format(style)

    def _set_style_font(self, style, bold: bool) -> None:
        style.font.name = "Times New Roman"
        style.font.size = Pt(12)
        style.font.bold = bold
        style.font.color.rgb = RGBColor(0, 0, 0)

    def _set_style_paragraph_format(self, style) -> None:
        if style.type != WD_STYLE_TYPE.PARAGRAPH:
            return

        paragraph_format = style.paragraph_format
        paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)
        paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE

    def _format_paragraph(self, paragraph) -> None:
        paragraph_format = paragraph.paragraph_format
        paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)
        paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE

    def _reference_tags(self, references: list[ReferenceRead]) -> dict[UUID, str]:
        return {reference.id: f"Ref_{reference.id.hex[:16]}" for reference in references}

    def _add_cover(self, document: Document, thesis: ThesisRead) -> None:
        metadata = thesis.thesis_metadata
        self._format_paragraph(document.add_paragraph(metadata.get("institution") or ""))
        title = document.add_paragraph()
        title.style = document.styles["Title"]
        title.add_run(metadata.get("title") or "Tesis").bold = True
        self._format_paragraph(title)

        if metadata.get("author"):
            self._format_paragraph(document.add_paragraph(f"Autor: {metadata['author']}"))
        if metadata.get("year"):
            self._format_paragraph(document.add_paragraph(str(metadata["year"])))

        document.add_page_break()

    def _add_table_of_contents(self, document: Document) -> None:
        title = document.add_paragraph()
        title.add_run("Tabla de contenido").bold = True
        self._format_paragraph(title)
        paragraph = document.add_paragraph()
        self._format_paragraph(paragraph)
        self._add_field(
            paragraph,
            r'TOC \o "1-3" \h \z \u',
            "Actualice el campo en Word para generar el índice.",
        )
        document.add_page_break()

    def _enable_auto_field_update(self, document: Document) -> None:
        settings = document.settings.element
        update_fields = settings.find(qn("w:updateFields"))
        if update_fields is None:
            update_fields = OxmlElement("w:updateFields")
            settings.append(update_fields)
        update_fields.set(qn("w:val"), "true")

    def _add_field(self, paragraph, instruction: str, placeholder: str = "") -> None:
        run = paragraph.add_run()

        fld_char_begin = OxmlElement("w:fldChar")
        fld_char_begin.set(qn("w:fldCharType"), "begin")

        instr_text = OxmlElement("w:instrText")
        instr_text.set(qn("xml:space"), "preserve")
        instr_text.text = instruction

        fld_char_separate = OxmlElement("w:fldChar")
        fld_char_separate.set(qn("w:fldCharType"), "separate")

        fld_char_end = OxmlElement("w:fldChar")
        fld_char_end.set(qn("w:fldCharType"), "end")

        run._r.append(fld_char_begin)
        run._r.append(instr_text)
        run._r.append(fld_char_separate)
        if placeholder:
            text = OxmlElement("w:t")
            text.text = placeholder
            run._r.append(text)
        run._r.append(fld_char_end)

    def _add_sections(
        self,
        document: Document,
        sections: list[SectionRead],
        references: list[ReferenceRead],
        reference_tags: dict[UUID, str],
    ) -> None:
        reference_lookup = self._reference_lookup(references, reference_tags)
        for section in sorted(sections, key=lambda item: (item.order, item.created_at)):
            heading = section.title
            if section.subtitle:
                heading = f"{heading}: {section.subtitle}"
            paragraph = document.add_heading(heading, level=section.level)
            self._format_paragraph(paragraph)

            for block in section.content.split("\n\n"):
                text = block.strip()
                if text:
                    self._add_content_paragraph(document, text, reference_lookup)

    def _reference_lookup(
        self,
        references: list[ReferenceRead],
        reference_tags: dict[UUID, str],
    ) -> dict[str, tuple[str, str]]:
        lookup: dict[str, tuple[str, str]] = {}
        for reference in references:
            reference_id = reference.id
            value = (reference_tags[reference_id], self._citation_placeholder(reference))
            lookup[str(reference_id).lower()] = value
            lookup[reference_id.hex.lower()] = value
            lookup[reference_tags[reference_id].lower()] = value
        return lookup

    def _add_content_paragraph(
        self,
        document: Document,
        text: str,
        reference_lookup: dict[str, tuple[str, str]],
    ) -> None:
        paragraph = document.add_paragraph()
        self._format_paragraph(paragraph)

        cursor = 0
        for match in CITATION_TOKEN_PATTERN.finditer(text):
            if match.start() > cursor:
                paragraph.add_run(text[cursor : match.start()])

            token_value = (match.group(2) or match.group(3) or "").lower()
            citation = reference_lookup.get(token_value)
            if citation:
                tag, placeholder = citation
                self._add_field(
                    paragraph,
                    rf"CITATION {tag} \l {FIELD_LOCALE}",
                    placeholder,
                )
            else:
                paragraph.add_run(match.group(0))
            cursor = match.end()

        if cursor < len(text):
            paragraph.add_run(text[cursor:])

    def _citation_placeholder(self, reference: ReferenceRead) -> str:
        authors = [author.last_name for author in reference.authors if author.last_name]
        if not authors:
            author_text = reference.title
        elif len(authors) == 1:
            author_text = authors[0]
        elif len(authors) == 2:
            author_text = f"{authors[0]} & {authors[1]}"
        else:
            author_text = f"{authors[0]} et al."

        year = reference.year if reference.year is not None else "s. f."
        return f"({author_text}, {year})"

    def _add_references(
        self,
        document: Document,
        references: list[ReferenceRead],
        style: CitationStyle,
    ) -> None:
        if document.paragraphs:
            document.paragraphs[-1].add_run().add_break(WD_BREAK.PAGE)
        heading = document.add_heading("Referencias", level=1)
        self._format_paragraph(heading)

        paragraph = document.add_paragraph()
        self._format_paragraph(paragraph)
        self._add_field(
            paragraph,
            rf"BIBLIOGRAPHY \l {FIELD_LOCALE}",
            "Actualice el campo en Word para generar la bibliografía APA.",
        )

    def _write_bibliography_metadata(
        self,
        path: Path,
        references: list[ReferenceRead],
        style: CitationStyle,
        reference_tags: dict[UUID, str],
    ) -> None:
        with ZipFile(path, "r") as archive:
            names = set(archive.namelist())
            content_types_xml = archive.read(CONTENT_TYPES) if CONTENT_TYPES in names else None
            root_rels_xml = archive.read(ROOT_RELS) if ROOT_RELS in names else None
            item_rels_xml = archive.read(BIBLIOGRAPHY_ITEM_RELS) if BIBLIOGRAPHY_ITEM_RELS in names else None
            item_props_xml = archive.read(BIBLIOGRAPHY_ITEM_PROPS) if BIBLIOGRAPHY_ITEM_PROPS in names else None

        replacements = {
            BIBLIOGRAPHY_ITEM: self._build_sources_xml(references, style, reference_tags),
            BIBLIOGRAPHY_ITEM_PROPS: item_props_xml or self._build_item_props_xml(),
            BIBLIOGRAPHY_ITEM_RELS: self._ensure_item_rels_xml(item_rels_xml),
            CONTENT_TYPES: self._ensure_content_types_xml(content_types_xml),
            ROOT_RELS: self._ensure_root_rels_xml(root_rels_xml),
        }

        with NamedTemporaryFile(dir=path.parent, delete=False, suffix=".docx") as tmp:
            tmp_path = Path(tmp.name)

        try:
            with ZipFile(path, "r") as source, ZipFile(tmp_path, "w", ZIP_DEFLATED) as target:
                for item in source.infolist():
                    if item.filename in replacements:
                        continue
                    target.writestr(item, source.read(item.filename))

                for filename, payload in replacements.items():
                    target.writestr(filename, payload)
            tmp_path.replace(path)
        finally:
            if tmp_path.exists():
                tmp_path.unlink()

    def _build_sources_xml(
        self,
        references: list[ReferenceRead],
        style: CitationStyle,
        reference_tags: dict[UUID, str],
    ) -> bytes:
        sources = ET.Element(
            f"{{{BIBLIOGRAPHY_NS}}}Sources",
            {
                "SelectedStyle": "/APA.XSL" if style == CitationStyle.APA7 else f"/{style.value}.XSL",
                "StyleName": "APA" if style == CitationStyle.APA7 else style.value,
            },
        )

        for index, reference in enumerate(references, start=1):
            source = ET.SubElement(sources, f"{{{BIBLIOGRAPHY_NS}}}Source")
            self._append_bibliography_text(source, "Tag", reference_tags[reference.id])
            self._append_bibliography_text(source, "SourceType", self._word_source_type(reference.type))
            self._append_bibliography_text(source, "Guid", f"{{{str(reference.id).upper()}}}")
            self._append_bibliography_text(source, "LCID", "0")
            self._append_bibliography_text(source, "RefOrder", str(index))
            self._append_word_authors(source, reference)
            self._append_bibliography_text(source, "Title", reference.title)
            self._append_bibliography_text(source, "Year", str(reference.year) if reference.year else None)
            self._append_bibliography_text(source, "Publisher", reference.publisher)
            self._append_bibliography_text(source, "JournalName", reference.journal)
            self._append_bibliography_text(source, "URL", reference.url)
            self._append_bibliography_text(source, "DOI", reference.doi)

            if reference.accessed_at:
                self._append_bibliography_text(source, "YearAccessed", str(reference.accessed_at.year))
                self._append_bibliography_text(source, "MonthAccessed", str(reference.accessed_at.month))
                self._append_bibliography_text(source, "DayAccessed", str(reference.accessed_at.day))

        return ET.tostring(sources, encoding="UTF-8", xml_declaration=True)

    def _word_source_type(self, reference_type: ReferenceType) -> str:
        return {
            ReferenceType.BOOK: "Book",
            ReferenceType.ARTICLE: "JournalArticle",
            ReferenceType.WEB: "InternetSite",
        }[reference_type]

    def _append_word_authors(self, source: ET.Element, reference: ReferenceRead) -> None:
        author_root = ET.SubElement(source, f"{{{BIBLIOGRAPHY_NS}}}Author")
        author = ET.SubElement(author_root, f"{{{BIBLIOGRAPHY_NS}}}Author")
        name_list = ET.SubElement(author, f"{{{BIBLIOGRAPHY_NS}}}NameList")

        for item in reference.authors:
            person = ET.SubElement(name_list, f"{{{BIBLIOGRAPHY_NS}}}Person")
            self._append_bibliography_text(person, "Last", item.last_name)
            self._append_bibliography_text(person, "First", item.first_name)

    def _append_bibliography_text(self, parent: ET.Element, tag: str, value: object | None) -> None:
        if value is None or value == "":
            return
        element = ET.SubElement(parent, f"{{{BIBLIOGRAPHY_NS}}}{tag}")
        element.text = str(value)

    def _build_item_props_xml(self) -> bytes:
        item = ET.Element(
            f"{{{CUSTOM_XML_NS}}}datastoreItem",
            {f"{{{CUSTOM_XML_NS}}}itemID": f"{{{str(uuid4()).upper()}}}"},
        )
        schema_refs = ET.SubElement(item, f"{{{CUSTOM_XML_NS}}}schemaRefs")
        ET.SubElement(schema_refs, f"{{{CUSTOM_XML_NS}}}schemaRef", {f"{{{CUSTOM_XML_NS}}}uri": BIBLIOGRAPHY_NS})
        return ET.tostring(item, encoding="UTF-8", xml_declaration=True)

    def _ensure_item_rels_xml(self, xml: bytes | None) -> bytes:
        root = self._relationships_root(xml)
        if not self._has_relationship(root, CUSTOM_XML_PROPS_REL, "itemProps1.xml"):
            ET.SubElement(
                root,
                f"{{{RELATIONSHIPS_NS}}}Relationship",
                {
                    "Id": self._next_relationship_id(root),
                    "Type": CUSTOM_XML_PROPS_REL,
                    "Target": "itemProps1.xml",
                },
            )
        return ET.tostring(root, encoding="UTF-8", xml_declaration=True)

    def _ensure_root_rels_xml(self, xml: bytes | None) -> bytes:
        root = self._relationships_root(xml)
        if not self._has_relationship(root, CUSTOM_XML_REL, BIBLIOGRAPHY_ITEM):
            ET.SubElement(
                root,
                f"{{{RELATIONSHIPS_NS}}}Relationship",
                {
                    "Id": self._next_relationship_id(root),
                    "Type": CUSTOM_XML_REL,
                    "Target": BIBLIOGRAPHY_ITEM,
                },
            )
        return ET.tostring(root, encoding="UTF-8", xml_declaration=True)

    def _relationships_root(self, xml: bytes | None) -> ET.Element:
        if xml:
            return ET.fromstring(xml)
        return ET.Element(f"{{{RELATIONSHIPS_NS}}}Relationships")

    def _has_relationship(self, root: ET.Element, relationship_type: str, target: str) -> bool:
        for relationship in root.findall(f"{{{RELATIONSHIPS_NS}}}Relationship"):
            if relationship.get("Type") == relationship_type and relationship.get("Target") == target:
                return True
        return False

    def _next_relationship_id(self, root: ET.Element) -> str:
        used = {relationship.get("Id") for relationship in root.findall(f"{{{RELATIONSHIPS_NS}}}Relationship")}
        index = 1
        while f"rId{index}" in used:
            index += 1
        return f"rId{index}"

    def _ensure_content_types_xml(self, xml: bytes | None) -> bytes:
        if xml:
            root = ET.fromstring(xml)
        else:
            root = ET.Element(f"{{{CONTENT_TYPES_NS}}}Types")
            ET.SubElement(root, f"{{{CONTENT_TYPES_NS}}}Default", {"Extension": "xml", "ContentType": "application/xml"})
            ET.SubElement(
                root,
                f"{{{CONTENT_TYPES_NS}}}Default",
                {
                    "Extension": "rels",
                    "ContentType": "application/vnd.openxmlformats-package.relationships+xml",
                },
            )

        for override in root.findall(f"{{{CONTENT_TYPES_NS}}}Override"):
            if override.get("PartName") == f"/{BIBLIOGRAPHY_ITEM_PROPS}":
                override.set("ContentType", BIBLIOGRAPHY_CONTENT_TYPE)
                break
        else:
            ET.SubElement(
                root,
                f"{{{CONTENT_TYPES_NS}}}Override",
                {
                    "PartName": f"/{BIBLIOGRAPHY_ITEM_PROPS}",
                    "ContentType": BIBLIOGRAPHY_CONTENT_TYPE,
                },
            )

        return ET.tostring(root, encoding="UTF-8", xml_declaration=True)
