from datetime import UTC, datetime
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from docx import Document
from docx.oxml.ns import qn

from app.models.references import Author, ReferenceRead, ReferenceType
from app.models.thesis import SectionRead, ThesisRead
from app.services.docx_service import DocxService
from app.services.template_service import TemplateService


def test_template_service_recreates_empty_apa7_template(tmp_path) -> None:
    template_path = tmp_path / "apa7.docx"
    template_path.write_bytes(b"")
    service = TemplateService()
    service.templates_dir = tmp_path

    resolved = service.get_template_path()

    assert resolved == template_path
    assert template_path.stat().st_size > 0
    Document(template_path)


def test_docx_generation_writes_heading_levels(tmp_path) -> None:
    template_service = TemplateService()
    template_service.templates_dir = tmp_path / "templates"
    docx_service = DocxService(template_service=template_service)
    docx_service.generated_dir = tmp_path / "generated"
    tesis_id = uuid4()
    now = datetime.now(UTC)

    thesis = ThesisRead(
        id=tesis_id,
        data={
            "metadata": {
                "title": "Tesis demo",
                "author": "Ada Lovelace",
                "institution": "Universidad Demo",
            }
        },
        version=1,
        created_at=now,
        updated_at=now,
    )
    sections = [
        SectionRead(
            id=uuid4(),
            tesis_id=tesis_id,
            title="Capítulo I",
            subtitle=None,
            level=1,
            content="Contenido inicial.",
            order=1,
            version=1,
            created_at=now,
            updated_at=now,
        ),
        SectionRead(
            id=uuid4(),
            tesis_id=tesis_id,
            title="Problema general",
            subtitle=None,
            level=2,
            content="Detalle.",
            order=2,
            version=1,
            created_at=now,
            updated_at=now,
        ),
        SectionRead(
            id=uuid4(),
            tesis_id=tesis_id,
            title="Problema específico",
            subtitle=None,
            level=3,
            content="Detalle específico.",
            order=3,
            version=1,
            created_at=now,
            updated_at=now,
        ),
    ]
    reference_id = uuid4()
    references = [
        ReferenceRead(
            id=reference_id,
            tesis_id=tesis_id,
            authors=[Author(first_name="Roberto", last_name="Hernández")],
            year=2014,
            title="Metodología de la investigación",
            type=ReferenceType.BOOK,
            publisher="McGraw-Hill",
            version=1,
            created_at=now,
            updated_at=now,
        )
    ]
    sections[0].content = f"Contenido inicial con cita [cite:{reference_id}]."

    response = docx_service.generate(thesis, sections, references)
    document = Document(response.path)
    styles_by_text = {paragraph.text: paragraph.style.name for paragraph in document.paragraphs}
    heading_style = document.styles["Heading 1"]
    normal_style = document.styles["Normal"]

    assert response.path.exists()
    assert document.settings.element.find(qn("w:updateFields")).get(qn("w:val")) == "true"
    assert document.core_properties.title == "Tesis demo"
    assert normal_style.font.name == "Times New Roman"
    assert normal_style.paragraph_format.space_before.pt == 0
    assert normal_style.paragraph_format.space_after.pt == 0
    assert heading_style.font.name == "Times New Roman"
    assert heading_style.font.bold is True
    assert heading_style.font.size.pt == 12
    assert styles_by_text["Tesis demo"] == "Title"
    assert styles_by_text["Tabla de contenido"] != "Heading 1"
    assert styles_by_text["Capítulo I"] == "Heading 1"
    assert styles_by_text["Problema general"] == "Heading 2"
    assert styles_by_text["Problema específico"] == "Heading 3"

    with ZipFile(response.path) as archive:
        document_xml = archive.read("word/document.xml").decode()
        sources_xml = archive.read("customXml/item1.xml")
        root_rels_xml = archive.read("_rels/.rels").decode()

    source_root = ET.fromstring(sources_xml)
    namespace = {"b": "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"}
    source = source_root.find("b:Source", namespace)

    assert "TOC" in document_xml
    assert "BIBLIOGRAPHY" in document_xml
    assert f"CITATION Ref_{reference_id.hex[:16]}" in document_xml
    assert source_root.get("SelectedStyle") == "/APA.XSL"
    assert source is not None
    assert source.findtext("b:SourceType", namespaces=namespace) == "Book"
    assert source.findtext("b:Tag", namespaces=namespace) == f"Ref_{reference_id.hex[:16]}"
    assert source.findtext("b:Title", namespaces=namespace) == "Metodología de la investigación"
    assert source.findtext("b:Publisher", namespaces=namespace) == "McGraw-Hill"
    assert "relationships/customXml" in root_rels_xml
