from pathlib import Path

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.shared import Pt

from app.core.config import get_settings
from app.models.references import CitationStyle


class TemplateService:
    def __init__(self) -> None:
        self.templates_dir = get_settings().templates_dir

    def get_template_path(self, style: CitationStyle = CitationStyle.APA7) -> Path:
        filename = {
            CitationStyle.APA7: "apa7.docx",
            CitationStyle.IEEE: "ieee.docx",
            CitationStyle.VANCOUVER: "vancouver.docx",
            CitationStyle.ISO690: "iso690.docx",
            CitationStyle.MLA: "mla.docx",
        }[style]
        path = self.templates_dir / filename

        self._ensure_valid_docx(path)

        return path

    def _ensure_valid_docx(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists() and path.stat().st_size > 0:
            try:
                Document(path)
                return
            except PackageNotFoundError:
                pass

        document = Document()
        styles = document.styles
        styles["Normal"].font.name = "Times New Roman"
        styles["Normal"].font.size = Pt(12)
        document.save(path)
