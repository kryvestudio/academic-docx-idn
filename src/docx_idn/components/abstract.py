from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


class Abstract:
    def __init__(self, abstrak_id: str = "", abstrak_en: str = ""):
        self.abstrak_id = abstrak_id
        self.abstrak_en = abstrak_en

    def render(self, doc):
        if self.abstrak_id:
            self._render_section(doc, "ABSTRAK", self.abstrak_id)
        if self.abstrak_en:
            self._render_section(doc, "ABSTRACT", self.abstrak_en)

    def _render_section(self, doc, title: str, text: str):
        doc.add_heading(title, level=1)
        doc.add_paragraph(text)
        doc.add_page_break()
