from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx_idn.core.config import DocumentConfig
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


class StyleManager:
    def __init__(self, doc: Document, config: DocumentConfig):
        self.doc = doc
        self.cfg = config

    def setup_page(self):
        section = self.doc.sections[0]
        section.page_width = Cm(self.cfg.paper_width / 10)
        section.page_height = Cm(self.cfg.paper_height / 10)
        section.left_margin = Cm(self.cfg.margin_left)
        section.right_margin = Cm(self.cfg.margin_right)
        section.top_margin = Cm(self.cfg.margin_top)
        section.bottom_margin = Cm(self.cfg.margin_bottom)

    def setup_styles(self):
        self._setup_normal()
        self._setup_heading1()
        self._setup_heading2()
        self._setup_heading3()
        self._setup_toc_styles()

    def _set_font_times_new_roman(self, style):
        """Set font to Times New Roman including East Asian font."""
        style.font.name = self.cfg.font_name
        # Set East Asian font
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        rFonts.set(qn("w:ascii"), "Times New Roman")
        rFonts.set(qn("w:hAnsi"), "Times New Roman")
        rFonts.set(qn("w:eastAsia"), "Times New Roman")
        rFonts.set(qn("w:cs"), "Times New Roman")

    def _setup_normal(self):
        style = self.doc.styles["Normal"]
        self._set_font_times_new_roman(style)
        style.font.size = Pt(self.cfg.font_size)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.line_spacing = self.cfg.line_spacing
        style.paragraph_format.first_line_indent = Cm(self.cfg.first_line_indent)
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after = Pt(0)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def _setup_heading1(self):
        style = self.doc.styles["Heading 1"]
        self._set_font_times_new_roman(style)
        style.font.size = Pt(self.cfg.font_size_h1)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.space_before = Pt(self.cfg.space_before_heading1)
        style.paragraph_format.space_after = Pt(self.cfg.space_after_heading1)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style.paragraph_format.first_line_indent = Cm(0)
        style.paragraph_format.line_spacing = 1.5

    def _setup_heading2(self):
        style = self.doc.styles["Heading 2"]
        self._set_font_times_new_roman(style)
        style.font.size = Pt(self.cfg.font_size_h2)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.space_before = Pt(12)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        style.paragraph_format.first_line_indent = Cm(0)
        style.paragraph_format.line_spacing = 1.5

    def _setup_heading3(self):
        style = self.doc.styles["Heading 3"]
        self._set_font_times_new_roman(style)
        style.font.size = Pt(self.cfg.font_size_h3)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.space_before = Pt(6)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        style.paragraph_format.first_line_indent = Cm(0)
        style.paragraph_format.line_spacing = 1.5

    def _setup_toc_styles(self):
        """Setup TOC styles: TOC 1 = Bold, TOC 2/3 = Normal."""
        # TOC 1 (BAB entries) - Bold
        try:
            style = self.doc.styles["TOC 1"]
            self._set_font_times_new_roman(style)
            style.font.size = Pt(12)
            style.font.bold = True
            style.font.color.rgb = RGBColor(0, 0, 0)
            style.paragraph_format.line_spacing = 1.5
        except KeyError:
            pass

        # TOC 2 (sub-bab entries) - Normal
        try:
            style = self.doc.styles["TOC 2"]
            self._set_font_times_new_roman(style)
            style.font.size = Pt(12)
            style.font.bold = False
            style.font.color.rgb = RGBColor(0, 0, 0)
            style.paragraph_format.line_spacing = 1.5
            style.paragraph_format.left_indent = Cm(1.0)
        except KeyError:
            pass

        # TOC 3 (sub-sub-bab entries) - Normal
        try:
            style = self.doc.styles["TOC 3"]
            self._set_font_times_new_roman(style)
            style.font.size = Pt(12)
            style.font.bold = False
            style.font.color.rgb = RGBColor(0, 0, 0)
            style.paragraph_format.line_spacing = 1.5
            style.paragraph_format.left_indent = Cm(2.0)
        except KeyError:
            pass

    def setup_default_paragraph(self, paragraph):
        """Apply default body formatting to a paragraph."""
        paragraph.paragraph_format.line_spacing = self.cfg.line_spacing
        paragraph.paragraph_format.first_line_indent = Cm(self.cfg.first_line_indent)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in paragraph.runs:
            run.font.name = self.cfg.font_name
            run.font.size = Pt(self.cfg.font_size)
            run.font.color.rgb = RGBColor(0, 0, 0)
            # Set East Asian font for run
            rPr = run._r.get_or_add_rPr()
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is None:
                rFonts = OxmlElement("w:rFonts")
                rPr.insert(0, rFonts)
            rFonts.set(qn("w:ascii"), "Times New Roman")
            rFonts.set(qn("w:hAnsi"), "Times New Roman")
            rFonts.set(qn("w:eastAsia"), "Times New Roman")
            rFonts.set(qn("w:cs"), "Times New Roman")
