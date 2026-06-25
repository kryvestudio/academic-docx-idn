from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx_idn.core.document import _set_run_font


class TableOfContents:
    def __init__(self, title: str = "DAFTAR ISI", levels: str = "1-3"):
        self.title = title
        self.levels = levels

    def render(self, doc):
        # Create TOC field code using SDT (Structured Document Tag)
        # This is the proper way Word handles TOC
        sdt = OxmlElement("w:sdt")
        sdtPr = OxmlElement("w:sdtPr")
        docPartObj = OxmlElement("w:docPartObj")
        docPartGallery = OxmlElement("w:docPartGallery")
        docPartGallery.set(qn("w:val"), "Table of Contents")
        docPartUnique = OxmlElement("w:docPartUnique")
        docPartUnique.set(qn("w:val"), "true")
        docPartObj.append(docPartGallery)
        docPartObj.append(docPartUnique)
        sdtPr.append(docPartObj)
        sdt.append(sdtPr)

        sdtContent = OxmlElement("w:sdtContent")

        # TOC title paragraph
        p = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")
        pStyle = OxmlElement("w:pStyle")
        pStyle.set(qn("w:val"), "TOC Heading")
        pPr.append(pStyle)
        p.append(pPr)
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = self.title
        r.append(t)
        p.append(r)
        sdtContent.append(p)

        # TOC field paragraph
        p2 = OxmlElement("w:p")
        r2 = OxmlElement("w:r")

        # Field begin
        fldChar1 = OxmlElement("w:fldChar")
        fldChar1.set(qn("w:fldCharType"), "begin")
        r2.append(fldChar1)

        # Field instruction
        instrText = OxmlElement("w:instrText")
        instrText.set(qn("xml:space"), "preserve")
        instrText.text = f' TOC \\o "{self.levels}" \\h \\z \\u '
        r2.append(instrText)

        # Field separate
        fldChar2 = OxmlElement("w:fldChar")
        fldChar2.set(qn("w:fldCharType"), "separate")
        r2.append(fldChar2)

        # Placeholder text
        r3 = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), "Times New Roman")
        rFonts.set(qn("w:hAnsi"), "Times New Roman")
        rFonts.set(qn("w:eastAsia"), "Times New Roman")
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "24")
        rPr.append(rFonts)
        rPr.append(sz)
        r3.append(rPr)
        t2 = OxmlElement("w:t")
        t2.text = "[Daftar Isi akan muncul setelah Update Field]"
        r3.append(t2)
        r2.append(r3)

        # Field end
        fldChar3 = OxmlElement("w:fldChar")
        fldChar3.set(qn("w:fldCharType"), "end")
        r2.append(fldChar3)

        p2.append(r2)
        sdtContent.append(p2)

        sdt.append(sdtContent)

        # Insert SDT before section properties
        doc.document._element.body.insert_element_before(sdt, *("w:sectPr",))

        doc.add_page_break()
