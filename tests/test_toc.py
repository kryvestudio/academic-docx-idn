from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.toc import TableOfContents


def test_toc_renders():
    doc = DocxDocument(DocumentConfig())
    toc = TableOfContents()
    toc.render(doc)
    xml_str = doc.document._element.body.xml
    assert "DAFTAR ISI" in xml_str
