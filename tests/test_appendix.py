from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.appendix import Appendix


def test_appendix_renders():
    doc = DocxDocument(DocumentConfig())
    appendix = Appendix(
        items=[
            {"judul": "Lampiran A: Kuesioner", "isi": ["Isi kuesioner..."]},
            {"judul": "Lampiran B: Data", "isi": ["Data tabel..."]},
        ]
    )
    appendix.render(doc)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("LAMPIRAN" in t for t in texts)
    assert any("Kuesioner" in t for t in texts)
