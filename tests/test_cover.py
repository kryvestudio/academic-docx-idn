from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.cover import CoverPage


def test_cover_renders():
    doc = DocxDocument(DocumentConfig())
    cover = CoverPage(
        judul="Dampak AI terhadap Pendidikan",
        penulis="Ahmad Fauzi",
        nim="123456789",
        institusi="Universitas Indonesia",
        fakultas="Ilmu Komputer",
        prodi="Teknik Informatika",
        angkatan="2024",
        tanggal="Juni 2026",
    )
    cover.render(doc)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("UNIVERSITAS INDONESIA" in t for t in texts)
    assert any("DAMPAK AI" in t for t in texts)
    assert any("Ahmad Fauzi" in t for t in texts)
