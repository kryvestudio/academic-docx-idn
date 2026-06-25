from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.chapter import Chapter


def test_chapter_renders():
    doc = DocxDocument(DocumentConfig())
    chapter = Chapter(
        judul="Pendahuluan",
        sections=[
            {
                "judul": "Latar Belakang",
                "isi": ["Ini paragraf latar belakang.", "Paragraf kedua."]
            },
            {
                "judul": "Rumusan Masalah",
                "isi": ["1. Apa dampak AI?"]
            }
        ]
    )
    chapter.render(doc, number=1)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("BAB I" in t for t in texts)
    assert any("PENDAHULUAN" in t for t in texts)
    assert any("Latar Belakang" in t for t in texts)


def test_chapter_with_subsections():
    doc = DocxDocument(DocumentConfig())
    chapter = Chapter(
        judul="Tinjauan Pustaka",
        sections=[
            {
                "judul": "Pengertian AI",
                "subsections": [
                    {
                        "judul": "Definisi Umum",
                        "isi": ["AI adalah..."]
                    }
                ],
                "isi": ["AI merupakan..."]
            }
        ]
    )
    chapter.render(doc, number=2)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("BAB II" in t for t in texts)
    assert any("Definisi Umum" in t for t in texts)
