from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.bibliography import Bibliography


def test_bibliography_renders():
    doc = DocxDocument(DocumentConfig())
    bib = Bibliography(
        entries=[
            {
                "penulis": "Sugiyono",
                "judul": "Metode Penelitian Kuantitatif",
                "tahun": "2019",
                "penerbit": "Bandung: Alfabeta",
            },
            {
                "penulis": "Moleong, L.J.",
                "judul": "Metodologi Penelitian Kualitatif",
                "edisi": "37",
                "tahun": "2018",
                "penerbit": "Bandung: Remaja Rosdakarya",
            },
        ]
    )
    bib.render(doc)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("DAFTAR PUSTAKA" in t for t in texts)
    assert any("Sugiyono" in t for t in texts)
    assert any("Moleong" in t for t in texts)
