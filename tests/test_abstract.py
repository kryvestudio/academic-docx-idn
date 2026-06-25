from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.abstract import Abstract


def test_abstract_id_only():
    doc = DocxDocument(DocumentConfig.makalah())
    abstract = Abstract(abstrak_id="Ini abstrak Indonesia.")
    abstract.render(doc)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("ABSTRAK" in t for t in texts)
    assert any("Ini abstrak Indonesia" in t for t in texts)


def test_abstract_both():
    doc = DocxDocument(DocumentConfig.kti())
    abstract = Abstract(
        abstrak_id="Abstrak Indonesia.",
        abstrak_en="English abstract.",
    )
    abstract.render(doc)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("ABSTRAK" in t for t in texts)
    assert any("ABSTRACT" in t for t in texts)
