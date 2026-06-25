from docx_idn.types.base import BaseDocument
from docx_idn.core.config import DocumentConfig


def test_base_document_minimal():
    doc = BaseDocument(
        config=DocumentConfig.makalah(),
        judul="Test",
        penulis="Author",
        institusi="Universitas",
        fakultas="Fakultas",
        bab=[{"judul": "Bab 1", "sections": [{"judul": "S1", "isi": ["Content"]}]}],
    )
    doc.save("test_base_output.docx")
    assert True


def test_base_document_with_all_sections():
    doc = BaseDocument(
        config=DocumentConfig.kti(),
        judul="Test Lengkap",
        penulis="Author",
        institusi="Universitas",
        fakultas="Fakultas",
        judul_en="Test Complete",
        abstrak_id="Abstrak ID",
        abstrak_en="Abstract EN",
        kata_pengantar="Kata pengantar di sini.",
        bab=[
            {"judul": "Bab 1", "sections": [{"judul": "S1", "isi": ["Content"]}]},
        ],
        daftar_pustaka=[
            {"penulis": "Sugiyono", "judul": "Metode", "tahun": "2019", "penerbit": "Alfabeta"}
        ],
        lampiran=[{"judul": "Lampiran A", "isi": ["Isi lampiran"]}],
    )
    doc.save("test_base_output_lengkap.docx")
    assert True
