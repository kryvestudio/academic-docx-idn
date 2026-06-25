from docx_idn.types.laporan import Laporan


def test_laporan_generates():
    l = Laporan(
        judul="Laporan Observasi",
        penulis="Budi Santoso",
        institusi="Universitas Gadjah Mada",
        fakultas="Ekonomi",
        bab=[
            {"judul": "Pendahuluan", "sections": [{"judul": "Latar Belakang", "isi": ["Observasi..."]}]},
            {"judul": "Hasil Observasi", "sections": [{"judul": "Temuan", "isi": ["Data menunjukkan..."]}]},
        ],
    )
    l.save("test_output_laporan.docx")
    assert True
