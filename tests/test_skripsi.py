from docx_idn.types.skripsi import Skripsi


def test_skripsi_generates():
    s = Skripsi(
        judul="Skripsi: Pemodelan ML",
        penulis="Dewi Lestari",
        institusi="Universitas Padjadjaran",
        fakultas="Ilmu Komputer",
        judul_en="ML Modeling Thesis",
        abstrak_id="Skripsi ini membahas...",
        abstrak_en="This thesis discusses...",
        pembimbing="Dr. Andi Wijaya",
        bab=[
            {"judul": "Pendahuluan", "sections": [{"judul": "Latar Belakang", "isi": ["ML..."]}]},
        ],
    )
    s.save("test_output_skripsi.docx")
    assert True
