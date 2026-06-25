from docx_idn.types.makalah import Makalah


def test_makalah_generates():
    m = Makalah(
        judul="Dampak AI terhadap Pendidikan",
        penulis="Ahmad Fauzi",
        institusi="Universitas Indonesia",
        fakultas="Ilmu Komputer",
        bab=[
            {
                "judul": "Pendahuluan",
                "sections": [
                    {"judul": "Latar Belakang", "isi": ["AI berkembang pesat..."]},
                ]
            }
        ],
        daftar_pustaka=[
            {"penulis": "Sugiyono", "judul": "Metode Penelitian", "tahun": "2019", "penerbit": "Alfabeta"}
        ],
    )
    m.save("test_output_makalah.docx")
    assert True
