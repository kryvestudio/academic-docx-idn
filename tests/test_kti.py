from docx_idn.types.kti import KTI


def test_kti_generates():
    k = KTI(
        judul="KTI: Analisis Sentimen",
        penulis="Siti Aminah",
        institusi="Institut Teknologi Bandung",
        fakultas="Matematika dan Ilmu Pengetahuan Alam",
        judul_en="Sentiment Analysis KTI",
        abstrak_id="Penelitian ini menganalisis...",
        abstrak_en="This research analyzes...",
        bab=[
            {"judul": "Pendahuluan", "sections": [{"judul": "Latar Belakang", "isi": ["Sentimen..."]}]},
        ],
        daftar_pustaka=[
            {"penulis": "Manning", "judul": "Foundations of NLP", "tahun": "2021", "penerbit": "Cambridge"}
        ],
    )
    k.save("test_output_kti.docx")
    assert True
