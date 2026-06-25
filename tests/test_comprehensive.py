"""
Comprehensive dummy testing for docx-idn
Tests all document types with realistic Indonesian academic content
"""
import pytest
from docx_idn import Makalah, Laporan, KTI, Skripsi, DocumentConfig, DocxDocument


# ============================================================
# DUMMY DATA - Realistic Indonesian Academic Content
# ============================================================

DUMMY_BAB_SEDERHANA = [
    {
        "judul": "PENDAHULUAN",
        "sections": [
            {
                "judul": "Latar Belakang",
                "isi": [
                    "Indonesia merupakan negara dengan jumlah penduduk terbesar keempat di dunia. "
                    "Dengan populasi lebih dari 270 juta jiwa, tantangan di bidang pendidikan, "
                    "kesehatan, dan ekonomi menjadi sangat kompleks dan memerlukan solusi inovatif.",
                    
                    "Teknologi kecerdasan buatan (Artificial Intelligence/AI) telah mengalami "
                    "perkembangan yang sangat pesat dalam dekade terakhir. AI tidak lagi menjadi "
                    "sekadar konsep fiksi ilmiah, melainkan telah menjadi bagian dari kehidupan "
                    "sehari-hari masyarakat modern.",
                    
                    "Berdasarkan data dari Kementerian Pendidikan dan Kebudayaan, tingkat literasi "
                    "digital masyarakat Indonesia masih berada pada angka 53,7% (Kemdikbud, 2024). "
                    "Angka ini menunjukkan bahwa masih banyak masyarakat yang belum melek teknologi, "
                    "padahal transformasi digital sudah di depan mata."
                ]
            },
            {
                "judul": "Rumusan Masalah",
                "isi": [
                    "Berdasarkan latar belakang yang telah diuraikan, maka rumusan masalah dalam "
                    "penelitian ini adalah:",
                    "1. Bagaimana pengaruh teknologi AI terhadap efektivitas pembelajaran di "
                    "jenjang pendidikan menengah?",
                    "2. Faktor-faktor apa saja yang mempengaruhi tingkat adopsi AI dalam "
                    "sistem pendidikan Indonesia?",
                    "3. Bagaimana strategi implementasi AI yang tepat untuk meningkatkan "
                    "kualitas pendidikan di Indonesia?"
                ]
            },
            {
                "judul": "Tujuan Penelitian",
                "isi": [
                    "Penelitian ini bertujuan untuk:",
                    "1. Menganalisis pengaruh teknologi AI terhadap efektivitas pembelajaran",
                    "2. Mengidentifikasi faktor-faktor yang mempengaruhi adopsi AI dalam "
                    "sistem pendidikan",
                    "3. Merumuskan strategi implementasi AI yang sesuai dengan kondisi "
                    "pendidikan Indonesia"
                ]
            }
        ]
    },
    {
        "judul": "TINJAUAN PUSTAKA",
        "sections": [
            {
                "judul": "Kecerdasan Buatan",
                "isi": [
                    "Kecerdasan buatan atau artificial intelligence (AI) adalah bidang ilmu "
                    "komputer yang berfokus pada penciptaan sistem yang mampu melakukan tugas "
                    "yang biasanya memerlukan kecerdasan manusia (Russell & Norvig, 2021). "
                    "AI meliputi berbagai sub-bidang seperti machine learning, natural language "
                    "processing, dan computer vision."
                ],
                "subsections": [
                    {
                        "judul": "Machine Learning",
                        "isi": [
                            "Machine learning adalah subset dari AI yang memungkinkan sistem "
                            "untuk belajar dari data tanpa diprogram secara eksplisit. "
                            "Algoritma machine learning dapat mengenali pola dalam data dan "
                            "membuat prediksi atau keputusan berdasarkan pola tersebut."
                        ]
                    },
                    {
                        "judul": "Natural Language Processing",
                        "isi": [
                            "Natural Language Processing (NLP) adalah cabang AI yang berfokus "
                            "pada interaksi antara komputer dan bahasa manusia. NLP memungkinkan "
                            "komputer untuk memahami, menafsirkan, dan menghasilkan teks dalam "
                            "bahasa manusia."
                        ]
                    }
                ]
            },
            {
                "judul": "Pendidikan di Indonesia",
                "isi": [
                    "Sistem pendidikan Indonesia mengalami berbagai tantangan, mulai dari "
                    "keterbatasan infrastruktur hingga kesenjangan kualitas pendidikan antara "
                    "daerah perkotaan dan pedesaan. Menurut data UNESCO (2023), Indonesia "
                    "masih berada di peringkat 69 dari 134 negara dalam hal kualitas pendidikan."
                ]
            }
        ]
    },
    {
        "judul": "METODOLOGI PENELITIAN",
        "sections": [
            {
                "judul": "Desain Penelitian",
                "isi": [
                    "Penelitian ini menggunakan pendekatan kuantitatif dengan desain "
                    "survey research. Sampel penelitian terdiri dari 500 guru dan 2000 "
                    "siswa dari 10 provinsi di Indonesia yang dipilih secara stratified "
                    "random sampling."
                ]
            },
            {
                "judul": "Instrumen Penelitian",
                "isi": [
                    "Instrumen yang digunakan dalam penelitian ini berupa kuesioner "
                    "dengan skala Likert 5 poin. Validitas instrumen diuji menggunakan "
                    "validitas konstruk dengan Confirmatory Factor Analysis (CFA), "
                    "sedangkan reliabilitas diuji menggunakan Cronbach's Alpha."
                ]
            }
        ]
    },
    {
        "judul": "HASIL DAN PEMBAHASAN",
        "sections": [
            {
                "judul": "Hasil Penelitian",
                "isi": [
                    "Berdasarkan analisis data yang telah dilakukan, diperoleh hasil "
                    "bahwa pengaruh teknologi AI terhadap efektivitas pembelajaran "
                    "menunjukkan hasil yang signifikan secara statistik (p < 0,05). "
                    "Nilai R² sebesar 0,67 menunjukkan bahwa 67% variasi dalam "
                    "efektivitas pembelajaran dapat dijelaskan oleh variabel AI."
                ]
            },
            {
                "judul": "Pembahasan",
                "isi": [
                    "Hasil penelitian ini konsisten dengan penelitian terdahulu yang "
                    "dilakukan oleh Smith et al. (2022) yang menemukan bahwa implementasi "
                    "AI dalam pendidikan dapat meningkatkan hasil belajar siswa hingga "
                    "30%. Namun, terdapat perbedaan signifikan dalam tingkat adopsi "
                    "antara wilayah perkotaan dan pedesaan."
                ]
            }
        ]
    },
    {
        "judul": "PENUTUP",
        "sections": [
            {
                "judul": "Kesimpulan",
                "isi": [
                    "Penelitian ini menyimpulkan bahwa teknologi AI memiliki pengaruh "
                    "positif yang signifikan terhadap efektivitas pembelajaran di "
                    "jenjang pendidikan menengah. Faktor utama yang mempengaruhi "
                    "tingkat adopsi AI adalah ketersediaan infrastruktur teknologi "
                    "dan tingkat literasi digital guru."
                ]
            },
            {
                "judul": "Saran",
                "isi": [
                    "Berdasarkan temuan penelitian ini, penulis memberikan saran sebagai "
                    "berikut:",
                    "1. Pemerintah perlu meningkatkan infrastruktur teknologi di daerah "
                    "terpencil",
                    "2. Program pelatihan AI untuk guru perlu diperluas dan "
                    "diperdalam",
                    "3. Penelitian selanjutnya perlu mengembangkan model implementasi "
                    "AI yang kontekstual untuk berbagai daerah di Indonesia"
                ]
            }
        ]
    }
]

DUMMY_DAFTAR_PUSTAKA = [
    {
        "penulis": "Russell, S. J., & Norvig, P.",
        "judul": "Artificial Intelligence: A Modern Approach",
        "tahun": "2021",
        "edisi": "4",
        "penerbit": "Pearson"
    },
    {
        "penulis": "Sugiyono",
        "judul": "Metode Penelitian Kuantitatif, Kualitatif, dan R&D",
        "tahun": "2019",
        "edisi": "",
        "penerbit": "Alfabeta"
    },
    {
        "penulis": "Smith, J., Johnson, A., & Williams, R.",
        "judul": "The Impact of AI on Education: A Global Perspective",
        "tahun": "2022",
        "edisi": "",
        "penerbit": "Cambridge University Press"
    },
    {
        "penulis": "Kementerian Pendidikan dan Kebudayaan",
        "judul": "Laporan Statistik Pendidikan Nasional 2024",
        "tahun": "2024",
        "edisi": "",
        "penerbit": "Kemdikbud",
        "url": "https://statistik.kemdikbud.go.id"
    },
    {
        "penulis": "Brown, T.",
        "judul": "Deep Learning for Natural Language Processing",
        "tahun": "2023",
        "edisi": "2",
        "penerbit": "MIT Press"
    }
]

DUMMY_LAMPIRAN = [
    {
        "judul": "Kuesioner Penelitian",
        "isi": [
            "Berikut adalah kuesioner yang digunakan dalam penelitian ini:",
            "Pertanyaan 1: Seberapa sering Anda menggunakan teknologi dalam proses "
            "pembelajaran? (1 = Tidak Pernah, 5 = Selalu)",
            "Pertanyaan 2: Menurut Anda, seberapa efektif penggunaan AI dalam "
            "meningkatkan pemahaman siswa?"
        ]
    },
    {
        "judul": "Data Statistik",
        "isi": [
            "Tabel 1: Distribusi Responden Berdasarkan Provinsi",
            "Tabel 2: Hasil Uji Validitas Instrumen",
            "Tabel 3: Ringkasan Statistik Deskriptif"
        ]
    }
]


# ============================================================
# TEST CASES
# ============================================================

class TestMakalah:
    """Test tipe dokumen Makalah"""
    
    def test_makalah_minimal(self, tmp_path):
        """Test makalah dengan data minimal"""
        doc = Makalah(
            judul="Dampak AI terhadap Pendidikan",
            penulis="Ahmad Fauzi",
            institusi="Universitas Indonesia",
            bab=[
                {
                    "judul": "PENDAHULUAN",
                    "sections": [
                        {"judul": "Latar Belakang", "isi": ["Isi sederhana"]}
                    ]
                }
            ]
        )
        out = tmp_path / "makalah_minimal.docx"
        doc.save(str(out))
        assert out.exists()
        assert out.stat().st_size > 0
    
    def test_makalah_lengkap(self, tmp_path):
        """Test makalah dengan semua fitur"""
        doc = Makalah(
            judul="Dampak Teknologi Kecerdasan Buatan terhadap Efektivitas Pembelajaran",
            penulis="Ahmad Fauzi",
            nim="123456789",
            institusi="Universitas Indonesia",
            fakultas="Ilmu Komputer",
            prodi="Teknik Informatika",
            angkatan="2024",
            tanggal="Juni 2026",
            bab=DUMMY_BAB_SEDERHANA,
            daftar_pustaka=DUMMY_DAFTAR_PUSTAKA
        )
        out = tmp_path / "makalah_lengkap.docx"
        doc.save(str(out))
        assert out.exists()
        assert out.stat().st_size > 10000  # Should be substantial
    
    def test_makalah_long_content(self, tmp_path):
        """Test makalah dengan konten panjang"""
        long_bab = [
            {
                "judul": "BAB PANJANG",
                "sections": [
                    {
                        "judul": "Section dengan Konten Sangat Panjang",
                        "isi": ["Paragraf ke-" + str(i) + " " + "kata " * 50 for i in range(20)]
                    }
                ]
            }
        ]
        doc = Makalah(
            judul="Tes Konten Panjang",
            penulis="Tester",
            institusi="Universitas Tes",
            bab=long_bab
        )
        out = tmp_path / "makalah_long.docx"
        doc.save(str(out))
        assert out.exists()


class TestLaporan:
    """Test tipe dokumen Laporan"""
    
    def test_laporan_generates(self, tmp_path):
        """Test laporan basic generation"""
        doc = Laporan(
            judul="Laporan Penelitian AI",
            penulis="Budi Santoso",
            institusi="Institut Teknologi Bandung",
            bab=[
                {
                    "judul": "BAB I",
                    "sections": [
                        {"judul": "Section 1", "isi": ["Content"]}
                    ]
                }
            ]
        )
        out = tmp_path / "laporan.docx"
        doc.save(str(out))
        assert out.exists()
    
    def test_laporan_with_lampiran(self, tmp_path):
        """Test laporan dengan lampiran"""
        doc = Laporan(
            judul="Laporan Lengkap",
            penulis="Test User",
            institusi="Universitas Test",
            bab=DUMMY_BAB_SEDERHANA[:2],  # First 2 chapters only
            lampiran=DUMMY_LAMPIRAN
        )
        out = tmp_path / "laporan_lampiran.docx"
        doc.save(str(out))
        assert out.exists()


class TestKTI:
    """Test tipe dokumen KTI (Karya Tulis Ilmiah)"""
    
    def test_kti_has_english_abstract(self, tmp_path):
        """Test KTI harus punya abstrak Inggris"""
        config = DocumentConfig.kti()
        assert config.has_abstrak_en is True
        
        doc = KTI(
            judul="Analisis Sentimen pada Ulasan Aplikasi E-Learning",
            judul_en="Sentiment Analysis on E-Learning App Reviews",
            penulis="Siti Rahma",
            nim="987654321",
            institusi="Universitas Gadjah Mada",
            abstrak_id="Penelitian ini menganalisis sentimen ulasan aplikasi e-learning.",
            abstrak_en="This study analyzes sentiment in e-learning app reviews.",
            bab=[
                {
                    "judul": "PENDAHULUAN",
                    "sections": [
                        {"judul": "Latar Belakang", "isi": ["Content"]}
                    ]
                }
            ]
        )
        out = tmp_path / "kti.docx"
        doc.save(str(out))
        assert out.exists()


class TestSkripsi:
    """Test tipe dokumen Skripsi"""
    
    def test_skripsi_has_all_features(self, tmp_path):
        """Test skripsi dengan semua fitur"""
        config = DocumentConfig.skripsi()
        assert config.has_abstrak_en is True
        assert config.has_lampiran is True
        
        doc = Skripsi(
            judul="Pengembangan Sistem Informasi Akademik Berbasis Web",
            judul_en="Development of Web-Based Academic Information System",
            penulis="Dewi Lestari",
            nim="2024001",
            institusi="Universitas Airlangga",
            fakultas="Ilmu Komputer",
            prodi="Sistem Informasi",
            angkatan="2024",
            pembimbing="Dr. Hendra Wijaya, M.Kom.",
            tanggal="Juli 2026",
            abstrak_id="Skripsi ini membahas pengembangan sistem informasi akademik.",
            abstrak_en="This thesis discusses the development of academic information system.",
            bab=DUMMY_BAB_SEDERHANA,
            daftar_pustaka=DUMMY_DAFTAR_PUSTAKA,
            lampiran=DUMMY_LAMPIRAN
        )
        out = tmp_path / "skripsi.docx"
        doc.save(str(out))
        assert out.exists()
        assert out.stat().st_size > 20000  # Skripsi should be large


class TestEdgeCases:
    """Test edge cases and error handling"""
    
    def test_bab_without_judul_raises(self):
        """Test that bab without judul raises error"""
        with pytest.raises(ValueError, match="tidak punya key 'judul'"):
            Makalah(
                judul="Test",
                penulis="Test",
                institusi="Test",
                bab=[{"sections": []}]  # Missing judul
            )
    
    def test_section_without_judul_raises(self):
        """Test that section without judul raises error"""
        with pytest.raises(ValueError, match="tidak punya key 'judul'"):
            Makalah(
                judul="Test",
                penulis="Test",
                institusi="Test",
                bab=[{"judul": "BAB", "sections": [{"isi": ["text"]}]}]
            )
    
    def test_empty_bab_list(self, tmp_path):
        """Test with empty bab list"""
        doc = Makalah(
            judul="Test Empty",
            penulis="Test",
            institusi="Test",
            bab=[]
        )
        out = tmp_path / "empty_bab.docx"
        doc.save(str(out))
        assert out.exists()
    
    def test_special_characters(self, tmp_path):
        """Test dengan karakter spesial"""
        doc = Makalah(
            judul="Analisis & Perbandingan <Metode> 'Canggih' \"AI\"",
            penulis="Nama dengan \"tanda kutip\"",
            institusi="Universitas @#$%",
            bab=[
                {
                    "judul": "BAB & <>",
                    "sections": [
                        {"judul": "Section \"1\"", "isi": ["Konten dengan karakter spesial: & < > \" '"]}
                    ]
                }
            ]
        )
        out = tmp_path / "special_chars.docx"
        doc.save(str(out))
        assert out.exists()
    
    def test_unicode_indonesian(self, tmp_path):
        """Test karakter unicode Indonesia"""
        doc = Makalah(
            judul="Analisis Huruf: àáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ",
            penulis="Penulis",
            institusi="Universitas",
            bab=[
                {
                    "judul": "Bab",
                    "sections": [
                        {"judul": "Section", "isi": ["Teks dengan unicode: é ñ ü ö ä"]}
                    ]
                }
            ]
        )
        out = tmp_path / "unicode.docx"
        doc.save(str(out))
        assert out.exists()
    
    def test_very_long_judul(self, tmp_path):
        """Test judul sangat panjang"""
        long_title = "Judul yang Sangat Panjang " * 20
        doc = Makalah(
            judul=long_title,
            penulis="Test",
            institusi="Test",
            bab=[{"judul": "BAB", "sections": [{"judul": "S", "isi": ["X"]}]}]
        )
        out = tmp_path / "long_title.docx"
        doc.save(str(out))
        assert out.exists()
    
    def test_multiple_page_breaks(self, tmp_path):
        """Test banyak bab untuk multiple page breaks"""
        many_babs = [
            {
                "judul": f"BAB {i}",
                "sections": [
                    {"judul": f"Section {i}.1", "isi": [f"Content for chapter {i}"]}
                ]
            }
            for i in range(1, 10)
        ]
        doc = Makalah(
            judul="Banyak Bab",
            penulis="Test",
            institusi="Test",
            bab=many_babs
        )
        out = tmp_path / "many_babs.docx"
        doc.save(str(out))
        assert out.exists()


class TestDocumentConfig:
    """Test DocumentConfig presets"""
    
    def test_makalah_config(self):
        cfg = DocumentConfig.makalah()
        assert cfg.has_cover is False
        assert cfg.has_abstrak_en is False
        assert cfg.has_lampiran is False
    
    def test_laporan_config(self):
        cfg = DocumentConfig.laporan()
        assert cfg.has_cover is False
        assert cfg.has_abstrak_en is False
        assert cfg.has_lampiran is True
    
    def test_kti_config(self):
        cfg = DocumentConfig.kti()
        assert cfg.has_cover is False
        assert cfg.has_abstrak_en is True
        assert cfg.has_lampiran is True
    
    def test_skripsi_config(self):
        cfg = DocumentConfig.skripsi()
        assert cfg.has_cover is False
        assert cfg.has_abstrak_en is True
        assert cfg.has_lampiran is True
    
    def test_custom_config(self):
        cfg = DocumentConfig(
            margin_left=3.0,
            margin_right=3.0,
            font_size=11,
            line_spacing=2.0
        )
        assert cfg.margin_left == 3.0
        assert cfg.font_size == 11
        assert cfg.line_spacing == 2.0


class TestDocxDocument:
    """Test DocxDocument directly"""
    
    def test_add_paragraph(self, tmp_path):
        doc = DocxDocument()
        p = doc.add_paragraph("Test paragraph")
        assert p is not None
        out = tmp_path / "direct.docx"
        doc.save(str(out))
        assert out.exists()
    
    def test_add_heading(self, tmp_path):
        doc = DocxDocument()
        h = doc.add_heading("Test Heading", level=1)
        assert h is not None
        out = tmp_path / "heading.docx"
        doc.save(str(out))
        assert out.exists()
    
    def test_multiple_headings(self, tmp_path):
        doc = DocxDocument()
        doc.add_heading("Heading 1", level=1)
        doc.add_heading("Heading 2", level=2)
        doc.add_heading("Heading 3", level=3)
        out = tmp_path / "headings.docx"
        doc.save(str(out))
        assert out.exists()


class TestBibliographyFormats:
    """Test various bibliography entry formats"""
    
    def test_book_reference(self, tmp_path):
        doc = Makalah(
            judul="Test",
            penulis="Test",
            institusi="Test",
            bab=[{"judul": "BAB", "sections": [{"judul": "S", "isi": ["X"]}]}],
            daftar_pustaka=[{
                "penulis": "Sugiyono",
                "judul": "Metode Penelitian",
                "tahun": "2019",
                "penerbit": "Alfabeta"
            }]
        )
        out = tmp_path / "bib_book.docx"
        doc.save(str(out))
        assert out.exists()
    
    def test_online_reference(self, tmp_path):
        doc = Makalah(
            judul="Test",
            penulis="Test",
            institusi="Test",
            bab=[{"judul": "BAB", "sections": [{"judul": "S", "isi": ["X"]}]}],
            daftar_pustaka=[{
                "penulis": "WHO",
                "judul": "Health Statistics",
                "tahun": "2024",
                "url": "https://who.int/data"
            }]
        )
        out = tmp_path / "bib_online.docx"
        doc.save(str(out))
        assert out.exists()
    
    def test_edition_reference(self, tmp_path):
        doc = Makalah(
            judul="Test",
            penulis="Test",
            institusi="Test",
            bab=[{"judul": "BAB", "sections": [{"judul": "S", "isi": ["X"]}]}],
            daftar_pustaka=[{
                "penulis": "Cormen et al.",
                "judul": "Introduction to Algorithms",
                "tahun": "2022",
                "edisi": "4",
                "penerbit": "MIT Press"
            }]
        )
        out = tmp_path / "bib_edition.docx"
        doc.save(str(out))
        assert out.exists()


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
