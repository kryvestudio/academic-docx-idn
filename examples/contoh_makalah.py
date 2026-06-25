"""Contoh pembuatan Makalah."""
from docx_idn import Makalah

doc = Makalah(
    judul="Dampak Kecerdasan Buatan terhadap Dunia Pendidikan di Indonesia",
    penulis="Ahmad Fauzi",
    nim="123456789",
    institusi="Universitas Indonesia",
    fakultas="Ilmu Komputer",
    prodi="Teknik Informatika",
    angkatan="2024",
    tanggal="Juni 2026",
    kata_pengantar="Puji syukur kami panjatkan ke hadirat Tuhan Yang Maha Esa...",
    abstrak_id="Penelitian ini membahas dampak kecerdasan buatan (AI) terhadap "
               "dunia pendidikan di Indonesia. Hasil penelitian menunjukkan bahwa "
               "AI memiliki potensi besar untuk meningkatkan kualitas pembelajaran.",
    bab=[
        {
            "judul": "Pendahuluan",
            "sections": [
                {
                    "judul": "Latar Belakang",
                    "isi": [
                        "Kecerdasan buatan (Artificial Intelligence/AI) telah menjadi "
                        "salah satu teknologi yang paling berpengaruh di era digital. "
                        "Dalam konteks pendidikan, AI menawarkan berbagai kemungkinan "
                        "baru untuk meningkatkan kualitas pembelajaran.",
                        "Indonesia sebagai negara berkembang memiliki tantangan "
                        "tersendiri dalam mengadopsi teknologi AI di sektor pendidikan. "
                        "Kesenjangan digital dan infrastruktur yang belum merata menjadi "
                        "hambatan utama.",
                    ],
                },
                {
                    "judul": "Rumusan Masalah",
                    "isi": [
                        "1. Bagaimana dampak kecerdasan buatan terhadap kualitas "
                        "pendidikan di Indonesia?",
                        "2. Apa saja tantangan dan peluang implementasi AI dalam "
                        "dunia pendidikan?",
                    ],
                },
                {
                    "judul": "Tujuan Penelitian",
                    "isi": [
                        "Penelitian ini bertujuan untuk menganalisis dampak kecerdasan "
                        "buatan terhadap kualitas pendidikan di Indonesia."
                    ],
                },
            ],
        },
        {
            "judul": "Tinjauan Pustaka",
            "sections": [
                {
                    "judul": "Pengertian Kecerdasan Buatan",
                    "isi": [
                        "Kecerdasan buatan (AI) adalah bidang ilmu komputer yang "
                        "berfokus pada pembuatan sistem yang mampu melakukan tugas yang "
                        "biasanya memerlukan kecerdasan manusia.",
                    ],
                },
                {
                    "judul": "AI dalam Pendidikan",
                    "subsections": [
                        {
                            "judul": "Sistem Tutor Cerdas",
                            "isi": [
                                "Sistem tutor cerdas (Intelligent Tutoring System/ITS) "
                                "adalah sistem komputer yang menyesuaikan instruksi "
                                "berdasarkan kebutuhan siswa."
                            ],
                        },
                        {
                            "judul": "Analitik Pembelajaran",
                            "isi": [
                                "Analitik pembelajaran menggunakan data untuk memahami "
                                "dan mengoptimalkan proses belajar."
                            ],
                        },
                    ],
                    "isi": [
                        "AI telah diterapkan dalam berbagai aspek pendidikan.",
                    ],
                },
            ],
        },
        {
            "judul": "Metodologi Penelitian",
            "sections": [
                {
                    "judul": "Jenis Penelitian",
                    "isi": [
                        "Penelitian ini menggunakan pendekatan kualitatif dengan "
                        "metode studi literatur."
                    ],
                },
                {
                    "judul": "Sumber Data",
                    "isi": [
                        "Data dikumpulkan dari jurnal ilmiah, artikel konferensi, "
                        "dan laporan industri."
                    ],
                },
            ],
        },
        {
            "judul": "Hasil dan Pembahasan",
            "sections": [
                {
                    "judul": "Temuan Utama",
                    "isi": [
                        "Berdasarkan analisis yang dilakukan, ditemukan bahwa AI "
                        "memiliki dampak positif signifikan terhadap kualitas "
                        "pendidikan, terutama dalam aspek personalisasi pembelajaran."
                    ],
                },
            ],
        },
        {
            "judul": "Penutup",
            "sections": [
                {
                    "judul": "Kesimpulan",
                    "isi": [
                        "AI memiliki potensi besar untuk mentransformasi dunia "
                        "pendidikan di Indonesia, namun diperlukan kebijakan yang "
                        "tepat untuk mengatasi tantangan yang ada."
                    ],
                },
                {
                    "judul": "Saran",
                    "isi": [
                        "1. Pemerintah perlu mempercepat infrastruktur digital.",
                        "2. Guru perlu dilatih untuk mengintegrasikan AI dalam "
                        "pembelajaran.",
                    ],
                },
            ],
        },
    ],
    daftar_pustaka=[
        {
            "penulis": "Sugiyono",
            "judul": "Metode Penelitian Kuantitatif, Kualitatif, dan R&D",
            "tahun": "2019",
            "penerbit": "Bandung: Alfabeta",
        },
        {
            "penulis": "Russell, S.J. dan Norvig, P.",
            "judul": "Artificial Intelligence: A Modern Approach",
            "edisi": "4",
            "tahun": "2021",
            "penerbit": "Pearson",
        },
    ],
)

doc.save("output_v2.docx")
print("Makalah berhasil dibuat: output_v2.docx")
