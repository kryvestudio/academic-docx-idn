"""Contoh pembuatan KTI (Karya Tulis Ilmiah)."""
from docx_idn import KTI

doc = KTI(
    judul="Analisis Sentimen Review Aplikasi Mobile Banking di Indonesia",
    penulis="Siti Aminah",
    nim="1122334455",
    institusi="Institut Teknologi Bandung",
    fakultas="Matematika dan Ilmu Pengetahuan Alam",
    prodi="Informatika",
    angkatan="2022",
    tanggal="Juni 2026",
    judul_en="Sentiment Analysis of Mobile Banking App Reviews in Indonesia",
    abstrak_id="Penelitian ini menganalisis sentimen pengguna terhadap aplikasi "
               "mobile banking di Indonesia menggunakan teknik Natural Language "
               "Processing (NLP). Data dikumpulkan dari 10.000 review di Google "
               "Play Store. Hasil menunjukkan 65% sentimen positif, 20% negatif, "
               "dan 15% netral.",
    abstrak_en="This study analyzes user sentiment toward mobile banking "
               "applications in Indonesia using Natural Language Processing (NLP) "
               "techniques. Data was collected from 10,000 reviews on Google Play "
               "Store. Results show 65% positive sentiment, 20% negative, and "
               "15% neutral.",
    bab=[
        {
            "judul": "Pendahuluan",
            "sections": [
                {
                    "judul": "Latar Belakang",
                    "isi": [
                        "Mobile banking telah mengalami pertumbuhan pesat di "
                        "Indonesia dalam beberapa tahun terakhir. Dengan jumlah "
                        "pengguna internet yang terus meningkat, mobile banking "
                        "menjadi salah satu layanan perbankan digital yang paling "
                        "banyak digunakan.",
                        "Analisis sentimen merupakan teknik NLP yang digunakan "
                        "untuk mengidentifikasi dan mengkuantifikasi subjek "
                        "dalam teks. Dalam konteks mobile banking, analisis "
                        "sentimen dapat membantu bank memahami kebutuhan dan "
                        "keluhan pelanggan."
                    ],
                },
                {
                    "judul": "Rumusan Masalah",
                    "isi": [
                        "1. Bagaimana distribusi sentimen review aplikasi mobile "
                        "banking di Indonesia?",
                        "2. Aspek apa saja yang paling banyak mendapat sentimen "
                        "positif dan negatif?"
                    ],
                },
                {
                    "judul": "Tujuan Penelitian",
                    "isi": [
                        "Menganalisis sentimen pengguna terhadap aplikasi mobile "
                        "banking dan mengidentifikasi aspek-aspek yang perlu "
                        "ditingkatkan."
                    ],
                },
            ],
        },
        {
            "judul": "Tinjauan Pustaka",
            "sections": [
                {
                    "judul": "Natural Language Processing",
                    "isi": [
                        "NLP adalah bidang ilmu komputer yang berfokus pada "
                        "interaksi antara komputer dan bahasa manusia."
                    ],
                },
                {
                    "judul": "Analisis Sentimen",
                    "subsections": [
                        {
                            "judul": "Metode Lexicon-Based",
                            "isi": [
                                "Metode berbasis kamus yang menggunakan daftar "
                                "kata beserta bobot sentimennya."
                            ],
                        },
                        {
                            "judul": "Metode Machine Learning",
                            "isi": [
                                "Metode yang menggunakan model机器学习 untuk "
                                "klasifikasi sentimen."
                            ],
                        },
                    ],
                    "isi": [
                        "Analisis sentimen dapat dilakukan dengan berbagai metode."
                    ],
                },
                {
                    "judul": "Mobile Banking",
                    "isi": [
                        "Mobile banking adalah layanan perbankan yang dapat "
                        "diakses melalui perangkat mobile."
                    ],
                },
            ],
        },
        {
            "judul": "Metodologi",
            "sections": [
                {
                    "judul": "Desain Penelitian",
                    "isi": [
                        "Penelitian ini menggunakan pendekatan kuantitatif "
                        "dengan metode analisis sentimen."
                    ],
                },
                {
                    "judul": "Pengumpulan Data",
                    "isi": [
                        "Data dikumpulkan dari Google Play Store menggunakan "
                        "web scraping. Sampel berjumlah 10.000 review."
                    ],
                },
                {
                    "judul": "Analisis Data",
                    "isi": [
                        "Analisis dilakukan menggunakan Python dengan library "
                        "NLTK, Scikit-learn, dan Transformers."
                    ],
                },
            ],
        },
        {
            "judul": "Hasil dan Pembahasan",
            "sections": [
                {
                    "judul": "Distribusi Sentimen",
                    "isi": [
                        "Hasil analisis menunjukkan distribusi sentimen sebagai "
                        "berikut: positif 65%, negatif 20%, netral 15%."
                    ],
                },
                {
                    "judul": "Aspek Analisis",
                    "isi": [
                        "Aspek yang paling banyak mendapat sentimen positif "
                        "adalah kemudahan penggunaan (45%), sedangkan aspek "
                        "yang paling banyak mendapat sentimen negatif adalah "
                        "performa aplikasi (30%)."
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
                        "1. Mayoritas pengguna mobile banking di Indonesia "
                        "memberikan sentimen positif.",
                        "2. Aspek performa aplikasi menjadi perhatian utama "
                        "pengguna."
                    ],
                },
                {
                    "judul": "Saran",
                    "isi": [
                        "1. Pengembang perlu meningkatkan performa aplikasi.",
                        "2. Penelitian selanjutnya dapat menggunakan metode "
                        "deep learning untuk analisis yang lebih akurat."
                    ],
                },
            ],
        },
    ],
    daftar_pustaka=[
        {
            "penulis": "Manning, C.D. dan Schutze, H.",
            "judul": "Foundations of Statistical Natural Language Processing",
            "tahun": "2021",
            "penerbit": "MIT Press",
        },
        {
            "penulis": "Liu, B.",
            "judul": "Sentiment Analysis and Opinion Mining",
            "tahun": "2022",
            "penerbit": "Springer",
        },
    ],
    lampiran=[
        {
            "judul": "Lampiran A: Contoh Code Python",
            "isi": ["Kode Python untuk analisis sentimen."],
        },
    ],
)

doc.save("contoh_kti.docx")
print("KTI berhasil dibuat: contoh_kti.docx")
