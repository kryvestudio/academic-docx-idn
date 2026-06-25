"""Contoh pembuatan Skripsi."""
from docx_idn import Skripsi

doc = Skripsi(
    judul="Pemodelan Prediksi Harga Properti menggunakan Machine Learning",
    penulis="Dewi Lestari",
    nim="5566778899",
    institusi="Universitas Padjadjaran",
    fakultas="Ilmu Komputer",
    prodi="Informatika",
    angkatan="2021",
    tanggal="Juni 2026",
    pembimbing="Dr. Andi Wijaya, M.T.",
    judul_en="Property Price Prediction Modeling using Machine Learning",
    abstrak_id="Skripsi ini membahas pemodelan prediksi harga properti menggunakan "
               "teknik machine learning. Data yang digunakan adalah data transaksi "
               "properti di Kota Bandung selama tahun 2023-2025. Model terbaik "
               "yang dihasilkan adalah Random Forest dengan nilai R-squared 0.89.",
    abstrak_en="This thesis discusses property price prediction modeling using "
               "machine learning techniques. The data used is property transaction "
               "data in Bandung City during 2023-2025. The best model produced is "
               "Random Forest with R-squared value of 0.89.",
    bab=[
        {
            "judul": "Pendahuluan",
            "sections": [
                {
                    "judul": "Latar Belakang Masalah",
                    "isi": [
                        "Harga properti merupakan salah satu indikator penting "
                        "dalam perekonomian suatu daerah. Prediksi harga properti "
                        "yang akurat dapat membantu berbagai pihak dalam pengambilan "
                        "keputusan, baik investor, pengembang, maupun masyarakat "
                        "umum.",
                        "Machine learning menawarkan pendekatan yang lebih akurat "
                        "dibandingkan metode tradisional dalam memprediksi harga "
                        "properti. Berbagai algoritma seperti Linear Regression, "
                        "Random Forest, dan Gradient Boosting telah terbukti "
                        "efektif dalam tugas regresi."
                    ],
                },
                {
                    "judul": "Rumusan Masalah",
                    "isi": [
                        "1. Bagaimana performa berbagai algoritma machine learning "
                        "dalam memprediksi harga properti?",
                        "2. Faktor-faktor apa saja yang paling berpengaruh terhadap "
                        "harga properti?",
                        "3. Model mana yang memberikan prediksi paling akurat?"
                    ],
                },
                {
                    "judul": "Tujuan Penelitian",
                    "isi": [
                        "1. Membangun model prediksi harga properti menggunakan "
                        "machine learning.",
                        "2. Menganalisis faktor-faktor yang mempengaruhi harga "
                        "properti.",
                        "3. Membandingkan performa beberapa algoritma machine "
                        "learning."
                    ],
                },
                {
                    "judul": "Manfaat Penelitian",
                    "isi": [
                        "Manfaat teoritis: Menambah pengetahuan tentang penerapan "
                        "machine learning dalam prediksi harga properti.",
                        "Manfaat praktis: Dapat digunakan sebagai referensi oleh "
                        "pengembang properti dan investor dalam pengambilan "
                        "keputusan."
                    ],
                },
            ],
        },
        {
            "judul": "Tinjauan Pustaka",
            "sections": [
                {
                    "judul": "Harga Properti",
                    "isi": [
                        "Harga properti dipengaruhi oleh berbagai faktor "
                        "seperti lokasi, luas tanah, luas bangunan, jumlah "
                        "kamar, dan fasilitas pendukung."
                    ],
                },
                {
                    "judul": "Machine Learning untuk Regresi",
                    "subsections": [
                        {
                            "judul": "Linear Regression",
                            "isi": [
                                "Linear regression adalah metode statistik yang "
                                "digunakan untuk memodelkan hubungan linear antara "
                                "variabel dependen dan satu atau lebih variabel "
                                "independen."
                            ],
                        },
                        {
                            "judul": "Random Forest",
                            "isi": [
                                "Random forest adalah algoritma ensemble learning "
                                "yang menggunakan beberapa decision tree untuk "
                                "menghasilkan prediksi."
                            ],
                        },
                        {
                            "judul": "Gradient Boosting",
                            "isi": [
                                "Gradient boosting adalah teknik ensemble learning "
                                "yang membangun model secara sequential."
                            ],
                        },
                    ],
                    "isi": [
                        "Machine learning menawarkan pendekatan data-driven "
                        "untuk prediksi harga properti."
                    ],
                },
            ],
        },
        {
            "judul": "Metodologi Penelitian",
            "sections": [
                {
                    "judul": "Desain Penelitian",
                    "isi": [
                        "Penelitian ini menggunakan pendekatan kuantitatif "
                        "dengan metode eksperimen."
                    ],
                },
                {
                    "judul": "Data dan Sumber Data",
                    "isi": [
                        "Data transaksi properti dikumpulkan dari portal "
                        "properti online dan dinas terkait. Periode data "
                        "2023-2025 dengan total 5.000 transaksi."
                    ],
                },
                {
                    "judul": "Variabel Penelitian",
                    "isi": [
                        "Variabel dependen: Harga properti. "
                        "Variabel independen: Lokasi, luas tanah, luas bangunan, "
                        "jumlah kamar, jumlah kamar mandi, usia bangunan, "
                        "fasilitas."
                    ],
                },
                {
                    "judul": "Teknik Analisis",
                    "isi": [
                        "Analisis menggunakan Python dengan library pandas, "
                        "scikit-learn, dan XGBoost. Metrik evaluasi: "
                        "R-squared, MAE, RMSE."
                    ],
                },
            ],
        },
        {
            "judul": "Hasil Penelitian",
            "sections": [
                {
                    "judul": "Statistik Deskriptif",
                    "isi": [
                        "Rata-rata harga properti dalam sampel adalah "
                        "Rp 1.5 Miliar dengan standar deviasi Rp 800 Juta."
                    ],
                },
                {
                    "judul": "Hasil Model",
                    "isi": [
                        "Random Forest menghasilkan performa terbaik dengan "
                        "R-squared 0.89, MAE Rp 120 Juta, dan RMSE Rp 180 Juta."
                    ],
                },
            ],
        },
        {
            "judul": "Pembahasan",
            "sections": [
                {
                    "judul": "Analisis Hasil",
                    "isi": [
                        "Random Forest unggul karena mampu menangkap hubungan "
                        "non-linear antara fitur dan harga properti."
                    ],
                },
                {
                    "judul": "Faktor Penentu Harga",
                    "isi": [
                        "Faktor paling berpengaruh adalah lokasi (35%), "
                        "luas tanah (25%), dan luas bangunan (20%)."
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
                        "1. Random Forest memberikan prediksi paling akurat "
                        "untuk harga properti.",
                        "2. Lokasi merupakan faktor paling berpengaruh "
                        "terhadap harga properti.",
                        "3. Model dapat digunakan sebagai alat bantu "
                        "penilaian properti."
                    ],
                },
                {
                    "judul": "Saran",
                    "isi": [
                        "1. Penelitian selanjutnya dapat menggunakan data "
                        "yang lebih besar.",
                        "2. Eksplorasi algoritma deep learning seperti "
                        "Neural Network."
                    ],
                },
            ],
        },
    ],
    daftar_pustaka=[
        {
            "penulis": "James, G., Witten, D., Hastie, T., dan Tibshirani, R.",
            "judul": "An Introduction to Statistical Learning",
            "edisi": "2",
            "tahun": "2023",
            "penerbit": "Springer",
        },
        {
            "penulis": "Breiman, L.",
            "judul": "Random Forests",
            "tahun": "2022",
            "penerbit": "Springer",
        },
    ],
    lampiran=[
        {
            "judul": "Lampiran A: Kode Program",
            "isi": ["Kode Python lengkap untuk analisis."],
        },
        {
            "judul": "Lampiran B: Data Sampel",
            "isi": ["Contoh data yang digunakan dalam penelitian."],
        },
    ],
)

doc.save("contoh_skripsi.docx")
print("Skripsi berhasil dibuat: contoh_skripsi.docx")
