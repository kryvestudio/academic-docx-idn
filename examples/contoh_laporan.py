"""Contoh pembuatan Laporan."""
from docx_idn import Laporan

doc = Laporan(
    judul="Laporan Observasi: Implementasi E-Learning di Sekolah Dasar",
    penulis="Budi Santoso",
    nim="987654321",
    institusi="Universitas Gadjah Mada",
    fakultas="Ilmu Pendidikan",
    prodi="Teknologi Pendidikan",
    angkatan="2023",
    tanggal="Juni 2026",
    abstrak_id="Laporan ini menyajikan hasil observasi implementasi e-learning "
               "di sekolah dasar negeri di Yogyakarta.",
    bab=[
        {
            "judul": "Pendahuluan",
            "sections": [
                {
                    "judul": "Latar Belakang",
                    "isi": [
                        "Pandemi COVID-19 telah mempercepat transformasi digital "
                        "di sektor pendidikan. E-learning menjadi solusi utama "
                        "dalam menjaga kontinuitas pembelajaran."
                    ],
                },
                {
                    "judul": "Tujuan Observasi",
                    "isi": [
                        "Mengidentifikasi tingkat implementasi e-learning dan "
                        "efektivitasnya terhadap proses belajar mengajar."
                    ],
                },
            ],
        },
        {
            "judul": "Hasil Observasi",
            "sections": [
                {
                    "judul": "Kondisi Sekolah",
                    "isi": [
                        "Observasi dilakukan di 5 sekolah dasar negeri di "
                        "Yogyakarta selama bulan Maret-April 2026."
                    ],
                },
                {
                    "judul": "Temuan Utama",
                    "isi": [
                        "1. 80% sekolah sudah menggunakan platform e-learning.",
                        "2. Tingkat ketergantungan guru terhadap teknologi bervariasi.",
                        "3. Siswa di kelas rendah mengalami kesulitan lebih besar.",
                    ],
                },
            ],
        },
        {
            "judul": "Pembahasan",
            "sections": [
                {
                    "judul": "Analisis Temuan",
                    "isi": [
                        "Implementasi e-learning menunjukkan hasil yang variatif "
                        "tergantung pada kesiapan infrastruktur dan SDM."
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
                        "E-learning telah menjadi bagian tak terpisahkan dari "
                        "sistem pendidikan, namun masih memerlukan peningkatan "
                        "infrastruktur dan kompetensi guru."
                    ],
                },
                {
                    "judul": "Rekomendasi",
                    "isi": [
                        "1. Pelatihan intensif bagi guru tentang e-learning.",
                        "2. Penyediaan infrastruktur yang memadai.",
                    ],
                },
            ],
        },
    ],
    daftar_pustaka=[
        {
            "penulis": "Moleong, L.J.",
            "judul": "Metodologi Penelitian Kualitatif",
            "edisi": "37",
            "tahun": "2018",
            "penerbit": "Bandung: Remaja Rosdakarya",
        },
    ],
    lampiran=[
        {
            "judul": "Lampiran A: Formulir Observasi",
            "isi": ["Formulir observasi yang digunakan dalam penelitian ini."],
        },
        {
            "judul": "Lampiran B: Foto Kegiatan",
            "isi": ["Dokumentasi kegiatan observasi di lapangan."],
        },
    ],
)

doc.save("contoh_laporan.docx")
print("Laporan berhasil dibuat: contoh_laporan.docx")
