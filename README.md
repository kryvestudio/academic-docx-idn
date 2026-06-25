# docx-idn 🇮🇩

[![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![PyPI](https://img.shields.io/pypi/v/docx-idn.svg)](https://pypi.org/project/docx-idn/)
[![GitHub issues](https://img.shields.io/github/issues/kryvestudio/academic-docx-idn.svg)](https://github.com/kryvestudio/academic-docx-idn/issues)
[![GitHub stars](https://img.shields.io/github/stars/kryvestudio/academic-docx-idn.svg)](https://github.com/kryvestudio/academic-docx-idn/stargazers)

Python library untuk generate dokumen akademik Indonesia (makalah, laporan, KTI, skripsi) dengan format standar institusi.

## ⚠️ DISCLAIMER

> **Program ini hanya sebagai alat bantu (tool) untuk mempermudah pembuatan dokumen akademik dalam format standar Indonesia.**
>
> - Program ini **TIDAK menentukan** nilai, kualitas, atau keaslian dari isi dokumen yang dihasilkan.
> - Pengguna **bertanggung jawab penuh** atas konten yang dimasukkan ke dalam dokumen.
> - Format yang dihasilkan merupakan **standar umum** dan mungkin berbeda dengan ketentuan institusi tertentu. Selalu cocokkan dengan pedoman resmi dari institusi Anda.
> - Program ini **bukan pengganti** penulisan akademik yang baik dan benar.
> - Gunakan secara **etis dan bertanggung jawab**.

## Features

- A4 paper size with 4-3-4-3 margins (kiri 4cm, kanan 3cm, atas 4cm, bawah 3cm)
- Times New Roman font with proper sizing (12pt body, 14pt headings)
- 1.5 line spacing with structured ruler indentations
- Auto Table of Contents (updates when opened in Word)
- Cover page, Abstract (ID/EN), Bibliography, Appendix
- 4 document types: Makalah, Laporan, KTI, Skripsi
- Proper numbering format (BAB I, 1.1, 1.1.1)
- Hanging indent for numbered lists
- East Asian font support

## Requirements

### Python Version
- Python 3.10 atau lebih baru

### Dependencies
Dependencies akan otomatis terinstall saat Anda menjalankan perintah install:

| Package | Version | Description |
|---------|---------|-------------|
| `python-docx` | >= 1.1.0 | Core library untuk membuat dokumen Word |

### Dev Dependencies (opsional)
| Package | Version | Description |
|---------|---------|-------------|
| `pytest` | >= 8.0 | Untuk menjalankan test |

## Installation

```bash
# Clone repository
git clone https://github.com/kryvestudio/academic-docx-idn.git
cd docx-idn

# Install (hanya python-docx)
pip install -e .

# Atau termasuk dev dependencies (pytest untuk testing)
pip install -e ".[dev]"
```

### Verifikasi Install

```bash
python -c "from docx_idn import Makalah, Laporan, KTI, Skripsi; print('Install berhasil! ✅')"
```

## Quick Start

```python
from docx_idn import Makalah

doc = Makalah(
    judul="Dampak AI terhadap Pendidikan",
    penulis="Ahmad Fauzi",
    nim="123456789",
    institusi="Universitas Indonesia",
    fakultas="Ilmu Komputer",
    prodi="Teknik Informatika",
    bab=[
        {
            "judul": "Pendahuluan",
            "sections": [
                {"judul": "Latar Belakang", "isi": ["AI berkembang pesat..."]}
            ]
        }
    ],
    daftar_pustaka=[
        {"penulis": "Sugiyono", "judul": "Metode Penelitian", "tahun": "2019", "penerbit": "Alfabeta"}
    ],
)
doc.save("makalah.docx")
```

## Document Types

| Type | Use Case | Features |
|------|----------|----------|
| `Makalah` | Essay/paper | Basic structure, no English abstract |
| `Laporan` | Report | Tables, figures, appendix |
| `KTI` | Scientific paper | Full structure, English abstract |
| `Skripsi` | Thesis | Full structure, English abstract, adviser |

## Data Schema

```python
{
    "judul": "Judul Dokumen",
    "judul_en": "English Title",        # optional, required for Skripsi/KTI
    "penulis": "Nama Penulis",
    "nim": "123456789",                  # optional
    "institusi": "Universitas X",
    "fakultas": "Fakultas Y",
    "prodi": "Program Studi Z",          # optional
    "angkatan": "2024",                  # optional
    "pembimbing": "Dr. ABC",            # optional
    "tanggal": "Juni 2026",             # optional
    "kata_pengantar": "...",            # optional
    "abstrak_id": "...",                # optional
    "abstrak_en": "...",                # optional
    "daftar_pustaka": [...],            # optional
    "bab": [...],                        # required
    "lampiran": [...]                    # optional
}
```

## Chapter Structure

```python
"bab": [
    {
        "judul": "Pendahuluan",
        "sections": [
            {
                "judul": "Latar Belakang",
                "isi": ["Paragraf 1...", "Paragraf 2..."]
            },
            {
                "judul": "Sub Section",
                "subsections": [
                    {
                        "judul": "Sub Sub Section",
                        "isi": ["Content..."]
                    }
                ],
                "isi": ["Parent content..."]
            }
        ]
    }
]
```

## Custom Configuration

```python
from docx_idn import DocumentConfig, DocxDocument

config = DocumentConfig(
    margin_left=3.0,    # custom margin
    margin_right=3.0,
    font_size=11,       # custom font size
)

doc = DocxDocument(config)
doc.add_paragraph("Custom formatted document")
doc.save("custom.docx")
```

## Formatting Standards

| Aspek | Default |
|-------|---------|
| Paper | A4 (210 x 297 mm) |
| Margins | Kiri 4cm, Kanan 3cm, Atas 4cm, Bawah 3cm |
| Font | Times New Roman 12pt |
| Heading 1 | 14pt Bold, Center |
| Heading 2 | 12pt Bold, Left |
| Heading 3 | 12pt Bold |
| Line Spacing | 1.5 |
| First Line Indent | 1.25cm |
| Justify | Full |

## Running Tests

```bash
# Jalankan semua test
pytest tests/ -v

# Jalankan test tertentu
pytest tests/test_makalah.py -v
```

## Examples

See `examples/` directory for complete examples:

```bash
python examples/contoh_makalah.py
python examples/contoh_laporan.py
python examples/contoh_kti.py
python examples/contoh_skripsi.py
python examples/kti_keberagaman_indonesia.py  # KTI lengkap
```

## Troubleshooting

### Error: ModuleNotFoundError
```bash
# Pastikan sudah install
pip install -e ".[dev]"
```

### Error: Font tidak muncul
- Pastikan Microsoft Word terinstall
- Font Times New Roman biasanya sudah default di Windows

### Error: TOC tidak update
- Buka dokumen di Microsoft Word
- Klik kanan pada Daftar Isi → Update Field → Update entire table

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License - See [LICENSE](LICENSE) for details.

## 📝 Catatan Penting

Program ini dibuat untuk **mempermudah** pembuatan dokumen akademik, bukan untuk menggantikan proses penulisan yang baik. Pengguna tetap harus:

1. **Memahami isi** dokumen yang dibuat
2. **Memverifikasi** format sesuai ketentuan institusi
3. **Bertanggung jawab** atas konten yang dihasilkan
4. **Menggunakan secara etis** dalam kegiatan akademik

---

**Made with ❤️ for Indonesian Academic Community**
