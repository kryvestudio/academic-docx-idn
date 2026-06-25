# Contributing to docx-idn

Terima kasih telah tertarik untuk berkontribusi! 🎉

## Cara Berkontribusi

### 1. Fork Repository

```bash
# Fork melalui GitHub, lalu clone
git clone https://github.com/YOUR_USERNAME/academic-docx-idn.git
cd academic-docx-idn
```

### 2. Buat Branch Baru

```bash
git checkout -b fitur/nama-fitur-baru
```

### 3. Install Dependencies

```bash
pip install -e ".[dev]"
```

### 4. Buat Perubahan

- Tulis kode yang bersih dan terdokumentasi
- Ikuti style guide yang ada
- Tambahkan test jika diperlukan

### 5. Jalankan Test

```bash
pytest tests/ -v
```

### 6. Commit & Push

```bash
git add .
git commit -m "feat: tambahkan fitur baru"
git push origin fitur/nama-fitur-baru
```

### 7. Buat Pull Request

Buka GitHub dan buat Pull Request dari branch Anda ke `main`.

## Style Guide

- **Python**: Ikuti PEP 8
- **Commit messages**: Gunakan format conventional commits
  - `feat:` untuk fitur baru
  - `fix:` untuk perbaikan bug
  - `docs:` untuk dokumentasi
  - `test:` untuk penambahan test
  - `chore:` untuk maintenance

## Menambahkan Fitur Baru

1. Buat file baru di `src/docx_idn/components/` atau `src/docx_idn/types/`
2. Tambahkan import di `src/docx_idn/__init__.py`
3. Buat test di `tests/`
4. Update README jika diperlukan

## Melaporkan Bug

Buka [Issues](https://github.com/kryvestudio/academic-docx-idn/issues) dan jelaskan:
- Deskripsi bug
- Cara mereproduksi
- Expected behavior
- Screenshot (jika applicable)

## Pertanyaan?

Buka [Discussions](https://github.com/kryvestudio/academic-docx-idn/discussions) untuk bertanya atau berdiskusi.

---

Terima kasih! 🐱
