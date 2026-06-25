# python-docx-idn Implementation Plan

> **For Hermes:** Use subagent-driven-development skill to implement this plan task-by-task.

**Goal:** Build a Python library (`docx_idn`) that generates properly formatted Indonesian academic documents (Makalah, Laporan, KTI, Skripsi) from dict/JSON data with A4 paper, 4-4-3-3 margins, structured ruler indentations, and auto TOC.

**Architecture:** Component-based design where each document section (cover, abstract, TOC, chapters, bibliography, appendix) is a separate module. A central style system handles all formatting. Each document type (Makalah, Laporan, KTI, Skripsi) is a class that composes components with its own configuration defaults.

**Tech Stack:** Python 3.10+, python-docx, pytest

---

## Project Structure

```
python-docx-idn/
├── src/docx_idn/
│   ├── __init__.py              # Public API exports
│   ├── core/
│   │   ├── __init__.py
│   │   ├── document.py          # Base DocxDocument class
│   │   ├── style.py             # StyleManager: fonts, margins, spacing, indentation
│   │   └── config.py            # DocumentConfig dataclass per type
│   ├── components/
│   │   ├── __init__.py
│   │   ├── cover.py             # CoverPage component
│   │   ├── abstract.py          # Abstract (ID + EN)
│   │   ├── toc.py               # Table of Contents (auto field code)
│   │   ├── chapter.py           # BAB with sections/subsections
│   │   ├── bibliography.py      # Daftar Pustaka
│   │   └── appendix.py          # Lampiran
│   ├── types/
│   │   ├── __init__.py
│   │   ├── base.py              # Base document type with shared logic
│   │   ├── makalah.py           # Makalah type
│   │   ├── laporan.py           # Laporan type
│   │   ├── kti.py               # KTI type
│   │   └── skripsi.py           # Skripsi type
│   └── utils/
│       ├── __init__.py
│       └── helpers.py           # Numbering, text formatting helpers
├── tests/
│   ├── test_style.py
│   ├── test_cover.py
│   ├── test_abstract.py
│   ├── test_toc.py
│   ├── test_chapter.py
│   ├── test_bibliography.py
│   ├── test_makalah.py
│   ├── test_laporan.py
│   ├── test_kti.py
│   └── test_skripsi.py
├── examples/
│   ├── contoh_makalah.py
│   ├── contoh_laporan.py
│   ├── contoh_kti.py
│   └── contoh_skripsi.py
├── pyproject.toml
└── README.md
```

---

## Formatting Standards

| Aspek | Default |
|-------|---------|
| Paper | A4 (210 x 297 mm) |
| Margins | Kiri 4cm, Kanan 4cm, Atas 3cm, Bawah 3cm |
| Font Body | Times New Roman 12pt |
| Font Heading 1 | Times New Roman 14pt Bold |
| Font Heading 2 | Times New Roman 13pt Bold |
| Font Heading 3 | Times New Roman 12pt Bold |
| Line Spacing | 1.5 (body), 2.0 (between headings and body) |
| First Line Indent | 1.25cm (4 characters) for body paragraphs |
| Justify | Full (justified) |
| Page Numbers | Bottom center, Arabic numerals |
| Cover Page Numbers | None |
| TOC | Auto field code `TOC \o "1-3" \h \z \u` — updates on open in Word |

### Heading Hierarchy

```
BAB I PENDAHULUAN          → Heading 1, Bold, Uppercase, Center
  A. Latar Belakang        → Heading 2, Bold
    1. Rumusan Masalah     → Heading 3, Bold
```

---

## Data Schema

```python
# Common schema used across all document types
doc_data = {
    "judul": "Judul Dokumen",
    "judul_en": "English Title",  # optional, required for Skripsi
    "penulis": "Nama Penulis",
    "nim": "123456789",           # optional
    "institusi": "Universitas X",
    "fakultas": "Fakultas Y",
    "prodi": "Program Studi Z",   # optional
    "angkatan": "2024",           # optional
    "pembimbing": "Dr. ABC",     # optional
    "tanggal": "Juni 2026",      # optional, defaults to current month/year
    "kata_pengantar": "...",     # optional
    "abstrak_id": "...",         # optional
    "abstrak_en": "...",         # optional, required for Skripsi
    "daftar_pustaka": [
        {
            "penulis": "Sugiyono",
            "judul": "Metode Penelitian Kuantitatif",
            "edisi": "3",
            "tahun": "2019",
            "penerbit": "Bandung: Alfabeta",
            "url": "https://...",  # optional
        }
    ],
    "bab": [
        {
            "judul": "Pendahuluan",
            "sections": [
                {
                    "judul": "Latar Belakang",
                    "isi": [
                        "Paragraf 1 tanpa indentasi awal...",
                        "Paragraf 2..."
                    ]
                },
                {
                    "judul": "Rumusan Masalah",
                    "isi": ["1. Apa dampak AI...?"]
                }
            ]
        }
    ],
    "lampiran": [  # optional
        {
            "judul": "Lampiran A: Kuesioner",
            "isi": ["Isi lampiran..."]
        }
    ]
}
```

---

## Tasks

### Task 1: Project Setup — pyproject.toml and package skeleton

**Objective:** Initialize the Python package with proper project metadata and dependencies.

**Files:**
- Create: `pyproject.toml`
- Create: `src/docx_idn/__init__.py`
- Create: `src/docx_idn/core/__init__.py`
- Create: `src/docx_idn/components/__init__.py`
- Create: `src/docx_idn/types/__init__.py`
- Create: `src/docx_idn/utils/__init__.py`

**Step 1: Create pyproject.toml**

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[project]
name = "docx-idn"
version = "0.1.0"
description = "Generate Indonesian academic documents (makalah, laporan, KTI, skripsi) from dict/JSON"
readme = "README.md"
license = "MIT"
requires-python = ">=3.10"
dependencies = [
    "python-docx>=1.1.0",
]

[project.optional-dependencies]
dev = [
    "pytest>=8.0",
]

[tool.hatch.build.targets.wheel]
packages = ["src/docx_idn"]

[tool.pytest.ini_options]
testpaths = ["tests"]
pythonpath = ["src"]
```

**Step 2: Create empty __init__.py files**

```bash
touch src/docx_idn/__init__.py
touch src/docx_idn/core/__init__.py
touch src/docx_idn/components/__init__.py
touch src/docx_idn/types/__init__.py
touch src/docx_idn/utils/__init__.py
```

**Step 3: Install in dev mode and verify**

```bash
cd "D:\Docker\Project\python-docx-idn"
pip install -e ".[dev]"
python -c "import docx_idn; print('OK')"
```

Expected: prints "OK"

---

### Task 2: Core Config — DocumentConfig dataclass

**Objective:** Define the configuration dataclass that holds all formatting parameters per document type.

**Files:**
- Create: `src/docx_idn/core/config.py`

**Step 1: Write failing test**

Create `tests/test_config.py`:

```python
from docx_idn.core.config import DocumentConfig

def test_default_config():
    cfg = DocumentConfig()
    assert cfg.paper_width == 210  # mm
    assert cfg.paper_height == 297  # mm
    assert cfg.margin_left == 4.0  # cm
    assert cfg.margin_right == 4.0
    assert cfg.margin_top == 3.0
    assert cfg.margin_bottom == 3.0
    assert cfg.font_name == "Times New Roman"
    assert cfg.font_size == 12
    assert cfg.line_spacing == 1.5
    assert cfg.first_line_indent == 1.25  # cm

def test_makalah_config():
    cfg = DocumentConfig.makalah()
    assert cfg.has_abstrak_en is False
    assert cfg.has_daftar_gambar is False

def test_skripsi_config():
    cfg = DocumentConfig.skripsi()
    assert cfg.has_abstrak_en is True
    assert cfg.has_daftar_gambar is True
    assert cfg.has_daftar_tabel is True
```

**Step 2: Run test to verify failure**

```bash
pytest tests/test_config.py -v
```

Expected: FAIL — ModuleNotFoundError

**Step 3: Implement DocumentConfig**

```python
from dataclasses import dataclass, field

@dataclass
class DocumentConfig:
    # Paper
    paper_width: int = 210      # mm (A4)
    paper_height: int = 297     # mm (A4)

    # Margins in cm
    margin_left: float = 4.0
    margin_right: float = 4.0
    margin_top: float = 3.0
    margin_bottom: float = 3.0

    # Font
    font_name: str = "Times New Roman"
    font_size: int = 12
    font_size_h1: int = 14
    font_size_h2: int = 13
    font_size_h3: int = 12

    # Spacing
    line_spacing: float = 1.5
    first_line_indent: float = 1.25  # cm
    space_after_paragraph: float = 0.0  # pt
    space_before_heading1: float = 24.0  # pt
    space_after_heading1: float = 12.0   # pt
    space_before_heading2: float = 12.0  # pt
    space_after_heading2: float = 6.0    # pt

    # Document features
    has_kata_pengantar: bool = True
    has_abstrak_id: bool = True
    has_abstrak_en: bool = False
    has_daftar_isi: bool = True
    has_daftar_tabel: bool = False
    has_daftar_gambar: bool = False
    has_lampiran: bool = False

    # Page numbering
    cover_page_number: bool = False
    page_number_start: int = 1

    @classmethod
    def makalah(cls) -> "DocumentConfig":
        return cls(
            has_abstrak_en=False,
            has_daftar_tabel=False,
            has_daftar_gambar=False,
            has_lampiran=False,
        )

    @classmethod
    def laporan(cls) -> "DocumentConfig":
        return cls(
            has_abstrak_en=False,
            has_daftar_tabel=True,
            has_daftar_gambar=True,
            has_lampiran=True,
        )

    @classmethod
    def kti(cls) -> "DocumentConfig":
        return cls(
            has_abstrak_en=True,
            has_daftar_tabel=True,
            has_daftar_gambar=True,
            has_lampiran=True,
        )

    @classmethod
    def skripsi(cls) -> "DocumentConfig":
        return cls(
            has_abstrak_en=True,
            has_daftar_tabel=True,
            has_daftar_gambar=True,
            has_lampiran=True,
        )
```

**Step 4: Run test to verify pass**

```bash
pytest tests/test_config.py -v
```

Expected: PASS

---

### Task 3: Core Style — StyleManager

**Objective:** Create the centralized style manager that applies all formatting to a python-docx Document.

**Files:**
- Create: `src/docx_idn/core/style.py`

**Step 1: Write failing test**

Create `tests/test_style.py`:

```python
from docx import Document
from docx_idn.core.style import StyleManager
from docx_idn.core.config import DocumentConfig
from docx.shared import Cm

def test_setup_page():
    doc = Document()
    cfg = DocumentConfig()
    sm = StyleManager(doc, cfg)
    sm.setup_page()
    section = doc.sections[0]
    assert section.left_margin == Cm(4.0)
    assert section.right_margin == Cm(4.0)
    assert section.top_margin == Cm(3.0)
    assert section.bottom_margin == Cm(3.0)

def test_normal_style():
    doc = Document()
    cfg = DocumentConfig()
    sm = StyleManager(doc, cfg)
    sm.setup_styles()
    style = doc.styles["Normal"]
    font = style.font
    assert font.name == "Times New Roman"
    assert font.size.pt == 12
    pf = style.paragraph_format
    assert pf.line_spacing == 1.5

def test_heading1_style():
    doc = Document()
    cfg = DocumentConfig()
    sm = StyleManager(doc, cfg)
    sm.setup_styles()
    style = doc.styles["Heading 1"]
    font = style.font
    assert font.name == "Times New Roman"
    assert font.size.pt == 14
    assert font.bold is True
```

**Step 2: Run test to verify failure**

```bash
pytest tests/test_style.py -v
```

Expected: FAIL — ModuleNotFoundError

**Step 3: Implement StyleManager**

```python
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx_idn.core.config import DocumentConfig


class StyleManager:
    def __init__(self, doc: Document, config: DocumentConfig):
        self.doc = doc
        self.cfg = config

    def setup_page(self):
        section = self.doc.sections[0]
        section.page_width = Cm(self.cfg.paper_width / 10)
        section.page_height = Cm(self.cfg.paper_height / 10)
        section.left_margin = Cm(self.cfg.margin_left)
        section.right_margin = Cm(self.cfg.margin_right)
        section.top_margin = Cm(self.cfg.margin_top)
        section.bottom_margin = Cm(self.cfg.margin_bottom)

    def setup_styles(self):
        self._setup_normal()
        self._setup_heading1()
        self._setup_heading2()
        self._setup_heading3()

    def _setup_normal(self):
        style = self.doc.styles["Normal"]
        style.font.name = self.cfg.font_name
        style.font.size = Pt(self.cfg.font_size)
        style.paragraph_format.line_spacing = self.cfg.line_spacing
        style.paragraph_format.first_line_indent = Cm(self.cfg.first_line_indent)
        style.paragraph_format.space_after = Pt(self.cfg.space_after_paragraph)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def _setup_heading1(self):
        style = self.doc.styles["Heading 1"]
        style.font.name = self.cfg.font_name
        style.font.size = Pt(self.cfg.font_size_h1)
        style.font.bold = True
        style.paragraph_format.space_before = Pt(self.cfg.space_before_heading1)
        style.paragraph_format.space_after = Pt(self.cfg.space_after_heading1)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style.paragraph_format.first_line_indent = None

    def _setup_heading2(self):
        style = self.doc.styles["Heading 2"]
        style.font.name = self.cfg.font_name
        style.font.size = Pt(self.cfg.font_size_h2)
        style.font.bold = True
        style.paragraph_format.space_before = Pt(self.cfg.space_before_heading2)
        style.paragraph_format.space_after = Pt(self.cfg.space_after_heading2)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        style.paragraph_format.first_line_indent = None

    def _setup_heading3(self):
        style = self.doc.styles["Heading 3"]
        style.font.name = self.cfg.font_name
        style.font.size = Pt(self.cfg.font_size_h3)
        style.font.bold = True
        style.paragraph_format.first_line_indent = None

    def setup_default_paragraph(self, paragraph):
        """Apply default body formatting to a paragraph."""
        paragraph.paragraph_format.line_spacing = self.cfg.line_spacing
        paragraph.paragraph_format.first_line_indent = Cm(self.cfg.first_line_indent)
        paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in paragraph.runs:
            run.font.name = self.cfg.font_name
            run.font.size = Pt(self.cfg.font_size)
```

**Step 4: Run test to verify pass**

```bash
pytest tests/test_style.py -v
```

Expected: PASS

---

### Task 4: Core Document — Base DocxDocument class

**Objective:** Create the base document class that orchestrates style setup and component rendering.

**Files:**
- Create: `src/docx_idn/core/document.py`

**Step 1: Write failing test**

Create `tests/test_document.py`:

```python
from docx import Document
from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig

def test_docx_document_creates_doc():
    doc = DocxDocument(DocumentConfig())
    assert doc.document is not None

def test_docx_document_saves(tmp_path):
    doc = DocxDocument(DocumentConfig())
    doc.add_paragraph("Test")
    out = tmp_path / "test.docx"
    doc.save(str(out))
    assert out.exists()
```

**Step 2: Run test to verify failure**

```bash
pytest tests/test_document.py -v
```

Expected: FAIL — ModuleNotFoundError

**Step 3: Implement DocxDocument**

```python
from docx import Document
from docx_idn.core.config import DocumentConfig
from docx_idn.core.style import StyleManager


class DocxDocument:
    def __init__(self, config: DocumentConfig | None = None):
        self.config = config or DocumentConfig()
        self.document = Document()
        self.style_manager = StyleManager(self.document, self.config)
        self.style_manager.setup_page()
        self.style_manager.setup_styles()

    def add_paragraph(self, text: str, style: str | None = None):
        p = self.document.add_paragraph(text, style=style)
        if style is None:
            self.style_manager.setup_default_paragraph(p)
        return p

    def add_heading(self, text: str, level: int = 1):
        return self.document.add_heading(text, level=level)

    def add_page_break(self):
        self.document.add_page_break()

    def save(self, path: str):
        self.document.save(path)
```

**Step 4: Run test to verify pass**

```bash
pytest tests/test_document.py -v
```

Expected: PASS

---

### Task 5: Component — CoverPage

**Objective:** Build the cover page component with centered layout, institution info, and title.

**Files:**
- Create: `src/docx_idn/components/cover.py`

**Step 1: Write failing test**

Create `tests/test_cover.py`:

```python
from docx import Document
from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.cover import CoverPage

def test_cover_renders():
    doc = DocxDocument(DocumentConfig())
    cover = CoverPage(
        judul="Dampak AI terhadap Pendidikan",
        penulis="Ahmad Fauzi",
        nim="123456789",
        institusi="Universitas Indonesia",
        fakultas="Ilmu Komputer",
        prodi="Teknik Informatika",
        angkatan="2024",
        tanggal="Juni 2026",
    )
    cover.render(doc)
    doc.save("test_output_cover.docx")
    # Verify document has content
    assert len(doc.document.paragraphs) > 0
```

**Step 2: Run test to verify failure**

```bash
pytest tests/test_cover.py -v
```

Expected: FAIL — ModuleNotFoundError

**Step 3: Implement CoverPage**

```python
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm


class CoverPage:
    def __init__(
        self,
        judul: str,
        penulis: str,
        institusi: str,
        fakultas: str,
        nim: str = "",
        prodi: str = "",
        angkatan: str = "",
        tanggal: str = "",
        pembimbing: str = "",
    ):
        self.judul = judul
        self.penulis = penulis
        self.nim = nim
        self.institusi = institusi
        self.fakultas = fakultas
        self.prodi = prodi
        self.angkatan = angkatan
        self.tanggal = tanggal
        self.pembimbing = pembimbing

    def render(self, doc):
        cfg = doc.config
        document = doc.document

        # Remove first default paragraph
        if document.paragraphs:
            p = document.paragraphs[0]
            p.clear()

        # Add vertical spacing
        for _ in range(4):
            sp = document.add_paragraph()
            sp.paragraph_format.space_before = Pt(0)
            sp.paragraph_format.space_after = Pt(0)

        # Institution name
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(self.institusi.upper())
        run.font.name = cfg.font_name
        run.font.size = Pt(16)
        run.bold = True

        # Faculty
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(self.fakultas)
        run.font.name = cfg.font_name
        run.font.size = Pt(14)
        run.bold = True

        # Study program
        if self.prodi:
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(self.prodi)
            run.font.name = cfg.font_name
            run.font.size = Pt(13)

        # Spacing
        document.add_paragraph()

        # Title
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(12)
        run = p.add_run(self.judul.upper())
        run.font.name = cfg.font_name
        run.font.size = Pt(16)
        run.bold = True

        # Spacing
        for _ in range(3):
            document.add_paragraph()

        # Author info
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"Nama Penulis: {self.penulis}")
        run.font.name = cfg.font_name
        run.font.size = Pt(12)

        if self.nim:
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(f"NIM: {self.nim}")
            run.font.name = cfg.font_name
            run.font.size = Pt(12)

        if self.angkatan:
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(f"Angkatan: {self.angkatan}")
            run.font.name = cfg.font_name
            run.font.size = Pt(12)

        if self.pembimbing:
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(f"Pembimbing: {self.pembimbing}")
            run.font.name = cfg.font_name
            run.font.size = Pt(12)

        # Spacing
        for _ in range(2):
            document.add_paragraph()

        # Year
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(self.tanggal or "2026")
        run.font.name = cfg.font_name
        run.font.size = Pt(12)

        doc.add_page_break()
```

**Step 4: Run test to verify pass**

```bash
pytest tests/test_cover.py -v
```

Expected: PASS

---

### Task 6: Component — Abstract

**Objective:** Build the abstract component supporting Indonesian and/or English abstract.

**Files:**
- Create: `src/docx_idn/components/abstract.py`

**Step 1: Write failing test**

Create `tests/test_abstract.py`:

```python
from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.abstract import Abstract

def test_abstract_id_only():
    doc = DocxDocument(DocumentConfig.makalah())
    abstract = Abstract(abstrak_id="Ini abstrak Indonesia.")
    abstract.render(doc)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("Abstrak" in t for t in texts)
    assert any("Ini abstrak Indonesia" in t for t in texts)

def test_abstract_both():
    doc = DocxDocument(DocumentConfig.kti())
    abstract = Abstract(
        abstrak_id="Abstrak Indonesia.",
        abstrak_en="English abstract.",
    )
    abstract.render(doc)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("Abstrak" in t for t in texts)
    assert any("Abstract" in t for t in texts)
```

**Step 2: Run test to verify failure**

```bash
pytest tests/test_abstract.py -v
```

Expected: FAIL

**Step 3: Implement Abstract**

```python
from docx.enum.text import WD_ALIGN_PARAGRAPH


class Abstract:
    def __init__(self, abstrak_id: str = "", abstrak_en: str = ""):
        self.abstrak_id = abstrak_id
        self.abstrak_en = abstrak_en

    def render(self, doc):
        if self.abstrak_id:
            self._render_section(doc, "Abstrak", self.abstrak_id)
        if self.abstrak_en:
            self._render_section(doc, "Abstract", self.abstrak_en)

    def _render_section(self, doc, title: str, text: str):
        doc.add_heading(title, level=1)
        p = doc.add_paragraph(text)
        doc.add_page_break()
```

**Step 4: Run test to verify pass**

```bash
pytest tests/test_abstract.py -v
```

Expected: PASS

---

### Task 7: Component — Table of Contents (TOC)

**Objective:** Insert auto-updating TOC field code into the document.

**Files:**
- Create: `src/docx_idn/components/toc.py`

**Step 1: Write failing test**

Create `tests/test_toc.py`:

```python
from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.toc import TableOfContents

def test_toc_renders():
    doc = DocxDocument(DocumentConfig())
    toc = TableOfContents()
    toc.render(doc)
    # Should have added paragraphs (heading + field)
    assert len(doc.document.paragraphs) > 0
    doc.save("test_output_toc.docx")
```

**Step 2: Run test to verify failure**

```bash
pytest tests/test_toc.py -v
```

Expected: FAIL

**Step 3: Implement TableOfContents**

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH


class TableOfContents:
    def __init__(self, title: str = "DAFTAR ISI", levels: str = "1-3"):
        self.title = title
        self.levels = levels

    def render(self, doc):
        # Add TOC heading
        heading = doc.add_heading(self.title, level=1)

        # Create TOC field code
        paragraph = doc.document.add_paragraph()
        run = paragraph.add_run()

        # Field begin
        fld_char_begin = OxmlElement("w:fldChar")
        fld_char_begin.set(qn("w:fldCharType"), "begin")
        run._r.append(fld_char_begin)

        # Field instruction
        instr_text = OxmlElement("w:instrText")
        instr_text.set(qn("xml:space"), "preserve")
        instr_text.text = f' TOC \\o "{self.levels}" \\h \\z \\u '
        run._r.append(instr_text)

        # Field separate
        fld_char_sep = OxmlElement("w:fldChar")
        fld_char_sep.set(qn("w:fldCharType"), "separate")
        run._r.append(fld_char_sep)

        # Placeholder text
        r2 = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = "Klik kanan → Update Field untuk memperbarui Daftar Isi"
        r2.append(t)
        paragraph._p.append(r2)

        # Field end
        r3 = OxmlElement("w:r")
        fld_char_end = OxmlElement("w:fldChar")
        fld_char_end.set(qn("w:fldCharType"), "end")
        r3.append(fld_char_end)
        paragraph._p.append(r3)

        doc.add_page_break()
```

**Step 4: Run test to verify pass**

```bash
pytest tests/test_toc.py -v
```

Expected: PASS

---

### Task 8: Component — Chapter

**Objective:** Build the chapter component that renders BAB headings with sections and subsections with proper ruler indentation.

**Files:**
- Create: `src/docx_idn/components/chapter.py`

**Step 1: Write failing test**

Create `tests/test_chapter.py`:

```python
from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.chapter import Chapter

def test_chapter_renders():
    doc = DocxDocument(DocumentConfig())
    chapter = Chapter(
        judul="Pendahuluan",
        sections=[
            {
                "judul": "Latar Belakang",
                "isi": ["Ini paragraf latar belakang.", "Paragraf kedua."]
            },
            {
                "judul": "Rumusan Masalah",
                "isi": ["1. Apa dampak AI?"]
            }
        ]
    )
    chapter.render(doc, number=1)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("BAB I" in t for t in texts)
    assert any("Pendahuluan" in t for t in texts)
    assert any("Latar Belakang" in t for t in texts)

def test_chapter_with_subsections():
    doc = DocxDocument(DocumentConfig())
    chapter = Chapter(
        judul="Tinjauan Pustaka",
        sections=[
            {
                "judul": "Pengertian AI",
                "subsections": [
                    {
                        "judul": "Definisi Umum",
                        "isi": ["AI adalah..."]
                    }
                ],
                "isi": ["AI merupakan..."]
            }
        ]
    )
    chapter.render(doc, number=2)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("BAB II" in t for t in texts)
    assert any("Definisi Umum" in t for t in texts)
```

**Step 2: Run test to verify failure**

```bash
pytest tests/test_chapter.py -v
```

Expected: FAIL

**Step 3: Implement Chapter**

```python
ROMAN_NUMERALS = {
    1: "I", 2: "II", 3: "III", 4: "IV", 5: "V",
    6: "VI", 7: "VII", 8: "VIII", 9: "IX", 10: "X",
}


class Chapter:
    def __init__(self, judul: str, sections: list[dict] | None = None):
        self.judul = judul
        self.sections = sections or []

    def render(self, doc, number: int):
        roman = ROMAN_NUMERALS.get(number, str(number))
        heading_text = f"BAB {roman} {self.judul.upper()}"
        doc.add_heading(heading_text, level=1)

        for i, section in enumerate(self.sections, start=1):
            self._render_section(doc, section, f"{number}.{i}")

    def _render_section(self, section: dict, number: str):
        label = f"{section['judul']}"
        doc.add_heading(label, level=2)

        # Render body paragraphs
        for text in section.get("isi", []):
            doc.add_paragraph(text)

        # Render subsections (level 3)
        for j, sub in enumerate(section.get("subsections", []), start=1):
            self._render_subsection(doc, sub, f"{number}.{j}")

    def _render_subsection(self, subsection: dict, number: str):
        label = f"{subsection['judul']}"
        doc.add_heading(label, level=3)

        for text in subsection.get("isi", []):
            doc.add_paragraph(text)
```

**Step 4: Run test to verify pass**

```bash
pytest tests/test_chapter.py -v
```

Expected: PASS

---

### Task 9: Component — Bibliography

**Objective:** Build the bibliography component that renders references in hanging indent format.

**Files:**
- Create: `src/docx_idn/components/bibliography.py`

**Step 1: Write failing test**

Create `tests/test_bibliography.py`:

```python
from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.bibliography import Bibliography

def test_bibliography_renders():
    doc = DocxDocument(DocumentConfig())
    bib = Bibliography(
        entries=[
            {
                "penulis": "Sugiyono",
                "judul": "Metode Penelitian Kuantitatif",
                "tahun": "2019",
                "penerbit": "Bandung: Alfabeta",
            },
            {
                "penulis": "Moleong, L.J.",
                "judul": "Metodologi Penelitian Kualitatif",
                "edisi": "37",
                "tahun": "2018",
                "penerbit": "Bandung: Remaja Rosdakarya",
            },
        ]
    )
    bib.render(doc)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("DAFTAR PUSTAKA" in t for t in texts)
    assert any("Sugiyono" in t for t in texts)
    assert any("Moleong" in t for t in texts)
```

**Step 2: Run test to verify failure**

```bash
pytest tests/test_bibliography.py -v
```

Expected: FAIL

**Step 3: Implement Bibliography**

```python
from docx.shared import Cm


class Bibliography:
    def __init__(self, entries: list[dict] | None = None):
        self.entries = entries or []

    def render(self, doc):
        doc.add_heading("DAFTAR PUSTAKA", level=1)

        for entry in self.entries:
            text = self._format_entry(entry)
            p = doc.add_paragraph(text)
            # Hanging indent: left indent 1.25cm, first line -1.25cm
            p.paragraph_format.left_indent = Cm(1.25)
            p.paragraph_format.first_line_indent = Cm(-1.25)
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.alignment = None  # left aligned

    def _format_entry(self, entry: dict) -> str:
        parts = []
        penulis = entry.get("penulis", "")
        if penulis:
            parts.append(penulis)

        tahun = entry.get("tahun", "")
        if tahun:
            parts.append(f"({tahun})")

        judul = entry.get("judul", "")
        if judul:
            parts.append(f"{judul}.")

        edisi = entry.get("edisi", "")
        if edisi:
            parts.append(f"Edisi {edisi}.")

        penerbit = entry.get("penerbit", "")
        if penerbit:
            parts.append(f"{penerbit}.")

        url = entry.get("url", "")
        if url:
            parts.append(f"Diperoleh dari {url}")

        return " ".join(parts)
```

**Step 4: Run test to verify pass**

```bash
pytest tests/test_bibliography.py -v
```

Expected: PASS

---

### Task 10: Component — Appendix

**Objective:** Build the appendix component.

**Files:**
- Create: `src/docx_idn/components/appendix.py`

**Step 1: Write failing test**

Create `tests/test_appendix.py`:

```python
from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.appendix import Appendix

def test_appendix_renders():
    doc = DocxDocument(DocumentConfig())
    appendix = Appendix(
        items=[
            {"judul": "Lampiran A: Kuesioner", "isi": ["Isi kuesioner..."]},
            {"judul": "Lampiran B: Data", "isi": ["Data tabel..."]},
        ]
    )
    appendix.render(doc)
    texts = [p.text for p in doc.document.paragraphs]
    assert any("LAMPIRAN" in t for t in texts)
    assert any("Kuesioner" in t for t in texts)
```

**Step 2: Run test to verify failure**

```bash
pytest tests/test_appendix.py -v
```

Expected: FAIL

**Step 3: Implement Appendix**

```python
class Appendix:
    def __init__(self, items: list[dict] | None = None):
        self.items = items or []

    def render(self, doc):
        doc.add_heading("LAMPIRAN", level=1)

        for item in self.items:
            doc.add_heading(item["judul"], level=2)
            for text in item.get("isi", []):
                doc.add_paragraph(text)
```

**Step 4: Run test to verify pass**

```bash
pytest tests/test_appendix.py -v
```

Expected: PASS

---

### Task 11: Type — Base document type

**Objective:** Create the base class for document types that orchestrates all components.

**Files:**
- Create: `src/docx_idn/types/base.py`

**Step 1: Write failing test**

Create `tests/test_base_type.py`:

```python
from docx_idn.types.base import BaseDocument
from docx_idn.core.config import DocumentConfig

def test_base_document_minimal():
    doc = BaseDocument(
        config=DocumentConfig.makalah(),
        judul="Test",
        penulis="Author",
        institusi="Universitas",
        fakultas="Fakultas",
        bab=[{"judul": "Bab 1", "sections": [{"judul": "S1", "isi": ["Content"]}] }],
    )
    doc.save("test_base_output.docx")
    assert True
```

**Step 2: Run test to verify failure**

```bash
pytest tests/test_base_type.py -v
```

Expected: FAIL

**Step 3: Implement BaseDocument**

```python
from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
from docx_idn.components.cover import CoverPage
from docx_idn.components.abstract import Abstract
from docx_idn.components.toc import TableOfContents
from docx_idn.components.chapter import Chapter
from docx_idn.components.bibliography import Bibliography
from docx_idn.components.appendix import Appendix


class BaseDocument:
    def __init__(
        self,
        config: DocumentConfig,
        judul: str,
        penulis: str,
        institusi: str,
        fakultas: str,
        bab: list[dict],
        nim: str = "",
        prodi: str = "",
        angkatan: str = "",
        tanggal: str = "",
        pembimbing: str = "",
        judul_en: str = "",
        kata_pengantar: str = "",
        abstrak_id: str = "",
        abstrak_en: str = "",
        daftar_pustaka: list[dict] | None = None,
        lampiran: list[dict] | None = None,
    ):
        self.config = config
        self.doc = DocxDocument(config)

        # Cover
        cover = CoverPage(
            judul=judul,
            penulis=penulis,
            nim=nim,
            institusi=institusi,
            fakultas=fakultas,
            prodi=prodi,
            angkatan=angkatan,
            tanggal=tanggal,
            pembimbing=pembimbing,
        )
        cover.render(self.doc)

        # Kata Pengantar
        if kata_pengantar and config.has_kata_pengantar:
            self.doc.add_heading("KATA PENGANTAR", level=1)
            self.doc.add_paragraph(kata_pengantar)
            self.doc.add_page_break()

        # Abstract
        if config.has_abstrak_id or config.has_abstrak_en:
            abstract = Abstract(
                abstrak_id=abstrak_id if config.has_abstrak_id else "",
                abstrak_en=abstrak_en if config.has_abstrak_en else "",
            )
            abstract.render(self.doc)

        # TOC
        if config.has_daftar_isi:
            toc = TableOfContents()
            toc.render(self.doc)

        # Chapters
        for i, b in enumerate(bab, start=1):
            chapter = Chapter(judul=b["judul"], sections=b.get("sections"))
            chapter.render(self.doc, number=i)
            # Page break between chapters
            if i < len(bab):
                self.doc.add_page_break()

        # Bibliography
        if daftar_pustaka:
            bib = Bibliography(entries=daftar_pustaka)
            bib.render(self.doc)

        # Appendix
        if lampiran and config.has_lampiran:
            appendix = Appendix(items=lampiran)
            appendix.render(self.doc)

    def save(self, path: str):
        self.doc.save(path)
```

**Step 4: Run test to verify pass**

```bash
pytest tests/test_base_type.py -v
```

Expected: PASS

---

### Task 12: Types — Makalah, Laporan, KTI, Skripsi

**Objective:** Create thin wrapper classes for each document type with their specific defaults.

**Files:**
- Create: `src/docx_idn/types/makalah.py`
- Create: `src/docx_idn/types/laporan.py`
- Create: `src/docx_idn/types/kti.py`
- Create: `src/docx_idn/types/skripsi.py`

**Step 1: Write failing test**

Create `tests/test_makalah.py`:

```python
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
```

Create `tests/test_laporan.py`:

```python
from docx_idn.types.laporan import Laporan

def test_laporan_generates():
    l = Laporan(
        judul="Laporan Observasi",
        penulis="Budi Santoso",
        institusi="Universitas Gadjah Mada",
        fakultas="Ekonomi",
        bab=[
            {"judul": "Pendahuluan", "sections": [{"judul": "Latar Belakang", "isi": ["Observasi..."]}]},
            {"judul": "Hasil Observasi", "sections": [{"judul": "Temuan", "isi": ["Data menunjukkan..."]}]},
        ],
    )
    l.save("test_output_laporan.docx")
    assert True
```

Create `tests/test_kti.py`:

```python
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
```

Create `tests/test_skripsi.py`:

```python
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
```

**Step 2: Run tests to verify failure**

```bash
pytest tests/test_makalah.py tests/test_laporan.py tests/test_kti.py tests/test_skripsi.py -v
```

Expected: FAIL

**Step 3: Implement all four types**

`src/docx_idn/types/makalah.py`:
```python
from docx_idn.types.base import BaseDocument
from docx_idn.core.config import DocumentConfig


class Makalah(BaseDocument):
    def __init__(self, **kwargs):
        kwargs.setdefault("config", DocumentConfig.makalah())
        super().__init__(**kwargs)
```

`src/docx_idn/types/laporan.py`:
```python
from docx_idn.types.base import BaseDocument
from docx_idn.core.config import DocumentConfig


class Laporan(BaseDocument):
    def __init__(self, **kwargs):
        kwargs.setdefault("config", DocumentConfig.laporan())
        super().__init__(**kwargs)
```

`src/docx_idn/types/kti.py`:
```python
from docx_idn.types.base import BaseDocument
from docx_idn.core.config import DocumentConfig


class KTI(BaseDocument):
    def __init__(self, **kwargs):
        kwargs.setdefault("config", DocumentConfig.kti())
        super().__init__(**kwargs)
```

`src/docx_idn/types/skripsi.py`:
```python
from docx_idn.types.base import BaseDocument
from docx_idn.core.config import DocumentConfig


class Skripsi(BaseDocument):
    def __init__(self, **kwargs):
        kwargs.setdefault("config", DocumentConfig.skripsi())
        super().__init__(**kwargs)
```

**Step 4: Run tests to verify pass**

```bash
pytest tests/test_makalah.py tests/test_laporan.py tests/test_kti.py tests/test_skripsi.py -v
```

Expected: PASS

---

### Task 13: Public API — __init__.py exports

**Objective:** Wire up the public API so users can `from docx_idn import Makalah`.

**Files:**
- Modify: `src/docx_idn/__init__.py`

**Step 1: Implement exports**

```python
from docx_idn.types.makalah import Makalah
from docx_idn.types.laporan import Laporan
from docx_idn.types.kti import KTI
from docx_idn.types.skripsi import Skripsi
from docx_idn.core.config import DocumentConfig
from docx_idn.core.document import DocxDocument

__all__ = ["Makalah", "Laporan", "KTI", "Skripsi", "DocumentConfig", "DocxDocument"]
__version__ = "0.1.0"
```

**Step 2: Verify import works**

```bash
python -c "from docx_idn import Makalah, Laporan, KTI, Skripsi; print('All imports OK')"
```

Expected: All imports OK

---

### Task 14: Run full test suite

**Objective:** Verify everything works together.

**Step 1: Run all tests**

```bash
pytest tests/ -v
```

Expected: All tests PASS

**Step 2: Generate example documents**

```bash
python examples/contoh_makalah.py
python examples/contoh_laporan.py
python examples/contoh_kti.py
python examples/contoh_skripsi.py
```

Verify `.docx` files are created.

---

### Task 15: README.md

**Objective:** Write documentation for the library.

**Files:**
- Create: `README.md`

Content should cover: installation, quick start, data schema, document types, configuration, and examples.

---

## Execution Order

| Task | Depends On | Description |
|------|-----------|-------------|
| 1 | — | Project setup |
| 2 | 1 | DocumentConfig |
| 3 | 1, 2 | StyleManager |
| 4 | 1, 2, 3 | DocxDocument base |
| 5 | 4 | CoverPage |
| 6 | 4 | Abstract |
| 7 | 4 | TableOfContents |
| 8 | 4 | Chapter |
| 9 | 4 | Bibliography |
| 10 | 4 | Appendix |
| 11 | 4, 5-10 | BaseDocument orchestrator |
| 12 | 11 | Type wrappers |
| 13 | 12 | Public API |
| 14 | 13 | Full test suite |
| 15 | 14 | Documentation |
