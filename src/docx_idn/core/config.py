from dataclasses import dataclass


@dataclass
class DocumentConfig:
    # Paper
    paper_width: int = 210      # mm (A4)
    paper_height: int = 297     # mm (A4)

    # Margins in cm (Standar Indonesia)
    margin_left: float = 4.0
    margin_right: float = 3.0
    margin_top: float = 4.0
    margin_bottom: float = 3.0

    # Font
    font_name: str = "Times New Roman"
    font_size: int = 12
    font_size_h1: int = 14
    font_size_h2: int = 12
    font_size_h3: int = 12

    # Spacing
    line_spacing: float = 1.5
    first_line_indent: float = 1.25  # cm (4 karakter)
    space_after_paragraph: float = 0.0  # pt
    space_before_heading1: float = 24.0  # pt (12pt before)
    space_after_heading1: float = 12.0   # pt (6pt after)
    space_before_heading2: float = 12.0  # pt
    space_after_heading2: float = 6.0    # pt

    # Document features
    has_cover: bool = False       # User handles cover manually
    has_kata_pengantar: bool = True
    has_abstrak_id: bool = True
    has_abstrak_en: bool = False
    has_daftar_isi: bool = True
    has_lampiran: bool = False

    @classmethod
    def makalah(cls) -> "DocumentConfig":
        return cls(
            has_cover=False,
            has_abstrak_en=False,
            has_lampiran=False,
        )

    @classmethod
    def laporan(cls) -> "DocumentConfig":
        return cls(
            has_cover=False,
            has_abstrak_en=False,
            has_lampiran=True,
        )

    @classmethod
    def kti(cls) -> "DocumentConfig":
        return cls(
            has_cover=False,
            has_abstrak_en=True,
            has_lampiran=True,
        )

    @classmethod
    def skripsi(cls) -> "DocumentConfig":
        return cls(
            has_cover=False,
            has_abstrak_en=True,
            has_lampiran=True,
        )
