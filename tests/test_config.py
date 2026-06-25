from docx_idn.core.config import DocumentConfig


def test_default_config():
    cfg = DocumentConfig()
    assert cfg.paper_width == 210  # mm
    assert cfg.paper_height == 297  # mm
    assert cfg.margin_left == 4.0  # cm
    assert cfg.margin_right == 3.0  # cm
    assert cfg.margin_top == 4.0   # cm (standar Indonesia)
    assert cfg.margin_bottom == 3.0  # cm
    assert cfg.font_name == "Times New Roman"
    assert cfg.font_size == 12
    assert cfg.line_spacing == 1.5
    assert cfg.first_line_indent == 1.25  # cm


def test_makalah_config():
    cfg = DocumentConfig.makalah()
    assert cfg.has_abstrak_en is False


def test_skripsi_config():
    cfg = DocumentConfig.skripsi()
    assert cfg.has_abstrak_en is True
