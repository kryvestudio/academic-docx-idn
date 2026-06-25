from docx_idn.types.base import BaseDocument
from docx_idn.core.config import DocumentConfig


class Laporan(BaseDocument):
    def __init__(self, **kwargs):
        kwargs.setdefault("config", DocumentConfig.laporan())
        super().__init__(**kwargs)
