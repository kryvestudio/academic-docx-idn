from docx_idn.types.base import BaseDocument
from docx_idn.core.config import DocumentConfig


class Skripsi(BaseDocument):
    def __init__(self, **kwargs):
        kwargs.setdefault("config", DocumentConfig.skripsi())
        super().__init__(**kwargs)
