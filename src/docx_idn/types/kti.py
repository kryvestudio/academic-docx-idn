from docx_idn.types.base import BaseDocument
from docx_idn.core.config import DocumentConfig


class KTI(BaseDocument):
    def __init__(self, **kwargs):
        kwargs.setdefault("config", DocumentConfig.kti())
        super().__init__(**kwargs)
