from docx_idn.core.document import DocxDocument
from docx_idn.core.config import DocumentConfig
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
        institusi: str = "",
        fakultas: str = "",
        bab: list[dict] | None = None,
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
        self.judul = judul

        # Kata Pengantar (optional)
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

        # TOC (field code - auto-update in Word)
        if config.has_daftar_isi:
            toc = TableOfContents()
            toc.render(self.doc)

        # Chapters
        if bab:
            for i, b in enumerate(bab, start=1):
                judul = b.get("judul")
                if not judul:
                    raise ValueError(f"Bab index {i-1} tidak punya key 'judul'")
                chapter = Chapter(judul=judul, sections=b.get("sections"))
                chapter.render(self.doc, number=i)
                # Page break between chapters (but not after last chapter if bibliography follows)
                if i < len(bab):
                    self.doc.add_page_break()

        # Bibliography - always on new page
        if daftar_pustaka:
            self.doc.add_page_break()
            bib = Bibliography(entries=daftar_pustaka)
            bib.render(self.doc)

        # Appendix
        if lampiran and config.has_lampiran:
            self.doc.add_page_break()
            appendix = Appendix(items=lampiran)
            appendix.render(self.doc)

    def save(self, path: str):
        self.doc.save(path)
