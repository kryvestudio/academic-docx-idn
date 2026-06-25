class Appendix:
    def __init__(self, items: list[dict] | None = None):
        self.items = items or []

    def render(self, doc):
        doc.add_heading("LAMPIRAN", level=1)

        for item in self.items:
            judul = item.get("judul")
            if not judul:
                raise ValueError(f"Lampiran tidak punya key 'judul'")
            doc.add_heading(judul, level=2)
            for text in item.get("isi", []):
                doc.add_paragraph(text)
