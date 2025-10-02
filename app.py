
import re
import zipfile
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from xml.sax.saxutils import escape as xml_escape

st.set_page_config(page_title="Gerador de PDFs para Projetos", layout="centered")

# ====== Configurações ======
TITULO = "PLATAFORMA LEONARDO - DISCIPLINA DE ÉTICA EM PESQUISA - PPGCIMH - FEFF/UFAM"

# --- util: nome base da coluna (sem sufixos .1, .2 etc.)
COL_BASE_RE = re.compile(r"\.\d+$")
def base_col(name: str) -> str:
    return COL_BASE_RE.sub("", str(name).strip())

# Canvas numerado com timestamp
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        self._saved_page_states = []
        self.print_ts = kwargs.pop("print_ts", None)
        super().__init__(*args, **kwargs)

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        super().showPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_footer(num_pages)
            super().showPage()
        super().save()

    def draw_footer(self, page_count):
        w, h = A4
        y = 10 * mm
        if self.print_ts:
            self.setFont("Helvetica", 9)
            self.drawString(12 * mm, y, f"Impresso em: {self.print_ts}")
        self.setFont("Helvetica", 9)
        texto = f"{self._pageNumber}/{page_count}"
        tw = self.string
