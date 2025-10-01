
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
        # timestamp (esquerda)
        if self.print_ts:
            self.setFont("Helvetica", 9)
            self.drawString(12 * mm, y, f"Impresso em: {self.print_ts}")
        # número da página (direita)
        self.setFont("Helvetica", 9)
        texto = f"{self._pageNumber}/{page_count}"
        tw = self.stringWidth(texto, "Helvetica", 9)
        self.drawString(w - 12 * mm - tw, y, texto)

# Estilos
base_styles = getSampleStyleSheet()
style_title = ParagraphStyle(
    "TituloCentro",
    parent=base_styles["Title"],
    alignment=TA_CENTER,
    fontSize=14,
    leading=18,
    spaceAfter=12,
)
style_item = ParagraphStyle(
    "ItemJustificado",
    parent=base_styles["Normal"],
    alignment=TA_JUSTIFY,
    fontSize=10,
    leading=14,
    spaceAfter=4,   # menos espaçamento entre itens
)

# Regex para identificar URLs
URL_RE = re.compile(r"(https?://[^\s<>]+)", flags=re.IGNORECASE)

def to_clickable(texto: str) -> str:
    """
    Converte URLs em: [clique aqui para acessar] clicável,
    azul e sublinhado. Mantém o restante do texto escapado.
    """
    parts, last = [], 0
    for m in URL_RE.finditer(texto):
        parts.append(xml_escape(texto[last:m.start()]))
        url = m.group(0).rstrip(').,;')
        parts.append(
            f'<font color="blue"><u>'
            f'<link href="{url}">[clique aqui para acessar]</link>'
            f'</u></font>'
        )
        last = m.end()
    parts.append(xml_escape(texto[last:]))
    return "".join(parts)

# ====== UI ======
st.title("Gerador de PDFs para Projetos")
st.write("Faça o upload de um arquivo Excel com os dados dos projetos. Um PDF será gerado para cada linha da planilha.")

arquivo = st.file_uploader("Escolha o arquivo .xlsx", type="xlsx")

# ====== Geração do PDF ======
def gerar_pdf(dados: pd.Series) -> BytesIO:
    buffer = BytesIO()
    ts = datetime.now().strftime("%d/%m/%Y %H:%M")

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        topMargin=20 * mm,
        bottomMargin=20 * mm,
        title="Relatório - Plataforma Leonardo",
        author="Plataforma Leonardo",
    )

    elementos = []
    elementos.append(Paragraph(TITULO, style_title))
    elementos.append(Spacer(1, 6))

    # Evitar duplicação
    vistos = set()

    for coluna, valor in dados.items():
        if pd.notna(valor) and str(valor).strip() != "":
            raw = str(valor)
            texto_valor = to_clickable(raw)
            html = f"<b>{xml_escape(str(coluna))}:</b> {texto_valor}"

            if html not in vistos:  # não repete parágrafos iguais
                elementos.append(Paragraph(html, style_item))
                vistos.add(html)

    def _canvasmaker(*args, **kwargs):
        kwargs["print_ts"] = ts
        return NumberedCanvas(*args, **kwargs)

    doc.build(elementos, canvasmaker=_canvasmaker)
    buffer.seek(0)
    return buffer

# ====== Processamento do Excel ======
if arquivo:
    try:
        df = pd.read_excel(arquivo)
    except Exception as e:
        st.error(f"Não consegui ler o Excel: {e}")
    else:
        st.success(f"{len(df)} projetos encontrados na planilha.")

        arquivos_pdfs = []
        for i, linha in df.iterrows():
            nome_base = linha.get("Nome", f"projeto_{i+1}")
            nome = str(nome_base).strip() or f"projeto_{i+1}"
            nome = (nome
                    .replace("/", "-")
                    .replace("\\", "-")
                    .replace(" ", "_"))
            buffer_pdf = gerar_pdf(linha)
            arquivos_pdfs.append((f"{nome}.pdf", buffer_pdf.read()))

        # ZIP para download
        buffer_zip = BytesIO()
        with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
            for nome_arquivo, conteudo in arquivos_pdfs:
                zipf.writestr(nome_arquivo, conteudo)
        buffer_zip.seek(0)

        st.download_button(
            label="📦 Baixar todos os PDFs em um ZIP",
            data=buffer_zip,
            file_name="projetos_pdfs.zip",
            mime="application/zip"
        )

        st.write("---")
        st.write("### Visualizar PDFs individualmente")
        for nome_arquivo, conteudo in arquivos_pdfs:
            st.download_button(
                label=f"⬇️ Baixar {nome_arquivo}",
                data=conteudo,
                file_name=nome_arquivo,
                mime="application/pdf"
            )
