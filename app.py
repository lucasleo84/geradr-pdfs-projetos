
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
        tw = self.stringWidth(texto, "Helvetica", 9)
        self.drawString(w - 12 * mm - tw, y, texto)

# ====== Estilos ======
base_styles = getSampleStyleSheet()
style_title = ParagraphStyle(
    "TituloCentro",
    parent=base_styles["Title"],
    alignment=TA_CENTER,
    fontSize=14,
    leading=16,
    spaceAfter=8,
)
style_item = ParagraphStyle(
    "ItemJustificado",
    parent=base_styles["Normal"],
    alignment=TA_JUSTIFY,
    fontSize=10,
    leading=12,   # linha simples
    spaceAfter=2, # respiro pequeno
)

# ====== Regex para identificar URLs ======
URL_RE = re.compile(r"(https?://[^\s<>]+)", flags=re.IGNORECASE)

def to_clickable(texto: str) -> str:
    """Converte URLs em [clique aqui para acessar] clicável, azul e sublinhado."""
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

# ====== Função de Geração do PDF ======
def gerar_pdf(dados: pd.Series) -> BytesIO:
    """
    Regras anti-duplicação:
    - Agrupa colunas por nome-base (sem .1, .2 ...).
    - Se TODAS as colunas de um mesmo nome-base têm o MESMO valor (após strip),
      renderiza APENAS UMA vez (com o nome-base).
    - Se existirem valores DISTINTOS, renderiza CADA coluna com seu rótulo original
      (ex.: Pesquisador, Pesquisador.1), preservando múltiplos pesquisadores/CPFs.
    """
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
    elementos.append(Spacer(1, 4))

    # 1) Agrupar colunas por nome-base
    grupos = {}  # base -> list of (col_name, value_str)
    for col, val in dados.items():
        if pd.isna(val):
            continue
        raw = str(val).strip()
        if raw == "":
            continue
        bname = base_col(col)
        grupos.setdefault(bname, []).append((str(col), raw))

    # 2) Renderização com regra de deduplicação
    for bname in grupos:
        colvals = grupos[bname]  # lista de (nome_col, valor)
        # valores únicos (normalizados) para detectar duplicata real
        valores_unicos = list({v for _, v in colvals})

        if len(valores_unicos) == 1:
            # Todos os campos duplicados desse "bname" têm o MESMO valor -> mostra 1x
            valor = valores_unicos[0]
            html = f"<b>{xml_escape(bname)}:</b> {to_clickable(valor)}"
            elementos.append(Paragraph(html, style_item))
        else:
            # Há valores distintos -> mostra cada um com o rótulo COMPLETO (preserva .1, .2)
            for colname, valor in colvals:
                html = f"<b>{xml_escape(colname)}:</b> {to_clickable(valor)}"
                elementos.append(Paragraph(html, style_item))

    # 3) Gerar com rodapé numerado e timestamp
    def _canvasmaker(*args, **kwargs):
        kwargs["print_ts"] = ts
        return NumberedCanvas(*args, **kwargs)

    doc.build(elementos, canvasmaker=_canvasmaker)
    buffer.seek(0)
    return buffer

# ====== Interface Streamlit ======
st.title("Gerador de PDFs para Projetos")
st.write("Faça o upload de um arquivo Excel com os dados dos projetos. Um PDF será gerado para cada linha da planilha.")

arquivo = st.file_uploader("Escolha o arquivo .xlsx", type="xlsx")

if arquivo:
    try:
        df = pd.read_excel(arquivo)
        # NÃO removemos colunas duplicadas aqui, porque você precisa de (Pesquisador, Pesquisador.1) etc.
        # A deduplicação “inteligente” acontece dentro do gerar_pdf() por nome-base + valor.
    except Exception as e:
        st.error(f"Não consegui ler o Excel: {e}")
    else:
        st.success(f"{len(df)} projetos encontrados na planilha.")

        arquivos_pdfs = []
        for i, linha in df.iterrows():
            nome_base = (
                linha.get("Nome")
                or linha.get("Nome do Pesquisador")
                or linha.get("Título")
                or f"projeto_{i+1}"
            )
            nome = str(nome_base).strip() or f"projeto_{i+1}"
            nome = (nome
                    .replace("/", "-")
                    .replace("\\", "-")
                    .replace(" ", "_"))
            buffer_pdf = gerar_pdf(linha)
            arquivos_pdfs.append((f"{nome}.pdf", buffer_pdf.read()))

        # ZIP (key único)
        buffer_zip = BytesIO()
        with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
            for nome_arquivo, conteudo in arquivos_pdfs:
                zipf.writestr(nome_arquivo, conteudo)
        buffer_zip.seek(0)

        st.download_button(
            label="📦 Baixar todos os PDFs em um ZIP",
            data=buffer_zip,
            file_name="projetos_pdfs.zip",
            mime="application/zip",
            key="btn_download_zip_all"
        )

        st.write("---")
        st.write("### Visualizar PDFs individualmente")
        for i, (nome_arquivo, conteudo) in enumerate(arquivos_pdfs):
            st.download_button(
                label=f"⬇️ Baixar {nome_arquivo}",
                data=conteudo,
                file_name=nome_arquivo,
                mime="application/pdf",
                key=f"btn_download_{i}_{nome_arquivo}"
            )
