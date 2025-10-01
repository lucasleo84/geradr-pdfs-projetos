import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import zipfile
import streamlit as st

st.set_page_config(page_title="Gerador de PDFs para Projetos", layout="centered")
st.title("Gerador de PDFs para Projetos")

uploaded_file = st.file_uploader("Escolha um arquivo Excel (.xlsx)", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.dropna(axis=1, how="all", inplace=True)

    colunas = df.columns.tolist()
    agrupamentos = []
    grupo = []
    for col in colunas:
        if col.lower().startswith("grupo") and grupo:
            agrupamentos.append(grupo)
            grupo = [col]
        else:
            grupo.append(col)
    if grupo:
        agrupamentos.append(grupo)

    class CustomCanvas(canvas.Canvas):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.pages = []

        def showPage(self):
            self.pages.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            page_count = len(self.pages)
            for page_num, page in enumerate(self.pages, start=1):
                self.__dict__.update(page)
                self.draw_footer(page_num)
                canvas.Canvas.showPage(self)
            canvas.Canvas.save(self)

        def draw_footer(self, page_num):
            self.setFont('Helvetica', 8)
            self.drawString(30, 15, "Plataforma Leonardo")
            self.drawRightString(A4[0] - 30, 15, f"{page_num}")

    def desenhar_titulo(canvas, doc):
        canvas.saveState()
        titulo = "PLATAFORMA LEONARDO - DISCIPLINA DE ÉTICA EM PESQUISA - PPGCIMH - FEFF/UFAM"
        canvas.setFont("Helvetica-Bold", 12)
        largura = A4[0] - 60
        altura = A4[1] - 80
        canvas.rect(30, altura - 20, largura, 40, stroke=1, fill=0)
        canvas.drawCentredString(A4[0] / 2, altura, titulo)
        canvas.restoreState()

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=100, bottomMargin=40)
    styles = getSampleStyleSheet()
    justified = ParagraphStyle(name="Justificado", parent=styles["Normal"], alignment=4)
    story = []

    for _, row in df.iterrows():
        story.append(Spacer(1, 60))  # espaço para o título desenhado no canvas
        for grupo in agrupamentos:
            data = []
            for col in grupo:
                valor = row[col]
                if pd.isna(valor) or str(valor).strip() == "":
                    continue
                texto = str(valor).strip()
                if texto.startswith("http"):
                    texto = f"<a href='{texto}' color='blue'><u>Clique aqui para acessar</u></a>"
                paragrafo = Paragraph(f"<b>{col}:</b> {texto}", justified)
                data.append([paragrafo])
            if data:
                tabela = Table(data, colWidths=[A4[0] - 60])
                tabela.setStyle(TableStyle([
                    ('BOX', (0, 0), (-1, -1), 1, colors.black),
                    ('INNERGRID', (0, 0), (-1, -1), 1, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP')
                ]))
                story.append(tabela)
                story.append(Spacer(1, 10))
        story.append(PageBreak())

    doc.build(story, onFirstPage=desenhar_titulo, onLaterPages=desenhar_titulo, canvasmaker=CustomCanvas)

    buffer.seek(0)
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr("relatorio.pdf", buffer.read())

    zip_buffer.seek(0)
    st.download_button(label="📥 Baixar PDF", data=zip_buffer, file_name="relatorio.zip", mime="application/zip")
