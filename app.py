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

    buffer = BytesIO()

    def header_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        canvas.drawString(30, 15, "Plataforma Leonardo")
        canvas.drawRightString(A4[0] - 30, 15, f"{canvas.getPageNumber()}")
        canvas.restoreState()

    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=60, bottomMargin=40)
    styles = getSampleStyleSheet()
    justified = ParagraphStyle(name="Justificado", parent=styles["Normal"], alignment=4)
    story = []

    titulo = Paragraph("<para align='center'><b><font size=14>PLATAFORMA LEONARDO - DISCIPLINA DE ÉTICA EM PESQUISA - PPGCIMH - FEFF/UFAM</font></b></para>", styles["Normal"])
    caixa_titulo = Table([[titulo]], colWidths=A4[0] - 60)
    caixa_titulo.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER')
    ]))
    story.append(caixa_titulo)
    story.append(Spacer(1, 20))

    for _, row in df.iterrows():
        for grupo in agrupamentos:
            data = []
            for col in grupo:
                valor = row[col]
                if isinstance(valor, str) and valor.startswith("http"):
                    texto = f"<a href='{valor}' color='blue'><u>Clique aqui para acessar</u></a>"
                else:
                    texto = str(valor)
                paragrafo = Paragraph(f"<b>{col}:</b> {texto}", justified)
                data.append([paragrafo])
            tabela = Table(data, colWidths=[A4[0] - 60])
            tabela.setStyle(TableStyle([
                ('BOX', (0, 0), (-1, -1), 1, colors.black),
                ('INNERGRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP')
            ]))
            story.append(tabela)
            story.append(Spacer(1, 10))
        story.append(PageBreak())

    doc.build(story, onLaterPages=header_footer, onFirstPage=header_footer)

    buffer.seek(0)
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr("relatorio.pdf", buffer.read())

    zip_buffer.seek(0)
    st.download_button(label="📥 Baixar PDF", data=zip_buffer, file_name="relatorio.zip", mime="application/zip")
