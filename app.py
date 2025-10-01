import streamlit as st
import pandas as pd
import os
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas
from zipfile import ZipFile
import shutil

# Configurações iniciais
st.set_page_config(page_title="Gerador de PDFs para Projetos")
st.title("Gerador de PDFs para Projetos")
st.markdown("Faça o upload de um arquivo Excel com os dados dos projetos. Um PDF será gerado para cada linha da planilha.")

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Escolha o arquivo .xlsx", type="xlsx")

# Função para criar o PDF
def gerar_pdf(dados, nome_arquivo):
    if not os.path.exists("pdfs_gerados"):
        os.makedirs("pdfs_gerados")

    caminho_arquivo = os.path.join("pdfs_gerados", nome_arquivo)
    doc = SimpleDocTemplate(caminho_arquivo, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=80, bottomMargin=30)

    styles = getSampleStyleSheet()
    style_normal = styles["Normal"]
    style_just = ParagraphStyle(name='Justify', parent=style_normal, alignment=4, fontSize=10, leading=12)
    story = []

    # Título em caixa alta dentro de uma caixa
    story.append(Paragraph("<para align='center'><b>PLATAFORMA LEONARDO - DISCIPLINA DE ÉTICA EM PESQUISA - PPGCiMH - FEFF/UFAM</b></para>", styles['Title']))
    story.append(Spacer(1, 24))

    for chave, valor in dados.items():
        if pd.notna(valor):
            story.append(Paragraph(f"<b>{chave}:</b> {valor}", style_just))
            story.append(Spacer(1, 6))

    doc.build(story, canvasmaker=lambda filename, pagesize: RodapeCanvas(filename, pagesize))

# Classe para rodapé personalizado
class RodapeCanvas(Canvas):
    def __init__(self, *args, **kwargs):
        Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_footer()
            Canvas.showPage(self)
        Canvas.save(self)

    def draw_footer(self):
        self.setFont("Helvetica", 8)
        self.drawString(cm, 1.5 * cm, "Plataforma Leonardo")
        self.drawRightString(A4[0] - cm, 1.5 * cm, f"{self._pageNumber}")

# Processamento do arquivo
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    total = len(df)
    for i, (_, row) in enumerate(df.iterrows()):
        dados = row.dropna().to_dict()
        nome_pdf = f"projeto_{i+1}.pdf"
        gerar_pdf(dados, nome_pdf)

    # Criar arquivo ZIP
    zip_path = "pdfs_gerados.zip"
    with ZipFile(zip_path, 'w') as zipf:
        for root, _, files in os.walk("pdfs_gerados"):
            for file in files:
                zipf.write(os.path.join(root, file), arcname=file)

    st.success(f"{total} PDFs foram gerados e salvos na pasta 'pdfs_gerados'.")

    with open(zip_path, "rb") as f:
        st.download_button("📦 Baixar todos em .ZIP", f, file_name="pdfs_projetos.zip")

    st.markdown("### 📄 Baixar arquivos individualmente")
    for file in os.listdir("pdfs_gerados"):
        with open(os.path.join("pdfs_gerados", file), "rb") as f:
            st.download_button(f"Baixar {file}", f, file_name=file, key=file)

    # Limpeza opcional
    # shutil.rmtree("pdfs_gerados")
    # os.remove(zip_path)
