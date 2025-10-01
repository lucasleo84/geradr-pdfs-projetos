import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
import streamlit as st
import os
import re

# Título da interface
st.title("Gerador de PDFs para Projetos")
st.write("Faça o upload de um arquivo Excel com os dados dos projetos. Um PDF será gerado para cada linha da planilha.")

# Upload do arquivo Excel
arquivo = st.file_uploader("Escolha o arquivo .xlsx", type="xlsx")

# Estilo de parágrafo justificado
estilos = getSampleStyleSheet()
justificado = ParagraphStyle(name='Justificado', parent=estilos['Normal'], alignment=TA_JUSTIFY)

# Função para gerar PDF a partir de uma linha do DataFrame
def gerar_pdf(dados, nome_arquivo):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elementos = []

    # Adiciona título centralizado em caixa com borda
    titulo = "PLATAFORMA LEONARDO - DISCIPLINA DE ÉTICA EM PESQUISA - PPGCIMH - FEFF/UFAM"
    elementos.append(Spacer(1, 12))
    elementos.append(Paragraph(f'<para align="center"><b>{titulo}</b></para>', estilos['Title']))
    elementos.append(Spacer(1, 12))

    # Adiciona informações do projeto, ignorando colunas vazias
    for coluna, valor in dados.items():
        if pd.notna(valor):
            valor = str(valor)
            if re.search(r'https?://', valor):
                valor = f'<a href="{valor}">[Clique aqui para acessar]</a>'
            texto = f"<b>{coluna}:</b> {valor}"
            elementos.append(Paragraph(texto, justificado))
            elementos.append(Spacer(1, 6))

    def rodape(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        canvas.drawString(2 * cm, 1.5 * cm, "Plataforma Leonardo")
        canvas.drawRightString(A4[0] - 2 * cm, 1.5 * cm, f"Página {doc.page}")
        canvas.restoreState()

    doc.build(elementos, onFirstPage=rodape, onLaterPages=rodape)

    with open(nome_arquivo, "wb") as f:
        f.write(buffer.getbuffer())

# Processamento
if arquivo:
    df = pd.read_excel(arquivo)
    pasta_saida = "pdfs_gerados"
    os.makedirs(pasta_saida, exist_ok=True)

    for i, linha in df.iterrows():
        nome = linha.get("Nome", f"projeto_{i+1}")
        nome = str(nome).strip().replace(" ", "_").replace("/", "-")
        nome_arquivo = os.path.join(pasta_saida, f"{nome}.pdf")
        gerar_pdf(linha, nome_arquivo)

    st.success(f"{len(df)} PDFs foram gerados e salvos na pasta '{pasta_saida}'.")

