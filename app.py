# streamlit_app.py
# Este script pode ser executado no Streamlit Cloud

import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors
import zipfile
import streamlit as st

st.set_page_config(page_title="Gerador de PDFs para Projetos", layout="centered")

# Título da interface
st.title("Gerador de PDFs para Projetos")
st.write("Faça o upload de um arquivo Excel com os dados dos projetos. Um PDF será gerado para cada linha da planilha.")

# Upload do arquivo Excel
arquivo = st.file_uploader("Escolha o arquivo .xlsx", type="xlsx")

# Função para gerar PDF a partir de uma linha do DataFrame
def gerar_pdf(dados):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    estilos = getSampleStyleSheet()
    elementos = []

    # Estilo do título centralizado
    titulo_style = ParagraphStyle(
        name="TituloCentralizado",
        parent=estilos["Title"],
        alignment=TA_CENTER,
        fontSize=14,
        textColor=colors.HexColor("#000000")
    )

    # Adicionar título principal
    titulo = "Plataforma Leonardo - Disciplina de Ética em Pesquisa - PPGCiMH - FEFF/UFAM"
    elementos.append(Paragraph(titulo, titulo_style))
    elementos.append(Spacer(1, 24))

    # Adicionar os dados preenchidos (ignorando colunas vazias e formatando links como "clique aqui")
    for coluna, valor in dados.items():
        if pd.notna(valor) and str(valor).strip() != "":
            valor_str = str(valor).strip()
            if valor_str.startswith("http://") or valor_str.startswith("https://"):
                texto = f"<b>{coluna}:</b> <a href='{valor_str}' color='blue'>clique aqui para acessar</a>"
            else:
                texto = f"<b>{coluna}:</b> {valor_str}"
            elementos.append(Paragraph(texto, estilos["Normal"]))
            elementos.append(Spacer(1, 12))

    doc.build(elementos)
    buffer.seek(0)
    return buffer

# Processamento
if arquivo:
    df = pd.read_excel(arquivo)
    st.success(f"{len(df)} projetos encontrados na planilha.")

    arquivos_pdfs = []

    for i, linha in df.iterrows():
        nome = linha.get("Nome", f"projeto_{i+1}").replace(" ", "_")
        buffer_pdf = gerar_pdf(linha)
        arquivos_pdfs.append((f"{nome}.pdf", buffer_pdf.read()))

    # Geração do ZIP para download
    buffer_zip = BytesIO()
    with zipfile.ZipFile(buffer_zip, "w") as zipf:
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
