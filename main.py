import streamlit as st
import re
import fitz  # PyMuPDF
import pandas as pd
from docx import Document
from io import BytesIO

# Configurar a interface do Streamlit
st.title("Busca de Palavras em Relatórios PDF")
st.write("Insira até 3 palavras para buscar em documentos PDF e carregue o PDF para análise.")

# Entrada do usuário: palavras de busca
user_input = st.text_input("Digite até 3 palavras separadas por vírgulas:")
if user_input:
    palavras = [palavra.strip() for palavra in user_input.split(",")][:3]
else:
    palavras = []

# Upload de arquivos PDF
uploaded_file = st.file_uploader("Carregue um arquivo PDF", type="pdf")

if st.button("Iniciar Busca") and uploaded_file and palavras:
    results_text = []
    word_frequencies = {}

    # Processar o arquivo carregado
    pdf_file_path = uploaded_file.name
    pdf_doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")

    # Processar cada página no PDF
    for page_num, page in enumerate(pdf_doc, start=1):
        text = page.get_text("text")

        # Procurar cada palavra de busca
        for palavra in palavras:
            palavra_regex = re.compile(r'\b' + re.escape(palavra) + r'\b', re.IGNORECASE)
            if palavra not in word_frequencies:
                word_frequencies[palavra] = 0

            for match in palavra_regex.finditer(text):
                start = max(match.start() - 500, 0)
                end = min(match.end() + 500, len(text))
                context = text[start:end]

                results_text.append({
                    'Arquivo': pdf_file_path,
                    'Palavra': palavra,
                    'Página': page_num,
                    'Ocorrência': context
                })

                word_frequencies[palavra] += 1

    # Exibir os resultados
    if results_text:
        st.success("Busca concluída! Abaixo estão os resultados encontrados.")
        for result in results_text:
            st.write(f"**Arquivo**: {result['Arquivo']}")
            st.write(f"**Palavra**: {result['Palavra']}")
            st.write(f"**Página**: {result['Página']}")
            st.write(f"**Ocorrência**: {result['Ocorrência']}")
            st.write("---")

        # Criar arquivos de resultados para download
        # Frequências em Excel
        df_freq = pd.DataFrame(list(word_frequencies.items()), columns=['Palavra', 'Frequência'])
        excel_buffer = BytesIO()
        df_freq.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)

        # Ocorrências em Word
        doc = Document()
        doc.add_heading('Ocorrências de Palavras-Chave', 0)
        for result in results_text:
            doc.add_paragraph(f"Arquivo: {result['Arquivo']}")
            doc.add_paragraph(f"Palavra: {result['Palavra']}")
            doc.add_paragraph(f"Página: {result['Página']}")
            doc.add_paragraph(f"Ocorrência: {result['Ocorrência']}")
            doc.add_paragraph("\n")

        word_buffer = BytesIO()
        doc.save(word_buffer)
        word_buffer.seek(0)

        # Botões para download
        st.download_button("Baixar Frequências (Excel)", excel_buffer, file_name="frequencias_palavras.xlsx")
        st.download_button("Baixar Ocorrências (Word)", word_buffer, file_name="ocorrencias_palavras.docx")
    else:
        st.warning("Nenhuma ocorrência encontrada.")
else:
    if not uploaded_file:
        st.warning("Por favor, carregue um arquivo PDF.")
    if not palavras:
        st.warning("Por favor, insira até 3 palavras para busca.")
