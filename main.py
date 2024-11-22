import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
from io import BytesIO

# Ajustar o limite de upload de arquivos
st.set_option('server.maxUploadSize', 200)  # Ajuste conforme necessário (em MB)

st.title("Busca de Palavras em Relatórios PDF - Arquivos Grandes")

# Entrada do usuário
user_input = st.text_input("Digite até 3 palavras separadas por vírgulas:")
if user_input:
    palavras = [palavra.strip() for palavra in user_input.split(",")][:3]
else:
    palavras = []

# Upload do PDF
uploaded_file = st.file_uploader("Carregue um arquivo PDF", type="pdf")

if st.button("Iniciar Busca") and uploaded_file and palavras:
    pdf_doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    results_text = []
    word_frequencies = {}

    for page_num, page in enumerate(pdf_doc, start=1):
        text = page.get_text("text")
        for palavra in palavras:
            palavra_regex = re.compile(r'\b' + re.escape(palavra) + r'\b', re.IGNORECASE)
            if palavra not in word_frequencies:
                word_frequencies[palavra] = 0

            for match in palavra_regex.finditer(text):
                start = max(match.start() - 500, 0)
                end = min(match.end() + 500, len(text))
                context = text[start:end]

                results_text.append({
                    'Arquivo': uploaded_file.name,
                    'Palavra': palavra,
                    'Página': page_num,
                    'Ocorrência': context
                })
                word_frequencies[palavra] += 1

    if results_text:
        st.success("Busca concluída!")
        for result in results_text:
            st.write(f"**Arquivo**: {result['Arquivo']}")
            st.write(f"**Palavra**: {result['Palavra']}")
            st.write(f"**Página**: {result['Página']}")
            st.write(f"**Ocorrência**: {result['Ocorrência']}")
            st.write("---")
    else:
        st.warning("Nenhuma ocorrência encontrada.")

