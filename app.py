
import streamlit as st
from docx import Document
from io import BytesIO
from datetime import date

orgaos = [
    "AGÊNCIA DE DEFESA AGROPECUÁRIA DO PARANÁ (ADAPAR)",
    "SEC. EST. SEGURANÇA PÚBLICA (SESP - CONSOLIDADOR)",
    "UNIV. EST. DE MARINGÁ (UEM)",
    "SECRETARIA DE ESTADO DA SAÚDE (SESA)",
    "SECRETARIA DE ESTADO DA FAZENDA (SEFA)",
    "SECRETARIA DE ESTADO DO ESPORTE - SEES",
    "UNIV. EST. DO OESTE DO PR (UNIOESTE)",
]

st.title("Gerador de Despachos Oficiais")

tipo = st.selectbox("Tipo de Ação", [
    "Cadastro de Veículos",
    "Inativação de Veículos",
    "Inativação de Veículos sem cadastro na Prime",
    "Cessão de Veículos entre órgãos"
])

numero_despacho = st.text_input("Nº do Despacho", value="001/2025")
numero_protocolo = st.text_input("Nº do Protocolo", value="12.345.678-9")
interessado = st.text_input("Interessado", value="SEAP")
data_despacho = st.date_input("Data", value=date.today())
orgao_extenso = st.selectbox("Interessado por extenso", sorted(orgaos))
placa = st.text_input("Placa do Veículo", value="ABC-1234")

modelos = {
    "Cadastro de Veículos": "modelos/cadastro.docx",
    "Inativação de Veículos": "modelos/inativacao.docx",
    "Inativação de Veículos sem cadastro na Prime": "modelos/inativacao_sem_prime.docx",
    "Cessão de Veículos entre órgãos": "modelos/cessao.docx"
}

if st.button("Gerar Despacho"):
    doc = Document(modelos[tipo])

    for p in doc.paragraphs:
        if "XXX/2025" in p.text:
            p.text = p.text.replace("XXX/2025", numero_despacho)
        if "XX/XX/2025" in p.text or "XX/XX/XXXX" in p.text:
            p.text = p.text.replace("XX/XX/2025", data_despacho.strftime("%d/%m/%Y")).replace("XX/XX/XXXX", data_despacho.strftime("%d/%m/%Y"))
        if "XX.XXX.XXX-X" in p.text:
            p.text = p.text.replace("XX.XXX.XXX-X", numero_protocolo)
        if "NONONONO" in p.text or "NONONONONO" in p.text:
            p.text = p.text.replace("NONONONO", interessado).replace("NONONONONO", interessado)
        if "À Nonononono" in p.text:
            p.text = p.text.replace("À Nonononono", f"À {orgao_extenso}")
        if "placa XXX-XXXX" in p.text:
            p.text = p.text.replace("placa XXX-XXXX", f"placa {placa}")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.download_button("📄 Baixar Word", buffer, file_name="despacho.docx")
