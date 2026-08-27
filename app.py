import io
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches

TEMPLATE_PATH = Path("modelo/Evidencia_de_Testes__CORSAN__Modelo.docx")

QA_NAMES = ["Nome QA 1", "Nome QA 2", "Nome QA 3"]

st.set_page_config(page_title="Gerador de Evidência de Testes", layout="centered")
st.title("Gerador de Evidência de Testes - CORSAN")


def get_unique_cells(row):
    """Remove células repetidas causadas por merge horizontal, mantendo a ordem."""
    unique = []
    seen_ids = set()
    for cell in row.cells:
        if id(cell._tc) not in seen_ids:
            seen_ids.add(id(cell._tc))
            unique.append(cell)
    return unique


def set_cell_text(cell, texto):
    """Escreve o texto na célula preservando a formatação do primeiro run, se existir."""
    if cell.paragraphs and cell.paragraphs[0].runs:
        run = cell.paragraphs[0].runs[0]
        run.text = texto
        for paragraph in cell.paragraphs[1:]:
            for run_extra in paragraph.runs:
                run_extra.text = ""
    else:
        cell.text = texto


def fill_labeled_field(document, label, valor):
    """Procura o rótulo em qualquer tabela do documento e escreve o valor na célula seguinte."""
    for table in document.tables:
        for row in table.rows:
            cells = get_unique_cells(row)
            for idx, cell in enumerate(cells):
                if cell.text.strip().upper() == label.upper():
                    if idx + 1 < len(cells):
                        set_cell_text(cells[idx + 1], valor)
                    return True
    return False


def add_screenshots(document, screenshots):
    """Insere as capturas de tela no final do documento, uma por página."""
    if not screenshots:
        return

    document.add_page_break()
    titulo = document.add_paragraph()
    run = titulo.add_run("Evidências (Prints)")
    run.bold = True

    for i, arquivo in enumerate(screenshots, start=1):
        document.add_paragraph(f"Print {i}: {arquivo.name}")
        imagem = io.BytesIO(arquivo.getvalue())
        document.add_picture(imagem, width=Inches(6))
        if i < len(screenshots):
            document.add_page_break()


def fill_document(template_bytes, valores, screenshots):
    document = Document(io.BytesIO(template_bytes))

    for label, valor in valores.items():
        if valor:
            fill_labeled_field(document, label, valor)

    add_screenshots(document, screenshots)

    buffer = io.BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


if not TEMPLATE_PATH.exists():
    st.error(f"Modelo não encontrado em {TEMPLATE_PATH}. Coloque o arquivo .docx nesse caminho.")
    st.stop()

template_bytes = TEMPLATE_PATH.read_bytes()

arquivo_xlsx = st.file_uploader("Envie o arquivo .xlsx do Caderno de Testes", type=["xlsx"])

if arquivo_xlsx:
    df = pd.read_excel(arquivo_xlsx)

    # Cenário vem da coluna "Processos"
    processos = df["Processos"].dropna().unique().tolist()
    if len(processos) == 1:
        cenario = processos[0]
        st.write(f"Cenário: **{cenario}**")
    else:
        cenario = st.selectbox("Cenário", processos)

    # Caso de Teste vem da coluna "Caso de Teste"
    casos_teste = df["Caso de Teste"].dropna().unique().tolist()
    if len(casos_teste) == 1:
        caso_teste = casos_teste[0]
        st.write(f"Caso de Teste: **{caso_teste}**")
    else:
        caso_teste = st.selectbox("Caso de Teste", casos_teste)

    qa = st.selectbox("QA", QA_NAMES)
    dev = st.text_input("DEV")
    ipc = st.text_input("IPC")
    squad = st.text_input("Squad", value="CORSAN")

    st.divider()
    st.subheader("Prints de tela")

    if "num_prints" not in st.session_state:
        st.session_state.num_prints = 1

    col_add, col_remove = st.columns(2)
    with col_add:
        if st.button("+ Adicionar print"):
            st.session_state.num_prints += 1
    with col_remove:
        if st.session_state.num_prints > 1 and st.button("- Remover último print"):
            st.session_state.num_prints -= 1

    screenshots = []
    for i in range(st.session_state.num_prints):
        arquivo_print = st.file_uploader(
            f"Print {i + 1}",
            type=["png", "jpg", "jpeg"],
            key=f"print_{i}",
        )
        if arquivo_print is not None:
            screenshots.append(arquivo_print)
            st.image(arquivo_print, width=250)

    st.divider()

    if st.button("Gerar documento"):
        valores = {
            "Cenário": cenario,
            "Caso de Teste": caso_teste,
            "QA": qa,
            "DEV": dev,
            "IPC": ipc,
            "Squad": squad,
        }

        resultado = fill_document(template_bytes, valores, screenshots)

        nome_arquivo = f"Evidencia_de_Testes_{caso_teste.split(' - ')[0]}.docx"

        st.success("Documento gerado com sucesso.")
        st.download_button(
            label="Baixar documento",
            data=resultado,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
else:
    st.info("Envie o arquivo .xlsx do Caderno de Testes para começar.")
