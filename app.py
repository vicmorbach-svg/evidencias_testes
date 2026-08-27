import io
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
import re

TEMPLATE_PATH = Path("modelo/Evidencia de Testes - CORSAN - Modelo.docx")

QA_NAMES = ["Victor Morbach", "Aline Rodrigues Vieira Pinto"]

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


def fill_status(document, status_valor):
    """Localiza o parágrafo 'STATUS:' e acrescenta o valor escolhido na mesma linha."""
    for paragraph in document.paragraphs:
        if "STATUS" in paragraph.text.strip().upper():
            run = paragraph.add_run(f" {status_valor}")
            run.bold = False
            return True
    return False


def add_screenshots_after_table(document, screenshots):
    """Insere os prints logo após a tabela, um em seguida do outro, sem quebras de página."""
    if not document.tables or not screenshots:
        return

    anchor = document.tables[0]._tbl

    for idx, item in enumerate(screenshots, start=1):
        arquivo = item["arquivo"]
        legenda = item.get("legenda", "").strip()
        arquivo.seek(0)

        img_paragraph = document.add_paragraph()
        run = img_paragraph.add_run()
        run.add_picture(arquivo, width=Inches(6))
        anchor.addnext(img_paragraph._p)
        anchor = img_paragraph._p

        cap_paragraph = document.add_paragraph()
        cap_run = cap_paragraph.add_run(f"Print {idx}: {legenda}" if legenda else f"Print {idx}")
        cap_run.italic = True
        cap_run.font.size = Pt(9)
        anchor.addnext(cap_paragraph._p)
        anchor = cap_paragraph._p


def sanitize_filename(texto):
    """Remove caracteres inválidos para nome de arquivo."""
    texto = re.sub(r'[\\/*?:"<>|]', "", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def fill_document(template_bytes, valores, screenshots):
    document = Document(io.BytesIO(template_bytes))

    for label, valor in valores.items():
        if label != "Status":
            fill_labeled_field(document, label, valor)

    fill_status(document, valores["Status"])
    add_screenshots_after_table(document, screenshots)

    buffer = io.BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


import io  # mantenha este import junto aos demais no topo do arquivo

if not TEMPLATE_PATH.exists():
    st.error("Modelo .docx não encontrado na pasta 'modelo/'.")
    st.stop()

template_bytes = TEMPLATE_PATH.read_bytes()

arquivo_xlsx = st.file_uploader("Envie o Caderno de Testes (.xlsx)", type=["xlsx"])

if arquivo_xlsx is not None:
    df = pd.read_excel(arquivo_xlsx, sheet_name="plano de teste")

    cenarios = df["Processos"].dropna().unique().tolist()
    if len(cenarios) == 1:
        cenario = cenarios[0]
        st.write(f"**Cenário:** {cenario}")
    else:
        cenario = st.selectbox("Cenário", cenarios)

    df_filtrado = df[df["Processos"] == cenario]
    casos_teste = df_filtrado["Caso de Teste"].dropna().unique().tolist()
    if len(casos_teste) == 1:
        caso_teste = casos_teste[0]
        st.write(f"**Caso de Teste:** {caso_teste}")
    else:
        caso_teste = st.selectbox("Caso de Teste", casos_teste, key=f"caso_teste_select_{cenario}")

    if len(QA_NAMES) == 1:
        qa = QA_NAMES[0]
    else:
        qa = st.selectbox("QA", QA_NAMES)

    dev = st.text_input("DEV")
    ipc = st.text_input("IPC")
    squad = st.text_input("Squad", value="CORSAN")
    status = st.selectbox("Status", STATUS_OPCOES)

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
        st.markdown(f"**Print {i + 1}**")
        arquivo_print = st.file_uploader(
            f"Selecione a imagem do print {i + 1}",
            type=["png", "jpg", "jpeg"],
            key=f"print_{i}",
        )
        legenda = st.text_input(f"Legenda do print {i + 1} (opcional)", key=f"legenda_{i}")
        if arquivo_print is not None:
            screenshots.append({"arquivo": arquivo_print, "legenda": legenda})
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
            "Status": status,
        }

        resultado = fill_document(template_bytes, valores, screenshots)

        nome_arquivo = sanitize_filename(
            f"Evidencias de teste {squad} {cenario} {caso_teste}"
        ) + ".docx"

        st.success("Documento gerado com sucesso.")
        st.download_button(
            label="Baixar documento",
            data=resultado,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
else:
    st.info("Envie o arquivo .xlsx do Caderno de Testes para começar.")
