import io
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ------------------------------------------------------------------
# CONFIGURAÇÕES
# ------------------------------------------------------------------
TEMPLATE_PATH = Path("modelo/Evidencia_de_Testes_CORSAN_Modelo.docx")

QA_NAMES = [
    "Victor Morbach",
    "Aline Rodrigues Vieira Pinto",
]

SHEET_NAME = "plano de teste"
COL_PROCESSO = "Processos"
COL_CASO_TESTE = "Caso de Teste"

IMAGEM_LARGURA_POL = 6.0  # largura da imagem no documento, em polegadas

# ------------------------------------------------------------------
# FUNÇÕES AUXILIARES
# ------------------------------------------------------------------
def get_unique_cells(row):
    """Remove repetições de célula causadas por merge horizontal."""
    seen = set()
    cells = []
    for cell in row.cells:
        if id(cell) not in seen:
            seen.add(id(cell))
            cells.append(cell)
    return cells


def normalize(text: str) -> str:
    return text.strip().lower().rstrip(":").replace("*", "")


def set_cell_text(cell, text: str):
    """Escreve o texto na célula preservando a formatação existente."""
    if cell.paragraphs:
        paragraph = cell.paragraphs[0]
        if paragraph.runs:
            paragraph.runs[0].text = text
            for extra_run in paragraph.runs[1:]:
                extra_run.text = ""
        else:
            paragraph.add_run(text)
        for extra_paragraph in cell.paragraphs[1:]:
            extra_paragraph.text = ""
    else:
        cell.text = text


def fill_fields(doc: Document, values: dict):
    field_map = {normalize(k): v for k, v in values.items()}

    for table in doc.tables:
        for row in table.rows:
            unique_cells = get_unique_cells(row)
            for idx, cell in enumerate(unique_cells):
                label = normalize(cell.text)
                if label in field_map and idx + 1 < len(unique_cells):
                    set_cell_text(unique_cells[idx + 1], field_map[label])


def add_evidencias(doc: Document, imagens: list):
    """Adiciona um título 'Evidências' e as imagens (com legenda opcional) ao final do documento."""
    if not imagens:
        return

    doc.add_page_break()

    titulo = doc.add_paragraph()
    run_titulo = titulo.add_run("Evidências")
    run_titulo.bold = True
    run_titulo.font.size = Pt(14)

    for item in imagens:
        nome = item["nome"]
        legenda = item["legenda"]
        conteudo = item["conteudo"]

        doc.add_paragraph()  # espaçamento
        paragrafo_imagem = doc.add_paragraph()
        paragrafo_imagem.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_imagem = paragrafo_imagem.add_run()
        run_imagem.add_picture(io.BytesIO(conteudo), width=Inches(IMAGEM_LARGURA_POL))

        texto_legenda = legenda.strip() if legenda and legenda.strip() else nome
        paragrafo_legenda = doc.add_paragraph()
        paragrafo_legenda.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_legenda = paragrafo_legenda.add_run(texto_legenda)
        run_legenda.italic = True
        run_legenda.font.size = Pt(10)


def fill_document(template_bytes: bytes, values: dict, imagens: list) -> bytes:
    doc = Document(io.BytesIO(template_bytes))
    fill_fields(doc, values)
    add_evidencias(doc, imagens)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


@st.cache_data
def load_excel(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=SHEET_NAME)


# ------------------------------------------------------------------
# APP
# ------------------------------------------------------------------
st.set_page_config(page_title="Gerador de Evidência de Testes", layout="centered")
st.title("Gerador de Evidência de Testes - CORSAN")

xlsx_file = st.file_uploader(
    "Envie o Caderno de Testes (.xlsx)", type=["xlsx"], key="xlsx_uploader"
)

template_file = None
if not TEMPLATE_PATH.exists():
    st.warning(
        "Modelo padrão não encontrado no servidor. Envie o arquivo .docx do modelo."
    )
    template_file = st.file_uploader(
        "Envie o modelo do documento (.docx)", type=["docx"], key="docx_uploader"
    )

if xlsx_file is not None:
    df = load_excel(xlsx_file.getvalue())

    casos_teste = sorted(df[COL_CASO_TESTE].dropna().unique())
    caso_teste = st.selectbox("Caso de Teste", casos_teste)

    processos_relacionados = (
        df.loc[df[COL_CASO_TESTE] == caso_teste, COL_PROCESSO].dropna().unique()
    )

    if len(processos_relacionados) == 1:
        cenario = processos_relacionados[0]
        st.text_input("Cenário", value=cenario, disabled=True)
    elif len(processos_relacionados) > 1:
        cenario = st.selectbox("Cenário", sorted(processos_relacionados))
    else:
        cenario = st.text_input("Cenário")

    qa = st.selectbox("QA", QA_NAMES)
    dev = st.text_input("DEV")
    ipc = st.text_input("IPC")
    squad = st.text_input("Squad", value="CORSAN")

    st.subheader("Prints de tela (evidências)")
    prints_upload = st.file_uploader(
        "Envie os prints de tela",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="prints_uploader",
    )

    imagens = []
    if prints_upload:
        for i, arquivo in enumerate(prints_upload):
            col1, col2 = st.columns([1, 2])
            with col1:
                st.image(arquivo, caption=arquivo.name, use_container_width=True)
            with col2:
                legenda = st.text_input(
                    f"Legenda para '{arquivo.name}'",
                    key=f"legenda_{i}",
                )
            imagens.append(
                {
                    "nome": arquivo.name,
                    "legenda": legenda,
                    "conteudo": arquivo.getvalue(),
                }
            )

    gerar = st.button("Gerar documento")

    if gerar:
        if TEMPLATE_PATH.exists():
            template_bytes = TEMPLATE_PATH.read_bytes()
        elif template_file is not None:
            template_bytes = template_file.getvalue()
        else:
            st.error("Nenhum modelo .docx disponível.")
            st.stop()

        valores = {
            "Cenário": cenario,
            "Caso de Teste": caso_teste,
            "IPC": ipc,
            "QA": qa,
            "DEV": dev,
            "Squad": squad,
        }

        resultado = fill_document(template_bytes, valores, imagens)

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
