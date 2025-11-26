import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import zipfile

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


st.set_page_config(page_title="Gerador de Relatórios TIC", layout="wide")
st.title("Gerador de Relatórios TIC → Tabelas em Word")


# ==== Funções auxiliares ==== #

def aba_e_numerica(nome: str) -> bool:
    return re.match(r"^\d", str(nome).strip()) is not None


def formatar_celula(celula, bold=False, branca=False):
    """Fonte Arial 12 na célula."""
    for par in celula.paragraphs:
        for run in par.runs:
            run.font.name = "Arial"
            run.font.size = Pt(12)
            run.bold = bold
            if branca:
                run.font.color.rgb = RGBColor(255, 255, 255)


def pintar_fundo(celula, rgb=(0, 0, 0)):
    """Aplica cor de fundo (w:shd) na célula."""
    cor = "%02X%02X%02X" % rgb  # (0,0,0) -> "000000"

    tc = celula._tc
    tcPr = tc.get_or_add_tcPr()

    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)

    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), cor)


def gerar_doc_para_aba(xls, aba_nome: str) -> bytes:
    """Gera um .docx (em bytes) para uma aba específica."""
    df = pd.read_excel(xls, sheet_name=aba_nome, header=None)

    # linha 1 tem os rótulos "Termo a ser preenchido"/"Campo de preenchimento"
    header = df.iloc[1]

    try:
        col_termo = header[header.astype(str).str.contains("Termo", case=False)].index[0]
        col_campo = header[header.astype(str).str.contains("Campo", case=False)].index[0]
    except IndexError:
        # não tem as colunas, não gera nada
        return None

    dados = df.iloc[2:, [col_termo, col_campo]].copy()
    dados.columns = ["Termo", "Campo"]
    dados = dados.replace({np.nan: None})
    dados = dados[~(dados["Termo"].isnull() & dados["Campo"].isnull())]

    if dados.empty:
        return None

    # Cria documento em paisagem
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    doc.add_paragraph(f"Relatório – {aba_nome}")

    # Tabela
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    hdr = table.rows[0].cells
    hdr[0].text = "Termo:"
    hdr[1].text = "Campo de preenchimento:"
    hdr[2].text = "Evidência"

    # Cabeçalho: preto, texto branco, Arial 12, centralizado
    for c in hdr:
        pintar_fundo(c, rgb=(0, 0, 0))
        formatar_celula(c, bold=True, branca=True)
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Linhas
    for _, row in dados.iterrows():
        termo = row["Termo"] if row["Termo"] else ""
        campo = row["Campo"] if row["Campo"] else ""

        nova = table.add_row().cells
        nova[0].text = str(termo)
        nova[1].text = str(campo)
        nova[2].text = ""

        for c in nova:
            formatar_celula(c)

    # Exporta para bytes
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ==== Interface Web ==== #

uploaded_file = st.file_uploader(
    "Envie o arquivo de parametrização TIC (.xlsx)",
    type=["xlsx"]
)

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    abas_numericas = [a for a in xls.sheet_names if aba_e_numerica(a)]

    if not abas_numericas:
        st.warning("Nenhuma aba numerada encontrada no arquivo.")
    else:
        st.markdown("### Abas numeradas detectadas")
        aba_sel = st.multiselect(
            "Selecione as abas para gerar relatório:",
            options=abas_numericas,
            default=abas_numericas
        )

        if st.button("Gerar relatórios (DOCX)"):
            if not aba_sel:
                st.error("Selecione pelo menos uma aba.")
            else:
                # Cria um ZIP em memória
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for aba in aba_sel:
                        doc_bytes = gerar_doc_para_aba(xls, aba)
                        if doc_bytes is None:
                            continue
                        nome_arquivo = f"Tabela_{aba}.docx"
                        zipf.writestr(nome_arquivo, doc_bytes)
                zip_buffer.seek(0)

                st.success("Relatórios gerados com sucesso!")

                st.download_button(
                    label="📥 Baixar ZIP com todos os DOCX",
                    data=zip_buffer,
                    file_name="relatorios_tic_tabelas.zip",
                    mime="application/zip"
                )
else:
    st.info("Envie o arquivo Excel para começar.")
