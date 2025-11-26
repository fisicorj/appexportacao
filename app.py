import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import zipfile
import sqlite3
from datetime import datetime

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ==========================
# Configuração da página
# ==========================
st.set_page_config(page_title="Gerador de Relatórios TIC", layout="wide")
st.title("Gerador de Relatórios TIC → Tabelas em Word")

# ==========================
# CSS CUSTOMIZADO – Azul + Laranja
# ==========================
st.markdown(
    """
    <style>

    .stApp {
        background-color: #E6F0FA !important;
    }

    h1, h2, h3, h4 {
        color: #003F88 !important;
        font-weight: 700 !important;
    }

    p, label, span {
        color: #1C1C1C !important;
    }

    .stFileUploader label {
        color: #003F88 !important;
        font-weight: 600 !important;
    }

    .stButton > button {
        background-color: #003F88 !important;
        color: white !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.2rem !important;
        font-weight: 600 !important;
        border: 2px solid #003F88 !important;
    }

    .stButton > button:hover {
        background-color: #F28C28 !important;
        border-color: #F28C28 !important;
        color: white !important;
    }

    input[type="text"] {
        border: 2px solid #003F88 !important;
        border-radius: 6px !important;
    }

    .stMultiSelect > div {
        border: 2px solid #003F88 !important;
        border-radius: 6px !important;
    }

    </style>
    """,
    unsafe_allow_html=True
)

# ==========================
# Banco de dados para histórico
# ==========================
conn = sqlite3.connect("historico_relatorios.db", check_same_thread=False)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS historico (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    data_hora TEXT,
    nome_trilha TEXT,
    nome_arquivo_xlsx TEXT,
    abas_selecionadas TEXT,
    qtd_relatorios INTEGER
)
""")
conn.commit()


def registrar_historico(nome_trilha, nome_arquivo_xlsx, abas, qtd_relatorios):
    data_hora = datetime.now().isoformat(timespec="seconds")
    abas_str = ", ".join(abas)
    cursor.execute(
        """
        INSERT INTO historico (data_hora, nome_trilha, nome_arquivo_xlsx, abas_selecionadas, qtd_relatorios)
        VALUES (?, ?, ?, ?, ?)
        """,
        (data_hora, nome_trilha, nome_arquivo_xlsx, abas_str, qtd_relatorios)
    )
    conn.commit()


# ==========================
# Funções auxiliares de negócio
# ==========================
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


def gerar_doc_para_aba(xls, aba_nome: str) -> bytes | None:
    """Gera o DOCX (em bytes) para uma aba específica."""
    df = pd.read_excel(xls, sheet_name=aba_nome, header=None)

    header = df.iloc[1]
    try:
        col_termo = header[header.astype(str).str.contains("Termo", case=False)].index[0]
        col_campo = header[header.astype(str).str.contains("Campo", case=False)].index[0]
    except IndexError:
        return None

    dados = df.iloc[2:, [col_termo, col_campo]].copy()
    dados.columns = ["Termo", "Campo"]
    dados = dados.replace({np.nan: None})
    dados = dados[~(dados["Termo"].isnull() & dados["Campo"].isnull())]

    if dados.empty:
        return None

    # Documento em paisagem
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

    # Cabeçalho formatado
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

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ==========================
# Interface principal
# ==========================
uploaded_file = st.file_uploader(
    "Envie o arquivo de parametrização TIC (.xlsx)",
    type=["xlsx"]
)

if uploaded_file is not None:
    nome_trilha = st.text_input(
        "Nome da trilha (usado no nome do arquivo ZIP):",
        placeholder="Ex.: Trilha_IA, TIC_Módulo_3, Material_Senac"
    )

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
                nome_zip_base = nome_trilha.strip() if nome_trilha.strip() else "relatorios_tic"
                nome_zip = nome_zip_base + ".zip"

                zip_buffer = io.BytesIO()
                qtd_ok = 0

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for aba in aba_sel:
                        doc_bytes = gerar_doc_para_aba(xls, aba)
                        if doc_bytes is None:
                            continue
                        nome_arquivo = f"Tabela_{aba}.docx"
                        zipf.writestr(nome_arquivo, doc_bytes)
                        qtd_ok += 1

                zip_buffer.seek(0)

                # registra no histórico
                registrar_historico(
                    nome_trilha=nome_trilha.strip() or "(sem nome)",
                    nome_arquivo_xlsx=uploaded_file.name,
                    abas=aba_sel,
                    qtd_relatorios=qtd_ok
                )

                st.success(f"Relatórios gerados com sucesso! ({qtd_ok} arquivos)")

                st.download_button(
                    label="📥 Baixar ZIP",
                    data=zip_buffer,
                    file_name=nome_zip,
                    mime="application/zip"
                )

else:
    st.info("Envie o arquivo Excel para começar.")


# ==========================
# Seção de histórico
# ==========================
st.markdown("---")
st.subheader("Histórico de relatórios gerados")

if st.checkbox("Mostrar histórico"):
    df_hist = pd.read_sql_query("SELECT * FROM historico ORDER BY id DESC", conn)
    if df_hist.empty:
        st.info("Ainda não há registros no histórico.")
    else:
        st.dataframe(df_hist, use_container_width=True)
