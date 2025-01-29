import streamlit as st
import pandas as pd
import camelot
import PyPDF2
import tempfile
import os
import base64
import numpy as np
from fuzzywuzzy import process

# Bibliotecas para gerar DOCX
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# Biblioteca para gerar PDF
from fpdf import FPDF

import re


# ==========================================
#   CONFIGURAÇÃO INICIAL DO STREAMLIT
# ==========================================
st.set_page_config(page_title="Analista de Extratos Bancários", layout="centered")

logo_path = "MP.png"           # Caminho para a logomarca
glossary_path = "Tarifas.txt"  # Caminho para o glossário


# ==========================================
#   FUNÇÃO PARA SANITIZAR O NOME DO CLIENTE
# ==========================================
def sanitize_nome_cliente(nome: str) -> str:
    """
    Remove acentos/caracteres especiais e troca espaços por underscores,
    para evitar problemas em nomes de arquivos.
    """
    nome = nome.strip()
    nome = nome.replace(" ", "_")
    nome = re.sub(r"[^\w\-_]", "", nome)
    return nome


# ==========================================
#   FUNÇÃO PARA TENTAR EXTRAIR NOME DO CLIENTE
# ==========================================
def extrair_nome_cliente(pdf_path):
    """
    Tenta extrair o nome do cliente de dentro do PDF,
    procurando pela substring 'Nome:' na primeira página.
    Caso não encontre, retorna 'Sem_Nome'.
    """
    try:
        with open(pdf_path, "rb") as file:
            pdf_reader = PyPDF2.PdfReader(file)
            first_page_text = pdf_reader.pages[0].extract_text() or ""

            if "Nome:" in first_page_text:
                pos = first_page_text.find("Nome:") + len("Nome:")
                restante = first_page_text[pos:].strip()
                linha = restante.split("\n")[0].strip()
                return sanitize_nome_cliente(linha) if linha else "Sem_Nome"
    except:
        pass

    return "Sem_Nome"


# ==========================================
#   FUNÇÃO PARA FORMATAR VALORES EM R$
# ==========================================
def formatar_valor_brl(valor):
    """
    Converte um valor numérico (float/int) para o formato BRL,
    ex: 1.234,56. Se não for conversível, retorna string original.
    """
    try:
        val = float(valor)
        txt_temp = f"{val:,.2f}"  # Exemplo: "1,234.56"
        txt_temp = txt_temp.replace(",", "X").replace(".", ",").replace("X", ".")
        return txt_temp
    except:
        return str(valor)


# ==========================================
#   FUNÇÕES PARA GERAÇÃO DE PDF/DOCX
# ==========================================

def df_to_pdf_bytes(df, titulo="Relatório", formatar_linhas_especiais=False, excluir_docto=False):
    """
    Gera um PDF (bytes) em formato A4 paisagem, com opção de:
      - destacar "Valor Total (R$)" / "Em dobro (R$)"
      - excluir a coluna "Docto." e somar sua largura à coluna "Histórico".
    """

    # Se solicitado, excluir 'Docto.' e somar a largura na coluna 'Histórico'
    HIST_WIDTH_NORMAL = 210
    DOCTO_WIDTH = 20
    if excluir_docto and "Docto." in df.columns:
        df = df.drop(columns=["Docto."], errors='ignore')
        # Ajustar a largura da coluna "Histórico" = 210 + 20 => 230
        hist_width = HIST_WIDTH_NORMAL + DOCTO_WIDTH
    else:
        hist_width = HIST_WIDTH_NORMAL

    class PDFTabela(FPDF):
        def __init__(self, orientation='L', unit='mm', format='A4'):
            super().__init__(orientation, unit, format)
            self.set_left_margin(10)
            self.set_right_margin(10)
            self.set_top_margin(10)
            self.set_auto_page_break(auto=True, margin=10)

            self.title_str = titulo
            self.col_names = df.columns.tolist()

            # Definir larguras
            self.col_widths = []
            for col in self.col_names:
                if col == "Histórico":
                    self.col_widths.append(hist_width)  # se excluiu "Docto.", é 230, senão 210
                elif col in ["Débito (R$)", "Crédito (R$)"]:
                    self.col_widths.append(27)
                else:
                    # "Data", ou qualquer outra
                    self.col_widths.append(20)

            # Alinhamentos
            self.col_aligns = []
            for col in self.col_names:
                if col == "Data":
                    self.col_aligns.append("C")
                elif col == "Histórico":
                    self.col_aligns.append("L")
                else:
                    self.col_aligns.append("R")

            self.row_height = 7
            self.font_size_data = 9
            self.font_size_header = 10
            self.font_size_title = 14

        def header(self):
            self.set_font("Arial", "B", self.font_size_title)
            self.cell(0, 10, self.title_str, 0, 1, "C")
            self.ln(2)

            self.set_font("Arial", "B", self.font_size_header)
            for i, col_name in enumerate(self.col_names):
                w = self.col_widths[i]
                a = self.col_aligns[i]
                self.cell(w, self.row_height, str(col_name), border=1, ln=0, align=a)
            self.ln(self.row_height)

        def footer(self):
            self.set_y(-15)
            self.set_font("Arial", "I", 8)
            self.cell(0, 10, f"Página {self.page_no()}/{{nb}}", 0, 0, "C")

        def gerar_tabela(self):
            self.set_font("Arial", "", self.font_size_data)
            for _, row in df.iterrows():
                if self.get_y() + self.row_height > (self.h - 10):
                    self.add_page()

                # Se é "Valor Total (R$)" ou "Em dobro (R$)"
                is_especial = (
                    "Histórico" in row
                    and row["Histórico"] in ["Valor Total (R$)", "Em dobro (R$)"]
                )
                if is_especial and formatar_linhas_especiais:
                    self.set_font("Arial", "B", 14)
                    self.set_text_color(255, 0, 0)
                else:
                    self.set_font("Arial", "", self.font_size_data)
                    self.set_text_color(0, 0, 0)

                for i, col_name in enumerate(self.col_names):
                    txt = str(row[col_name])
                    w = self.col_widths[i]
                    a = self.col_aligns[i]
                    self.cell(w, self.row_height, txt, border=1, ln=0, align=a)
                self.ln(self.row_height)

                if is_especial and formatar_linhas_especiais:
                    self.set_font("Arial", "", self.font_size_data)
                    self.set_text_color(0, 0, 0)

    if df.empty:
        pdf_empty = FPDF(orientation='L', unit='mm', format='A4')
        pdf_empty.add_page()
        pdf_empty.set_font("Arial", "B", 12)
        pdf_empty.cell(0, 10, "DataFrame vazio - nenhum dado para exibir.", 0, 1, "C")
        return pdf_empty.output(dest="S").encode("latin-1")

    pdf = PDFTabela()
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.gerar_tabela()
    return pdf.output(dest="S").encode("latin-1")


def df_to_doc_bytes(df, titulo="Relatório", adicionar_totais=False, excluir_docto=False):
    """
    Gera um DOCX em paisagem, com opção de:
      - inserir "Valor Total (R$)" / "Em dobro (R$)"
      - excluir "Docto." e somar sua largura à "Histórico".
    """

    # Se excluir 'Docto.', remover e somar a largura
    HIST_NORMAL = 227
    DOCTO_WIDTH = 20
    if excluir_docto and "Docto." in df.columns:
        df = df.drop(columns=["Docto."], errors='ignore')
        hist_width_mm = HIST_NORMAL + DOCTO_WIDTH  # 247 mm
    else:
        hist_width_mm = HIST_NORMAL  # 227 mm

    document = Document()

    for section in document.sections:
        section.orientation = WD_ORIENT.LANDSCAPE
        new_width, new_height = section.page_height, section.page_width
        section.page_width = new_width
        section.page_height = new_height

    title = document.add_heading(titulo, level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if df.empty:
        p = document.add_paragraph("DataFrame vazio - nenhum dado para exibir.")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        buf_empty = BytesIO()
        document.save(buf_empty)
        return buf_empty.getvalue()

    table = document.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid'

    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        hdr_cells[i].text = str(col_name)
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True

    # Define larguras heurísticas
    col_widths_inches = []
    for col in df.columns:
        if col == "Histórico":
            # Converte mm -> inches
            col_widths_inches.append(hist_width_mm / 25.4)
        elif col in ["Débito (R$)", "Crédito (R$)"]:
            col_widths_inches.append(20 / 25.4)
        else:
            col_widths_inches.append(15 / 25.4)
    for i, width in enumerate(col_widths_inches):
        table.columns[i].width = Inches(width)

    # Preencher linhas
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, item in enumerate(row):
            text = str(item)
            if df.columns[i] == "Histórico":
                text = text.replace('\n', ' ').replace('\r', ' ').strip()

            paragraph = row_cells[i].paragraphs[0]
            run = paragraph.add_run(text)

            if df.columns[i] == "Histórico":
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            run.font.size = Pt(9)
            paragraph.line_spacing = Pt(9)

            # destaque totals
            if "Histórico" in row and row["Histórico"] in ["Valor Total (R$)", "Em dobro (R$)"]:
                run.font.bold = True
                run.font.size = Pt(14)
                run.font.color.rgb = RGBColor(255, 0, 0)

    # Se quisermos gerar as linhas de soma
    if adicionar_totais:
        if "Débito (R$)" in df.columns:
            total_col = "Débito (R$)"
        else:
            total_col = "Crédito (R$)"

        numeros = (
            df[total_col]
            .str.replace("R$", "", regex=False)
            .str.replace(" ", "", regex=False)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        total = pd.to_numeric(numeros, errors='coerce').sum()

        total_row = table.add_row().cells
        if len(total_row) > 1:
            total_row[1].text = "Valor Total (R$)"
        if len(total_row) > 3:
            total_row[3].text = f"{total:.2f}"
        for i, cell in enumerate(total_row):
            paragraph = cell.paragraphs[0]
            run = paragraph.runs[0]
            if i == 1 or i == 3:
                run.font.bold = True
                run.font.size = Pt(14)
                run.font.color.rgb = RGBColor(255, 0, 0)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if i != 1 else WD_ALIGN_PARAGRAPH.LEFT

        double_total = total * 2
        double_row = table.add_row().cells
        if len(double_row) > 1:
            double_row[1].text = "Em dobro (R$)"
        if len(double_row) > 3:
            double_row[3].text = f"{double_total:.2f}"
        for i, cell in enumerate(double_row):
            paragraph = cell.paragraphs[0]
            run = paragraph.runs[0]
            if i == 1 or i == 3:
                run.font.bold = True
                run.font.size = Pt(14)
                run.font.color.rgb = RGBColor(255, 0, 0)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if i != 1 else WD_ALIGN_PARAGRAPH.LEFT

    buf = BytesIO()
    document.save(buf)
    return buf.getvalue()


# ==========================================
#   FUNÇÕES AUXILIARES
# ==========================================
def get_image_base64(file_path):
    if not os.path.exists(file_path):
        st.warning(f"Logomarca não encontrada em: {file_path}")
        return ""
    with open(file_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()


def carregar_glossario(path):
    try:
        with open(path, "r", encoding="utf-8") as file:
            return file.read().splitlines()
    except IOError:
        st.error(f"Erro ao ler o arquivo de glossário em: {path}")
        return []


def match_glossary(text, glossary, threshold=85):
    if not glossary or not text:
        return False
    result = process.extractOne(text, glossary)
    return (result is not None) and (result[1] >= threshold)


def filtrar_por_glossario(df, glossary, threshold=85):
    if df.empty or not glossary:
        return pd.DataFrame()
    mask = df["Histórico"].apply(lambda x: match_glossary(x, glossary, threshold))
    return df[mask]


def obter_numero_de_paginas(pdf_path):
    with open(pdf_path, "rb") as file:
        pdf_reader = PyPDF2.PdfReader(file)
        return len(pdf_reader.pages)


def ignorar_tabela(df):
    condicao = (
        (df == "Fone Fácil Bradesco").any(axis=1)
        | (df == "Se Preferir, fale com a BIA pelo").any(axis=1)
        | (df == "Saldo Invest Fácil").any(axis=1)
    )
    return not (condicao == False).all()


def converter_data_para_dois_digitos(data):
    if pd.isna(data) or data == "":
        return data
    try:
        partes = data.split('/')
        if len(partes) == 3 and len(partes[2]) == 4:
            partes[2] = partes[2][2:]
            return '/'.join(partes)
        return data
    except:
        return data


def processar_pdf(pdf_path):
    try:
        num_paginas = obter_numero_de_paginas(pdf_path)
        formato = ["90, 220, 320, 420, 520"] * num_paginas

        tables = camelot.read_pdf(
            pdf_path,
            pages="all",
            row_tol=15,
            flavor="stream",
            columns=formato,
        )

        extrato = pd.DataFrame(columns=["Data", "Histórico", "Docto.", "Crédito (R$)", "Débito (R$)", "Saldo (R$)"])

        with st.spinner('Processando tabelas...'):
            progress_bar = st.progress(0)
            for i, table in enumerate(tables):
                df = table.df
                if ignorar_tabela(df):
                    progress_bar.progress((i + 1) / len(tables))
                    continue

                check_start = (df == "Data").any(axis=1)
                if any(check_start):
                    idx = check_start.idxmax()
                    df = df[idx + 1:]

                # Ajusta colunas
                expected_cols = ["Data", "Histórico", "Docto.", "Crédito (R$)", "Débito (R$)", "Saldo (R$)"]
                if len(df.columns) >= 6:
                    df.columns = expected_cols
                else:
                    df = df.iloc[:, :6]
                    df.columns = expected_cols

                extrato = pd.concat([extrato, df], ignore_index=True)
                progress_bar.progress((i + 1) / len(tables))

        extrato["Data"] = extrato["Data"].replace("", np.nan).ffill()
        extrato["Data"] = extrato["Data"].apply(converter_data_para_dois_digitos)

        linhas_vazias = extrato[
            (extrato["Docto."] == "")
            & (extrato["Crédito (R$)"] == "")
            & (extrato["Débito (R$)"] == "")
            & (extrato["Saldo (R$)"] == "")
        ].index
        for idx in linhas_vazias:
            if idx + 1 < len(extrato):
                extrato.at[idx + 1, "Histórico"] = (
                    extrato.at[idx, "Histórico"] + " " + extrato.at[idx + 1, "Histórico"]
                )

        extrato = extrato.drop(linhas_vazias).reset_index(drop=True)
        return extrato

    except Exception as e:
        st.error(f"Erro ao processar o PDF: {str(e)}")
        return None


def filtrar_debitos(df):
    debitos = df[df["Débito (R$)"].notna() & (df["Débito (R$)"] != "")]
    cols_drop = ["Crédito (R$)", "Saldo (R$)"]
    debitos = debitos.drop(columns=cols_drop, errors='ignore')
    return debitos


# ==========================================
#     MAIN STREAMLIT (SOMENTE DÉBITO)
# ==========================================
def main():
    # Inicializar keys
    for key in [
        "df_extrato", "df_debito", "df_debito_gloss", "df_debito_gloss_filtrado",
        "nome_cliente"
    ]:
        if key not in st.session_state:
            st.session_state[key] = None

    # Exibir logomarca
    image_base64 = get_image_base64(logo_path)
    if image_base64:
        st.markdown(
            f"""
            <div style="display: flex; justify-content: center; align-items: center; margin-bottom: 20px;">
                <img src="data:image/png;base64,{image_base64}" alt="Logomarca" style="width: 300px;">
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.subheader("Análise de Extratos Bancários do Bradesco")

    # Carrega glossário
    glossary_terms = carregar_glossario(glossary_path)
    if not glossary_terms:
        st.warning("Glossário não encontrado ou vazio!")

    # Upload do PDF
    uploaded_file = st.file_uploader("Inserir Extrato Bancário do Bradesco", type="pdf")
    if uploaded_file is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            tmp_file.write(uploaded_file.read())
            pdf_path = tmp_file.name

        nome_cliente_encontrado = extrair_nome_cliente(pdf_path)
        st.session_state["nome_cliente"] = nome_cliente_encontrado or "Sem_Nome"

        try:
            df_extrato = processar_pdf(pdf_path)
            st.session_state["df_extrato"] = df_extrato

            if df_extrato is not None and not df_extrato.empty:
                st.success("PDF processado com sucesso!")
                st.markdown("### Extrato Completo")
                st.dataframe(df_extrato, use_container_width=True)
            else:
                st.warning("Não foi possível processar o PDF ou o extrato está vazio.")
        finally:
            os.unlink(pdf_path)

    # Se há extrato processado, vamos direto para "Análise de Débitos"
    nome_cliente = st.session_state.get("nome_cliente", "Sem_Nome")
    if st.session_state.get("df_extrato") is not None and not st.session_state["df_extrato"].empty:
        # --- PASSO 1) Filtrar Operações de Débito ---
        st.markdown("## Análise de Débitos")
        with st.form("filtrar_debitos_form"):
            st.markdown("### 1) Filtrar Operações de Débito")
            filtrar_debitos_submit = st.form_submit_button("Filtrar Débitos")

        if filtrar_debitos_submit:
            df_debito = filtrar_debitos(st.session_state["df_extrato"])
            st.session_state["df_debito"] = df_debito
            st.markdown("#### Resultado (Extrato de Débito)")
            st.dataframe(df_debito, use_container_width=True)

            pdf_debitos = df_to_pdf_bytes(df_debito, titulo="Extrato de Débitos")
            st.download_button(
                label="Baixar PDF (Débitos)",
                data=pdf_debitos,
                file_name=f"debitos_{nome_cliente}.pdf",
                mime="application/pdf",
            )

        # --- PASSO 2) Filtrar Débitos no Glossário ---
        if st.session_state.get("df_debito") is not None and not st.session_state["df_debito"].empty:
            with st.form("filtrar_glossario_debito_form"):
                st.markdown("### 2) Filtrar Débitos no Glossário (com Precisão Ajustável)")
                precision_debito = st.slider(
                    "Precisão da correspondência para Débitos (0.5 a 1.0):",
                    min_value=0.5,
                    max_value=1.0,
                    value=0.85,
                    step=0.025
                )
                filtrar_gloss_debito_submit = st.form_submit_button("Filtrar Débitos no Glossário")

            if filtrar_gloss_debito_submit:
                df_debito_gloss = filtrar_por_glossario(
                    st.session_state["df_debito"], glossary_terms, threshold=int(precision_debito * 100)
                )
                df_debito_gloss = df_debito_gloss.drop(columns=["Crédito (R$)", "Saldo (R$)"], errors='ignore')
                st.session_state["df_debito_gloss"] = df_debito_gloss
                st.session_state["df_debito_gloss_filtrado"] = None

                st.markdown("#### Resultado: Débitos + Glossário")
                st.dataframe(df_debito_gloss, use_container_width=True)

                pdf_gloss_debito = df_to_pdf_bytes(df_debito_gloss, titulo="Débitos (Filtrados no Glossário)")
                st.download_button(
                    label="Baixar PDF (Débitos Glossário)",
                    data=pdf_gloss_debito,
                    file_name=f"debitos_glossario_{nome_cliente}.pdf",
                    mime="application/pdf",
                )

        # --- PASSO 3) Lista Única de 'Histórico' ---
        if st.session_state.get("df_debito_gloss") is not None and not st.session_state["df_debito_gloss"].empty:
            with st.form("excluir_debitos_form"):
                st.markdown("### 3) Lista Única de 'Histórico' para Débitos + Inclusão")

                df_gloss_original_debito = st.session_state["df_debito_gloss"]
                df_base_exclusao_debito = (
                    st.session_state["df_debito_gloss_filtrado"]
                    if st.session_state["df_debito_gloss_filtrado"] is not None
                    else df_gloss_original_debito
                ).copy()

                valores_unicos_debito = sorted(df_base_exclusao_debito["Histórico"].unique())
                st.markdown("#### Lista Única de 'Histórico' (Débitos - sem repetições)")
                st.write("Marque os itens que deseja incluir:")

                selected_historicos_debito = []
                for i, hist in enumerate(valores_unicos_debito):
                    count_hist = df_base_exclusao_debito[df_base_exclusao_debito["Histórico"] == hist].shape[0]
                    rotulo = f"{i+1}- {hist} ({count_hist} {'vez' if count_hist == 1 else 'vezes'})"
                    if st.checkbox(rotulo, key=f"unique_hist_debito_{i}"):
                        selected_historicos_debito.append(hist)

                confirmar_inclusao_debito = st.form_submit_button("Confirmar Inclusão (Débitos)")
                if confirmar_inclusao_debito:
                    if selected_historicos_debito:
                        df_filtrado_debito = df_base_exclusao_debito[
                            df_base_exclusao_debito["Histórico"].isin(selected_historicos_debito)
                        ].reset_index(drop=True)

                        st.success("Operações de Débito incluídas com sucesso!")
                        st.session_state["df_debito_gloss_filtrado"] = df_filtrado_debito

                        st.markdown("#### Lista Restante após Inclusões (Débitos - sem repetições)")
                        if df_filtrado_debito.empty:
                            st.write("Nenhum histórico de Débito restante.")
                        else:
                            df_restante_unicos_debito = df_filtrado_debito["Histórico"].value_counts().reset_index()
                            df_restante_unicos_debito.columns = ["Histórico", "Ocorrências"]
                            st.dataframe(df_restante_unicos_debito, use_container_width=True)
                    else:
                        st.warning("Nenhuma descrição de Débito foi selecionada.")

        # --- PASSO 4) Apresentar Tarifas p/ Débitos (excluir Docto., somar largura) ---
        if st.session_state.get("df_debito_gloss_filtrado") is not None and not st.session_state["df_debito_gloss_filtrado"].empty:
            with st.form("apresentar_tarifas_debito_form"):
                st.markdown("### 4) Apresentar Tarifas para Débitos (DataFrame Final Ordenado)")
                apresentar_tarifas_debito_submit = st.form_submit_button("Apresentar Tarifas para Débitos")

            if apresentar_tarifas_debito_submit:
                df_para_exibir_debito = st.session_state["df_debito_gloss_filtrado"]
                if not df_para_exibir_debito.empty:
                    # Remove colunas não necessárias
                    df_para_exibir_debito = df_para_exibir_debito.drop(columns=["Crédito (R$)", "Saldo (R$)"], errors='ignore')

                    # Converter valores de Débito para float
                    numeros_debitos = (
                        df_para_exibir_debito["Débito (R$)"]
                        .str.replace("R$", "", regex=False)
                        .str.replace(" ", "", regex=False)
                        .str.replace(".", "", regex=False)
                        .str.replace(",", ".", regex=False)
                    )
                    valores_float_debito = pd.to_numeric(numeros_debitos, errors='coerce').fillna(0.0)

                    # Forçar soma positiva
                    valores_float_debito = valores_float_debito.abs()
                    total_debitos = valores_float_debito.sum()

                    # Formatar em BRL
                    df_para_exibir_debito["Débito (R$)"] = valores_float_debito.apply(formatar_valor_brl)

                    # Duas linhas extras: Valor Total, Em Dobro
                    valor_total = pd.DataFrame({
                        "Data": [""],
                        "Histórico": ["Valor Total (R$)"],
                        "Docto.": [""],
                        "Débito (R$)": [formatar_valor_brl(total_debitos)]
                    })
                    em_dobro = pd.DataFrame({
                        "Data": [""],
                        "Histórico": ["Em dobro (R$)"],
                        "Docto.": [""],
                        "Débito (R$)": [formatar_valor_brl(total_debitos * 2)]
                    })

                    extrato_debito_final = pd.concat([df_para_exibir_debito, valor_total, em_dobro], ignore_index=True)

                    # Ordenar por data
                    extrato_debito_final["Data"] = pd.to_datetime(
                        extrato_debito_final["Data"], format="%d/%m/%y", errors='coerce'
                    )
                    extrato_debito_final = extrato_debito_final.sort_values(by="Data")
                    extrato_debito_final["Data"] = extrato_debito_final["Data"].dt.strftime("%d/%m/%y")
                    extrato_debito_final["Data"] = extrato_debito_final["Data"].fillna("")

                    # Esvazia Data nas linhas de total
                    extrato_debito_final.loc[
                        extrato_debito_final["Histórico"].isin(["Valor Total (R$)", "Em dobro (R$)"]),
                        "Data"
                    ] = ""

                    st.markdown("#### DataFrame Final de Débitos (Cronológico)")
                    st.dataframe(extrato_debito_final, use_container_width=True)

                    # Gerar PDF: excluir "Docto." e somar largura
                    pdf_final_debito = df_to_pdf_bytes(
                        extrato_debito_final,
                        titulo="Extrato Final de Débitos (Cronológico)",
                        formatar_linhas_especiais=True,
                        excluir_docto=True  # <-- Aqui exclui "Docto." e amplia "Histórico"
                    )
                    st.download_button(
                        label="Baixar PDF (Débitos Final - Cronológico)",
                        data=pdf_final_debito,
                        file_name=f"debitos_final_cronologico_{nome_cliente}.pdf",
                        mime="application/pdf",
                    )

                    # Gerar DOCX: idem
                    doc_final_debito = df_to_doc_bytes(
                        extrato_debito_final,
                        titulo="Extrato Final de Débitos (Cronológico)",
                        adicionar_totais=False,
                        excluir_docto=True  # <-- Excluir "Docto." e soma largura
                    )
                    st.download_button(
                        label="Baixar DOCX (Débitos Final - Cronológico)",
                        data=doc_final_debito,
                        file_name=f"debitos_final_cronologico_{nome_cliente}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                else:
                    st.warning("Não há extrato final para Débitos.")


if __name__ == "__main__":
    main()

