import streamlit as st
import pandas as pd
import io
import zipfile
from pathlib import Path

st.set_page_config(page_title="Gerador de Planilhas", page_icon="🧮", layout="centered")

st.title("🧮 Planilhas, Limpador de Colunas")
st.write("Realize o upload de sua planilha, marque a caixa de seleção para usar o preset que limpará automaticamente determinadas colunas ou selecione manualmente as colunas que deseja deletar.")

# -----------------------
# Presets fixos
# -----------------------
PRESET_ORIGINAL = [
    "CPF", "DataNascimento", "CEP", "UF", "Cidade", "Cota", "Raça",
    "LocalProva", "Lingua", "Deficiente", "Tipo de deficiência", "Adaptação solicitada",
    "Isencao", "Isento", "Pagou", "Data Pagamento", "Sequencial", "Turma",
    "Data Inscricao", "Hora Inscricao", "Data de Atualização"
]

PRESET_SHEET1 = [
    "CPF", "DataNascimento", "CEP", "UF", "Cidade", "Cota", "Raça",
    "LocalProva", "Lingua", "Deficiente", "Tipo de deficiência", "Adaptação solicitada",
    "Isencao", "Isento", "Pagou", "Data Pagamento", "Sequencial", "Turma",
    "Data Inscricao", "Hora Inscricao", "Data de Atualização", "Campus2", "Curso2"
]

PRESET_SHEET2 = [
    "CPF", "DataNascimento", "CEP", "UF", "Cidade", "Cota", "Raça",
    "LocalProva", "Lingua", "Deficiente", "Tipo de deficiência", "Adaptação solicitada",
    "Isencao", "Isento", "Pagou", "Data Pagamento", "Sequencial", "Turma",
    "Data Inscricao", "Hora Inscricao", "Data de Atualização", "Campus1", "Curso1"
]

# -----------------------
# Funções auxiliares
# -----------------------


def normalize_name(s: str) -> str:
    if s is None:
        return ""
    return "".join(str(s).lower().split()).replace("_", "")

def find_matches(available_cols, desired_list):
    norm_map = {normalize_name(c): c for c in available_cols}
    found, not_found = [], []
    for d in desired_list:
        nd = normalize_name(d)
        if nd in norm_map:
            found.append(norm_map[nd])
            continue
        matched = None
        for nc, real in norm_map.items():
            if nd in nc or nc in nd:
                matched = real
                break
        if matched:
            found.append(matched)
        else:
            not_found.append(d)
    return found, not_found

# -----------------------
# Upload
# -----------------------
uploaded_file = st.file_uploader("📤 Envie sua planilha (.xlsx ou .xls)", type=["xlsx", "xls"])

if not uploaded_file:
    st.info("Faça upload de um arquivo Excel para começar.")
    st.stop()

# Lê o arquivo Excel
try:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
except Exception:
    df = pd.read_excel(uploaded_file)

st.success(f"✅ Arquivo carregado: {len(df)} linhas × {len(df.columns)} colunas")
st.dataframe(df.head())

all_columns = list(df.columns)

# -----------------------
# Configuração manual ou automática
# -----------------------
st.markdown("### ⚙️ Configuração de colunas a remover")

use_presets = st.checkbox("Usar presets automáticos (recomendado)", value=True)

preset_original_matches, preset_original_notfound = find_matches(all_columns, PRESET_ORIGINAL)
preset_sheet1_matches, preset_sheet1_notfound = find_matches(all_columns, PRESET_SHEET1)
preset_sheet2_matches, preset_sheet2_notfound = find_matches(all_columns, PRESET_SHEET2)

if preset_original_notfound or preset_sheet1_notfound or preset_sheet2_notfound:
    with st.expander("⚠️ Colunas dos presets não encontradas (nomes diferentes na planilha?)"):
        if preset_original_notfound:
            st.write("Original:", preset_original_notfound)
        if preset_sheet1_notfound:
            st.write("Planilha 1:", preset_sheet1_notfound)
        if preset_sheet2_notfound:
            st.write("Planilha 2:", preset_sheet2_notfound)

# Campos para edição manual (preenchidos com presets se selecionado)

cols_original = st.multiselect("Colunas a remover da planilha original:", all_columns, default=preset_original_matches if use_presets else [])

with st.container():
    st.markdown('<div class="sheet1">', unsafe_allow_html=True)
    cols_sheet1 = st.multiselect("Colunas a remover da planilha 1:", all_columns, default=preset_sheet1_matches if use_presets else [])
    st.markdown('</div>', unsafe_allow_html=True)

with st.container():
    st.markdown('<div class="sheet2">', unsafe_allow_html=True)
    cols_sheet2 = st.multiselect("Colunas a remover da planilha 2:", all_columns, default=preset_sheet2_matches if use_presets else [])
    st.markdown('</div>', unsafe_allow_html=True)



# -----------------------
# Botão de geração
# -----------------------
if st.button("🚀 Gerar planilhas"):
    with st.spinner("Gerando planilhas..."):
        # Cria cópias removendo colunas
        df_original_mod = df.drop(columns=[c for c in cols_original if c in df.columns], errors="ignore")
        df_sheet1 = df.drop(columns=[c for c in cols_sheet1 if c in df.columns], errors="ignore")
        df_sheet2 = df.drop(columns=[c for c in cols_sheet2 if c in df.columns], errors="ignore")

        # Cria buffer ZIP na memória
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
            # Salva as 3 planilhas dentro do zip
            for name, dataf in {
                "planilha_original.xlsx": df_original_mod,
                "planilha_1.xlsx": df_sheet1,
                "planilha_2.xlsx": df_sheet2
            }.items():
                file_buf = io.BytesIO()
                dataf.to_excel(file_buf, index=False, engine="openpyxl")
                z.writestr(name, file_buf.getvalue())

        buffer.seek(0)
        st.success("✅ Planilhas geradas com sucesso!")
        st.download_button(
            label="⬇️ Baixar arquivo ZIP com as 3 planilhas",
            data=buffer,
            file_name="planilhas_geradas.zip",
            mime="application/zip"
        )
