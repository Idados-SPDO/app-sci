import streamlit as st
import pandas as pd
from io import BytesIO
import os
from openpyxl import Workbook
from streamlit_sortables import sort_items

st.set_page_config(page_title="SPDO - App SCI", layout="wide", page_icon="logo_fgv.png")
st.title("SCI – Ajuste de Entrada de Dados")


st.logo("logo_ibre.png")
# ─── Sidebar ─────────────────────────────────────────────────────
with st.sidebar:
    st.header("1. Envio e pré-configuração")
    uploaded_file = st.file_uploader("Envie um Excel:", type=["xlsx", "xls"])
    if uploaded_file:
        nome_sem_ext = os.path.splitext(uploaded_file.name)[0]
        df = pd.read_excel(uploaded_file)

        st.markdown("---")
        st.subheader("2. Organize as colunas")
        containers = [
            {"header": "Usadas",    "items": list(df.columns)},
            {"header": "Ignoradas", "items": []}
        ]
        conts = sort_items(containers, multi_containers=True)

        st.markdown("**3. Renomeie**")
        df_map = pd.DataFrame({
            "antigo": conts[0]["items"],
            "novo":   conts[0]["items"]
        })
        df_map_edit = st.data_editor(df_map, 
                                     num_rows="fixed", 
                                     disabled=["antigo"], 
                                     hide_index=True)
        rename_map = dict(zip(df_map_edit["antigo"], df_map_edit["novo"]))
        st.session_state.df_filtrado = df[conts[0]["items"]].rename(columns=rename_map)
        st.session_state.nome_arquivo   = nome_sem_ext

# ─── Corpo principal ──────────────────────────────────────────────
if "df_filtrado" in st.session_state:
    st.subheader("Preview da planilha ajustada")
    st.data_editor(st.session_state.df_filtrado, use_container_width=True, disabled=True, hide_index=True)

    st.markdown("### ⬇️ Baixar resultado")
    buffer = BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    wb.create_sheet("Planilha1")
    wb.save(buffer)
    buffer.seek(0)
    with pd.ExcelWriter(buffer, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        st.session_state.df_filtrado.to_excel(w, index=False, sheet_name="Planilha1")
    buffer.seek(0)

    st.download_button(
        "Download Excel",
        data=buffer.getvalue(),
        file_name=f"{st.session_state.nome_arquivo} - Reorganizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
