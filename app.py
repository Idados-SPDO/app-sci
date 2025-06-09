import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from streamlit_sortables import sort_items

st.set_page_config(page_title="Validação + Reordenação com Sortables", layout="wide")
st.title("SCI - App Ajuste de Entrada de Dados")

# 2) Uploader para o usuário enviar o Excel
uploaded_file = st.file_uploader(
    "Envie um arquivo Excel (.xlsx ou .xls):",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)
    except Exception as e:
        st.error(f"Não foi possível ler o arquivo como Excel: {e}")
        st.stop()

    colunas_originais = list(df.columns)

    columns_padrão = "Solicitação | Código CI ext. | JOB | Código Elementar | Código Ins. Inf. | Tipo Insumo | Código Externo | Serviço | Descrição | Elementar | Descrição | Ins. Inf. | Código | Informante | Descrição | Informante | Última Mensagem "
    st.write(f'**Padrão das colunas a se seguir para reorganizar:** {columns_padrão}')
    # 1) Definindo os dois containers:
    containers = [
        {
            'header': 'Colunas Utilizadas:',
            'items': colunas_originais  # começam todas aqui
        },
        {
            'header': 'Colunas Ignoradas:',
            'items': []  # inicialmente vazio
        }
    ]

    # 2) Renderizando os dois painéis para drag & drop:
    containers_ordenados = sort_items(
        containers,
        multi_containers=True,
    )

    # 3) Extraindo as listas aprovadas e rejeitadas:
    colunas_usadas     = containers_ordenados[0]['items']
    colunas_ignoradas = containers_ordenados[1]['items']

    # 4) Reordenando e filtrando o DataFrame:
    df_filtrado = df[colunas_usadas]

    st.markdown("### Planilha Reorganizada")
    st.dataframe(df_filtrado)

    wb = Workbook()
    primeira_aba = wb.active
    wb.remove(primeira_aba)
    wb.create_sheet("Planilha1")

    temp_stream = BytesIO()
    wb.save(temp_stream)
    temp_stream.seek(0)
    with pd.ExcelWriter(temp_stream, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name="Planilha1")
    processed_data = temp_stream.getvalue()

    st.download_button(
            label="⬇️ Excel Reorganizado",
            data=processed_data,
            file_name="arquivo_reordenado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

        
