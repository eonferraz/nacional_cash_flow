import streamlit as st
import pandas as pd
import pyodbc
from io import BytesIO
from datetime import datetime
import base64

st.set_page_config(page_title="Exportar Fluxo de Caixa", layout="wide")

@st.cache_resource
def carregar_dados():
    conn_str = (
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=sx-global.database.windows.net;'
        'DATABASE=sx_comercial;'
        'UID=paulo.ferraz;'
        'PWD=Gs!^42j$G0f0^EI#ZjRv'
    )
    conn = pyodbc.connect(conn_str)
    query = """
        SELECT
            nro_unico AS numero_unico,
            tipo_fluxo,
            desdobramento,
            nro_nota,
            serie_nota,
            nro_unico_nota,
            data_faturamento,
            data_negociacao,
            data_movimentacao,
            data_vencimento,
            data_baixa,
            nome_parceiro,
            cnpj_cpf,
            desc_top,
            desc_projeto,
            historico,
            valor_desdobramento,
            valor_baixa,
            status_titulo
        FROM nacional_fluxo;
    """
    df = pd.read_sql(query, conn)
    conn.close()
    df['data_faturamento'] = pd.to_datetime(df['data_faturamento'], errors='coerce')
    return df

# Codifica imagem da logo
with open("nacional-escuro.svg", "rb") as image_file:
    encoded = base64.b64encode(image_file.read()).decode()
logo_img = f"data:image/svg+xml;base64,{encoded}"

# Header com logo e título alinhados verticalmente ao centro
st.markdown(f"""
    <div style='display: flex; align-items: center; gap: 20px;'>
        <img src='{logo_img}' width='80'>
        <h1 style='margin: 0;'>Fluxo de Caixa</h1>
    </div>
""", unsafe_allow_html=True)

# Carrega dados
original_df = carregar_dados()

# Filtros
st.sidebar.header("Filtros")
hoje = datetime.today()
data_inicio = st.sidebar.date_input("Data Inicial", value=datetime(hoje.year, 1, 1))
data_fim = st.sidebar.date_input("Data Final", value=hoje)

parceiros = original_df['nome_parceiro'].dropna().unique().tolist()
status_list = original_df['status_titulo'].dropna().unique().tolist()

filtro_parceiro = st.sidebar.multiselect("Parceiro", parceiros)
filtro_status = st.sidebar.multiselect("Status do Título", status_list)

# Aplica filtros
df = original_df.copy()
df = df[df['data_faturamento'].notna()]
df = df[(df['data_faturamento'] >= pd.to_datetime(data_inicio)) & (df['data_faturamento'] <= pd.to_datetime(data_fim))]
if filtro_parceiro:
    df = df[df['nome_parceiro'].isin(filtro_parceiro)]
if filtro_status:
    df = df[df['status_titulo'].isin(filtro_status)]

# Exibe dados
st.dataframe(df, use_container_width=True)

# Exporta para Excel com ajuste de largura
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Fluxo de Caixa', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Fluxo de Caixa']
    for i, col in enumerate(df.columns):
        largura = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, largura)

nome_arquivo = f"fluxo_filtrado_{datetime.today().strftime('%Y%m%d')}.xlsx"

st.download_button(
    label="📥 Baixar Excel",
    data=buffer.getvalue(),
    file_name=nome_arquivo,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
