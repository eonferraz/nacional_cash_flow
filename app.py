import streamlit as st
import pandas as pd
import pyodbc
from io import BytesIO
from datetime import datetime
import base64

st.set_page_config(page_title="Exportar Fluxo de Caixa", layout="wide")

@st.cache_resource
def carregar_dados():
    try:
        conn_str = (
            'DRIVER={ODBC Driver 17 for SQL Server};'
            'SERVER=benu.database.windows.net,1433;'
            'DATABASE=benu;'
            'UID=eduardo.ferraz;'
            'PWD=8h!0+a~jL8]B6~^5s5+v'
        )
        conn = pyodbc.connect(conn_str)
        query = """
            SELECT
                CAST(nro_unico AS INT) AS "Número Único",
                tipo_fluxo AS "Tipo de Fluxo",
                CAST(desdobramento AS INT) AS "Desdobramento",
                CAST(nro_nota AS INT) AS "Número da Nota",
                serie_nota AS "Série",
                CAST(nro_unico_nota AS INT) AS "Número Único Nota",
                FORMAT(data_faturamento, 'dd/MM/yyyy') AS "Data de Faturamento",
                FORMAT(data_negociacao, 'dd/MM/yyyy') AS "Data de Negociação",
                FORMAT(data_movimentacao, 'dd/MM/yyyy') AS "Data de Movimentação",
                FORMAT(data_vencimento, 'dd/MM/yyyy') AS "Data de Vencimento",
                FORMAT(data_baixa, 'dd/MM/yyyy') AS "Data de Baixa",
                nome_parceiro AS "Parceiro",
                cnpj_cpf AS "CNPJ/CPF",
                desc_top AS "Operação",
                desc_projeto AS "Projeto",
                historico AS "Histórico",
                FORMAT(valor_desdobramento, 'C', 'pt-BR') AS "Valor do Desdobramento",
                FORMAT(valor_baixa, 'C', 'pt-BR') AS "Valor da Baixa",
                CASE WHEN provisao = 'S' THEN 'Provisão' ELSE 'Efetivo' END AS "Tipo de Registro",
                status_titulo AS "Status do Título"
            FROM nacional_fluxo;
        """
        df = pd.read_sql(query, conn)
        conn.close()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        st.stop()

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

parceiros = original_df['Parceiro'].dropna().unique().tolist()
status_list = original_df['Status do Título'].dropna().unique().tolist()
tipo_fluxo_list = original_df['Tipo de Fluxo'].dropna().unique().tolist()
tipo_registro_list = original_df['Tipo de Registro'].dropna().unique().tolist()

filtro_parceiro = st.sidebar.multiselect("Parceiro", parceiros)
filtro_status = st.sidebar.multiselect("Status do Título", status_list)
filtro_fluxo = st.sidebar.multiselect("Tipo de Fluxo", tipo_fluxo_list)
filtro_tipo_registro = st.sidebar.radio("Tipo de Registro", options=["Todos"] + tipo_registro_list, index=0)

# Aplica filtros
df = original_df.copy()
df['Data de Faturamento'] = pd.to_datetime(df['Data de Faturamento'], dayfirst=True, errors='coerce')
df = df[df['Data de Faturamento'].notna()]
df = df[(df['Data de Faturamento'] >= pd.to_datetime(data_inicio)) & (df['Data de Faturamento'] <= pd.to_datetime(data_fim))]
if filtro_parceiro:
    df = df[df['Parceiro'].isin(filtro_parceiro)]
if filtro_status:
    df = df[df['Status do Título'].isin(filtro_status)]
if filtro_fluxo:
    df = df[df['Tipo de Fluxo'].isin(filtro_fluxo)]
if filtro_tipo_registro != "Todos":
    df = df[df['Tipo de Registro'] == filtro_tipo_registro]

# Totalizador
col1, col2 = st.columns(2)
df_valores = df.copy()
df_valores['Valor do Desdobramento'] = df_valores['Valor do Desdobramento'].replace('[R$\s]', '', regex=True).str.replace('.', '').str.replace(',', '.').astype(float)
df_valores['Valor da Baixa'] = df_valores['Valor da Baixa'].replace('[R$\s]', '', regex=True).str.replace('.', '').str.replace(',', '.').astype(float)

total_desdobramento = df_valores['Valor do Desdobramento'].sum()
total_baixa = df_valores['Valor da Baixa'].sum()

with col1:
    st.metric("Total Desdobrado", f"R$ {total_desdobramento:,.2f}".replace('.', '#').replace(',', '.').replace('#', ','))
with col2:
    st.metric("Total Baixado", f"R$ {total_baixa:,.2f}".replace('.', '#').replace(',', '.').replace('#', ','))

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
