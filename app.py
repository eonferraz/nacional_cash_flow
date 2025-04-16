import streamlit as st
import pandas as pd
import pyodbc
from io import BytesIO
from datetime import datetime
import base64

st.set_page_config(page_title="Exportar Faturamento", layout="wide")

@st.cache_resource
def carregar_dados():
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
          numero_nf AS "Número NF",
          data_negociacao AS "Data Negociação",
          data_faturamento AS "Data Faturamento",
          ano_mes AS "Ano-Mês",
          ano AS "Ano",
          mes AS "Mês",
          data_entrada AS "Data Entrada",
          cod_parceiro AS "Código Parceiro",
          cod_projeto AS "Código Projeto",
          abrev_projeto AS "Abrev. Projeto",
          projeto AS "Projeto",
          cnpj AS "CNPJ",
          parceiro AS "Parceiro",
          cod_top AS "Código TOP",
          [top] AS "TOP",
          movimento AS "Movimento",
          cliente AS "Cliente",
          fornecedor AS "Fornecedor",
          codigo AS "Código Produto",
          descricao AS "Descrição",
          ncm AS "NCM",
          grupo AS "Grupo",
          cfop AS "CFOP",
          operacao AS "Operação",
          qtd_negociada AS "Qtd. Negociada",
          qtd_entregue AS "Qtd. Entregue",
          status AS "Status",
          saldo AS "Saldo",
          valor_unitario AS "Valor Unitário",
          valor_total AS "Valor Total",
          valor_icms AS "Valor ICMS",
          valor_ipi AS "Valor IPI",
          receita AS "Receita"
        FROM
          nacional_faturamento;
    """
    df = pd.read_sql(query, conn)
    conn.close()
    df['Data Faturamento'] = pd.to_datetime(df['Data Faturamento'], errors='coerce')
    return df

# Codifica imagem da logo
with open("nacional-escuro.svg", "rb") as image_file:
    encoded = base64.b64encode(image_file.read()).decode()
logo_img = f"data:image/svg+xml;base64,{encoded}"

# Header com logo e título alinhados verticalmente ao centro
st.markdown(f"""
    <div style='display: flex; align-items: center; gap: 20px;'>
        <img src='{logo_img}' width='80'>
        <h1 style='margin: 0;'>Dados de Faturamento</h1>
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
operacoes = original_df['Operação'].dropna().unique().tolist()

filtro_parceiro = st.sidebar.multiselect("Parceiro", parceiros)
filtro_operacao = st.sidebar.multiselect("Operação", operacoes)

# Aplica filtros
df = original_df.copy()
df = df[df['Data Faturamento'].notna()]
df = df[(df['Data Faturamento'] >= pd.to_datetime(data_inicio)) & (df['Data Faturamento'] <= pd.to_datetime(data_fim))]
if filtro_parceiro:
    df = df[df['Parceiro'].isin(filtro_parceiro)]
if filtro_operacao:
    df = df[df['Operação'].isin(filtro_operacao)]

# Exibe dados
st.dataframe(df, use_container_width=True)

# Exporta para Excel com ajuste de largura
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Faturamento', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Faturamento']
    for i, col in enumerate(df.columns):
        largura = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, largura)

nome_arquivo = f"faturamento_filtrado_{datetime.today().strftime('%Y%m%d')}.xlsx"

st.download_button(
    label="📥 Baixar Excel",
    data=buffer.getvalue(),
    file_name=nome_arquivo,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
