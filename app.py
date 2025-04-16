import streamlit as st
import pandas as pd
import pyodbc
import plotly.express as px

# Conexão fixa com Azure SQL
conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=sx-global.database.windows.net;'
    'DATABASE=sx_comercial;'
    'UID=paulo.ferraz;'
    'PWD=Gs!^42j$G0f0^EI#ZjRv'
)

@st.cache_data(ttl=3600)
def carregar_dados():
    query = """
        SELECT
            nro_unico,
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
        FROM nacional_fluxo
    """
    return pd.read_sql(query, conn)

st.set_page_config(page_title="Fluxo Financeiro - Nacional", layout="wide")
st.title("Fluxo de Caixa - Nacional")

# Dados
with st.spinner("Carregando dados..."):
    df = carregar_dados()

# Filtros
col1, col2 = st.columns(2)

with col1:
    parceiros = st.multiselect("Filtrar por parceiro:", options=sorted(df["nome_parceiro"].dropna().unique()), default=None)

with col2:
    status = st.multiselect("Filtrar por status:", options=sorted(df["status_titulo"].dropna().unique()), default=None)

if parceiros:
    df = df[df["nome_parceiro"].isin(parceiros)]
if status:
    df = df[df["status_titulo"].isin(status)]

# Visão Geral
st.subheader("Resumo do Fluxo")
resumo = df.groupby("tipo_fluxo")["valor_desdobramento"].sum().reset_index()
st.dataframe(resumo, use_container_width=True)

# Evolução temporal
st.subheader("Evolução Temporal")
df_evol = df.copy()
df_evol["data"] = pd.to_datetime(df_evol["data_negociacao"])

fig = px.line(df_evol.groupby(["data", "tipo_fluxo"])["valor_desdobramento"].sum().reset_index(),
              x="data", y="valor_desdobramento", color="tipo_fluxo",
              title="Fluxo por Tipo ao Longo do Tempo")
st.plotly_chart(fig, use_container_width=True)

# Tabela detalhada
st.subheader("Tabela Detalhada")
st.dataframe(df.sort_values("data_negociacao", ascending=False), use_container_width=True)
