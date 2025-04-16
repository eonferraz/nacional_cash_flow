import streamlit as st
import pandas as pd
import pyodbc
from io import BytesIO
from datetime import datetime, timedelta
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
                FORMAT(data_negociacao, 'dd/MM/yyyy') AS "Data de Negociação",
                FORMAT(data_movimentacao, 'dd/MM/yyyy') AS "Data de Movimentação",
                data_vencimento AS "Data de Vencimento",
                FORMAT(data_baixa, 'dd/MM/yyyy') AS "Data de Baixa",
                cod_projeto AS "Código Projeto",
                desc_projeto AS "Projeto",
                nome_parceiro AS "Parceiro",
                cnpj_cpf AS "CNPJ/CPF",
                desc_top AS "Operação",
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

# Conversão de data
original_df['Data de Vencimento'] = pd.to_datetime(original_df['Data de Vencimento'], errors='coerce')
original_df = original_df[original_df['Data de Vencimento'].notna()]

# Filtros rápidos e por data
col1, col2, col3, col4, col5, col6 = st.columns([1,1,1,1,1,3])
hoje = datetime.today()

with col1:
    data_inicio = st.date_input("Início", value=datetime(hoje.year, 1, 1))
with col2:
    data_fim = st.date_input("Fim", value=hoje)
with col3:
    if st.button("Hoje"):
        data_inicio = data_fim = hoje
with col4:
    if st.button("Essa Semana"):
        data_inicio = hoje - timedelta(days=hoje.weekday())
        data_fim = hoje
with col5:
    if st.button("Esse Mês"):
        data_inicio = datetime(hoje.year, hoje.month, 1)
        data_fim = hoje
with col6:
    if st.button("Esse Ano"):
        data_inicio = datetime(hoje.year, 1, 1)
        data_fim = hoje

# Filtros adicionais
st.sidebar.header("Filtros Avançados")
parceiros = sorted(original_df['Parceiro'].dropna().unique().tolist())
status_list = sorted(original_df['Status do Título'].dropna().unique().tolist())
tipo_fluxo_list = sorted(original_df['Tipo de Fluxo'].dropna().unique().tolist())
tipo_registro_list = sorted(original_df['Tipo de Registro'].dropna().unique().tolist())
cod_projetos = sorted(original_df['Código Projeto'].dropna().unique().tolist())
nome_projetos = sorted(original_df['Projeto'].dropna().unique().tolist())

filtro_parceiro = st.sidebar.multiselect("Parceiro", parceiros)
filtro_status = st.sidebar.selectbox("Status do Título", options=["Todos"] + status_list)
filtro_fluxo = st.sidebar.selectbox("Tipo de Fluxo", options=["Todos"] + tipo_fluxo_list)
filtro_tipo_registro = st.sidebar.selectbox("Tipo de Registro", options=["Todos"] + tipo_registro_list)
filtro_cod_projeto = st.sidebar.selectbox("Código Projeto", options=["Todos"] + [str(p) for p in cod_projetos])
filtro_nome_projeto = st.sidebar.selectbox("Projeto", options=["Todos"] + nome_projetos)

# Aplica filtros
df = original_df.copy()
df = df[(df['Data de Vencimento'] >= pd.to_datetime(data_inicio)) & (df['Data de Vencimento'] <= pd.to_datetime(data_fim))]
if filtro_parceiro:
    df = df[df['Parceiro'].isin(filtro_parceiro)]
if filtro_status != "Todos":
    df = df[df['Status do Título'] == filtro_status]
if filtro_fluxo != "Todos":
    df = df[df['Tipo de Fluxo'] == filtro_fluxo]
if filtro_tipo_registro != "Todos":
    df = df[df['Tipo de Registro'] == filtro_tipo_registro]
if filtro_cod_projeto != "Todos":
    df = df[df['Código Projeto'].astype(str) == filtro_cod_projeto]
if filtro_nome_projeto != "Todos":
    df = df[df['Projeto'] == filtro_nome_projeto]

# Formata datas para exibição
df['Data de Vencimento'] = df['Data de Vencimento'].dt.strftime('%d/%m/%Y')

# Totalizador horizontal
df_valores = df.copy()
df_valores['Valor do Desdobramento'] = df_valores['Valor do Desdobramento'].replace('[R$\s]', '', regex=True).str.replace('.', '').str.replace(',', '.').astype(float)
df_valores['Valor da Baixa'] = df_valores['Valor da Baixa'].replace('[R$\s]', '', regex=True).str.replace('.', '').str.replace(',', '.').astype(float)

total_desdobramento = df_valores['Valor do Desdobramento'].sum()
total_baixa = df_valores['Valor da Baixa'].sum()

col1, col2 = st.columns(2)
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
