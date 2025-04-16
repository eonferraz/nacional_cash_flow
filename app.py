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
            'SERVER=sx-global.database.windows.net;'
            'DATABASE=sx_comercial;'
            'UID=paulo.ferraz;'
            'PWD=Gs!^42j$G0f0^EI#ZjRv'
        )
        conn = pyodbc.connect(conn_str)
        query = """
            SELECT
                CAST(nro_unico AS INT) AS numero_unico,
                tipo_fluxo,
                CAST(desdobramento AS INT) AS desdobramento,
                CAST(nro_nota AS INT) AS nro_nota,
                serie_nota,
                CAST(nro_unico_nota AS INT) AS nro_unico_nota,
                CAST(data_faturamento AS DATE) AS data_faturamento,
                CAST(data_negociacao AS DATE) AS data_negociacao,
                CAST(data_movimentacao AS DATE) AS data_movimentacao,
                CAST(data_vencimento AS DATE) AS data_vencimento,
                CAST(data_baixa AS DATE) AS data_baixa,
                nome_parceiro,
                cnpj_cpf,
                desc_top,
                desc_projeto,
                historico,
                CAST(valor_desdobramento AS FLOAT) AS valor_desdobramento,
                CAST(valor_baixa AS FLOAT) AS valor_baixa,
                status_titulo
            FROM nacional_fluxo;
        """
        df = pd.read_sql(query, conn)
        conn.close()
        df['data_faturamento'] = pd.to_datetime(df['data_faturamento'], errors='coerce')
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

# Exibe dados completos
st.dataframe(original_df, use_container_width=True)

# Exporta para Excel com ajuste de largura
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    original_df.to_excel(writer, sheet_name='Fluxo de Caixa', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Fluxo de Caixa']
    for i, col in enumerate(original_df.columns):
        largura = max(original_df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, largura)

nome_arquivo = f"fluxo_completo_{datetime.today().strftime('%Y%m%d')}.xlsx"

st.download_button(
    label="📥 Baixar Excel",
    data=buffer.getvalue(),
    file_name=nome_arquivo,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
