# db_utils.py
import pandas as pd
import pyodbc
import streamlit as st
from functools import lru_cache

# Usa cache para performance
@st.cache_data(show_spinner="Conectando ao banco de dados...")
def get_data():
    server = 'benu.database.windows.net'
    database = 'benu'
    username = 'eduardo.ferraz'
    password = '8h!0+a~jL8]B6~^5s5+v'

    conn_str = (
        f'DRIVER={{ODBC Driver 17 for SQL Server}};'
        f'SERVER={server};DATABASE={database};UID={username};PWD={password}'
    )

    query = "SELECT * FROM nacional_faturamento WHERE receita IS NOT NULL"

    with pyodbc.connect(conn_str) as conn:
        df = pd.read_sql(query, conn)

    return df
