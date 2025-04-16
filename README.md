# Dashboard de Faturamento - Nacional Indústria Mecânica

Este projeto é um dashboard interativo desenvolvido em Streamlit para visualização dos dados de faturamento da empresa Nacional Indústria Mecânica.

## Funcionalidades
- Gráfico de faturamento mensal
- Faturamento por operação
- Faturamento por cliente (Top 10)
- Tabela com heatmap de receitas

## Tecnologias utilizadas
- Python 3.10+
- Streamlit
- Pandas
- Plotly
- PyODBC
- SQL Server no Azure

## Rodando localmente

### 1. Clone o repositório
```bash
git clone https://github.com/seu-usuario/seu-repo.git
cd seu-repo
```

### 2. Crie um ambiente virtual e ative
```bash
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
.venv\Scripts\activate    # Windows
```

### 3. Instale as dependências
```bash
pip install -r requirements.txt
```

### 4. Execute o aplicativo
```bash
streamlit run app.py
```

## Estrutura de Arquivos
```
.
├── app.py                # Código principal do Streamlit
├── db_utils.py           # Conexão com o banco de dados
├── requirements.txt      # Dependências
├── .streamlit
│   └── config.toml       # Tema e configurações do app
├── nacional-branco.svg   # Logo branca da empresa
├── nacional-escuro.svg   # Logo azul escuro da empresa
```

## Observações
- As credenciais de banco estão em `db_utils.py`. Para maior segurança, recomenda-se usar variáveis de ambiente em produção.
- O acesso ao banco requer o driver ODBC para SQL Server (Windows/Linux).

---
Desenvolvido por Paulo Eduardo de Oliveira Ferraz para Nacional Indústria Mecânica. 🚀
