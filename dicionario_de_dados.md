# dicionario_de_dados.md

## Dicionário de Dados - Tabela `nacional_faturamento`

| Nome do Campo       | Tipo de Dado       | Descrição                                                                 |
|---------------------|--------------------|---------------------------------------------------------------------------|
| numero_nf           | int                | Número da nota fiscal                                                    |
| data_negociacao     | date               | Data em que a negociação foi realizada                                 |
| data_faturamento    | date               | Data de faturamento                                                       |
| ano_mes             | varchar(7)         | Ano e mês no formato AAAA-MM                                             |
| ano                 | int                | Ano da negociação ou faturamento                                       |
| mes                 | int                | Mês da negociação ou faturamento (1-12)                               |
| data_entrada        | date               | Data de entrada do produto                                                |
| cod_parceiro        | int                | Código do parceiro                                                        |
| cod_projeto         | varchar(20)        | Código interno do projeto                                                |
| abrev_projeto       | varchar(100)       | Abreviação do nome do projeto                                           |
| projeto             | varchar(200)       | Nome completo do projeto                                                  |
| cnpj                | varchar(20)        | CNPJ do parceiro                                                          |
| parceiro            | varchar(100)       | Nome do parceiro                                                          |
| cod_top             | int                | Código da TOP (Tipo de Operação)                                       |
| top                 | varchar(150)       | Descrição da TOP                                                        |
| movimento           | char(1)            | Indica se há movimentação ('S', 'N')                                     |
| cliente             | char(1)            | Marca se é cliente ('S', 'N')                                             |
| fornecedor          | char(1)            | Marca se é fornecedor ('S', 'N')                                          |
| codigo              | int                | Código do item/produto                                                    |
| descricao           | text               | Descrição do item/produto                                               |
| ncm                 | varchar(20)        | Código NCM                                                               |
| grupo               | varchar(100)       | Grupo do produto                                                          |
| cfop                | int                | Código fiscal da operação                                               |
| operacao            | varchar(50)        | Tipo de operação                                                       |
| qtd_negociada       | decimal(10,3)      | Quantidade negociada                                                     |
| qtd_entregue        | decimal(10,3)      | Quantidade entregue                                                      |
| status              | char(1)            | Status do pedido ou da nota                                              |
| saldo               | decimal(10,3)      | Saldo entre quantidade negociada e entregue                              |
| valor_unitario      | decimal(15,2)      | Valor unitário do item                                                   |
| valor_total         | decimal(15,2)      | Valor total da negociação ou nota                                       |
| valor_icms          | decimal(15,2)      | Valor do ICMS                                                             |
| valor_ipi           | decimal(15,2)      | Valor do IPI                                                              |
| receita             | decimal(15,2)      | Valor de receita considerado para o faturamento                          |
