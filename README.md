# 📊 Gerador de Dashboard Financeiro

Script em Python que lê um CSV de transações (receitas e despesas) e gera
automaticamente um arquivo Excel com uma aba de dados e um **dashboard
com KPIs e gráficos** — sem precisar montar nada manualmente na planilha.

Nasceu do uso pessoal: controlar as contas de casa (mercado, contas fixas,
pagamentos). Serve também para pequenos negócios que querem ver receitas,
despesas e saldo sem montar planilha na mão.

> Os arquivos `dados_exemplo.csv` e `dashboard_exemplo.xlsx` usam **dados fictícios**.

## O que ele gera

Na aba **Dashboard**:
- Cartões de KPI: Receitas, Despesas, Saldo e Ticket médio de venda
- Gráfico de barras: despesas por categoria
- Gráfico de linha: evolução do saldo mês a mês

Na aba **Dados**: todos os lançamentos organizados e formatados.

## Como usar

1. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

2. Prepare seu CSV com as colunas: `Data, Descricao, Categoria, Valor, Tipo`
   (veja `dados_exemplo.csv` como modelo — `Tipo` é `Receita` ou `Despesa`,
   e valores de despesa devem ser negativos).

3. Rode o script:
   ```bash
   python gerar_dashboard.py dados_exemplo.csv dashboard.xlsx
   ```

4. Abra o `dashboard.xlsx` gerado — o dashboard já vem pronto na segunda aba.

## Stack

Python · pandas · openpyxl

## Autor

Leonardo Reis de Freitas Andrade — [LinkedIn](https://www.linkedin.com/in/leonardo-reis-cloud) · leonardo.brasil.ra@gmail.com
