"""
Gerador de Dashboard Financeiro
--------------------------------
Le um CSV de transacoes (Data, Descricao, Categoria, Valor, Tipo) e gera
um arquivo Excel com uma aba de dados e uma aba de dashboard, com KPIs
e graficos (despesas por categoria e evolucao mensal do saldo).

Uso:
    python gerar_dashboard.py dados_exemplo.csv saida_dashboard.xlsx

Autor: Leonardo Reis de Freitas Andrade
"""

import sys
from datetime import datetime
from collections import defaultdict

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter

NAVY = "122A42"
TEAL = "1FAE9E"
LIGHT = "EAF7F4"
WHITE = "FFFFFF"


def carregar_dados(caminho_csv: str) -> pd.DataFrame:
    df = pd.read_csv(caminho_csv, parse_dates=["Data"])
    df["Valor"] = df["Valor"].astype(float)
    df["Mes"] = df["Data"].dt.to_period("M").astype(str)
    return df


def calcular_kpis(df: pd.DataFrame) -> dict:
    receitas = df.loc[df["Tipo"] == "Receita", "Valor"].sum()
    despesas = df.loc[df["Tipo"] == "Despesa", "Valor"].sum()
    saldo = receitas + despesas
    ticket_medio = df.loc[df["Tipo"] == "Receita", "Valor"].mean()
    return {
        "Receitas": receitas,
        "Despesas": abs(despesas),
        "Saldo": saldo,
        "Ticket medio de venda": ticket_medio,
    }


def montar_planilha(df: pd.DataFrame, kpis: dict, caminho_saida: str) -> None:
    wb = Workbook()

    # ---------- Aba Dados ----------
    ws_dados = wb.active
    ws_dados.title = "Dados"
    ws_dados.append(["Data", "Descricao", "Categoria", "Valor (R$)", "Tipo", "Mes"])
    for cell in ws_dados[1]:
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center")

    for _, row in df.sort_values("Data").iterrows():
        ws_dados.append([
            row["Data"].strftime("%d/%m/%Y"),
            row["Descricao"],
            row["Categoria"],
            row["Valor"],
            row["Tipo"],
            row["Mes"],
        ])

    for col, width in zip("ABCDEF", [12, 26, 16, 14, 10, 10]):
        ws_dados.column_dimensions[col].width = width
    for r in range(2, ws_dados.max_row + 1):
        ws_dados[f"D{r}"].number_format = "#,##0.00"

    # ---------- Aba Dashboard ----------
    ws = wb.create_sheet("Dashboard")
    ws["B2"] = "Dashboard Financeiro"
    ws["B2"].font = Font(bold=True, size=16, color=NAVY)
    ws["B3"] = f"Gerado em {datetime.now().strftime('%d/%m/%Y')} a partir de {len(df)} lancamentos"
    ws["B3"].font = Font(italic=True, size=9, color="51657A")

    # KPIs como cartoes
    kpi_items = [
        ("Receitas", kpis["Receitas"], TEAL),
        ("Despesas", kpis["Despesas"], "C0392B"),
        ("Saldo", kpis["Saldo"], NAVY),
        ("Ticket medio de venda", kpis["Ticket medio de venda"], "51657A"),
    ]
    start_col = 2  # B
    for i, (label, valor, cor) in enumerate(kpi_items):
        col = start_col + i * 2
        col_letter = get_column_letter(col)
        ws.merge_cells(f"{col_letter}5:{get_column_letter(col+1)}5")
        ws.merge_cells(f"{col_letter}6:{get_column_letter(col+1)}7")
        label_cell = ws[f"{col_letter}5"]
        label_cell.value = label
        label_cell.font = Font(bold=True, size=9, color="51657A")
        value_cell = ws[f"{col_letter}6"]
        value_cell.value = valor
        value_cell.number_format = '"R$" #,##0.00'
        value_cell.font = Font(bold=True, size=14, color=cor)
        value_cell.alignment = Alignment(vertical="center")
        for row in range(5, 8):
            for c in range(col, col + 2):
                ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=LIGHT)

    # Tabela auxiliar: despesas por categoria
    despesas_cat = (
        df.loc[df["Tipo"] == "Despesa"]
        .assign(Valor=lambda d: d["Valor"].abs())
        .groupby("Categoria")["Valor"].sum()
        .sort_values(ascending=False)
    )
    ws["B10"] = "Despesas por categoria"
    ws["B10"].font = Font(bold=True, color=NAVY)
    ws["B11"] = "Categoria"
    ws["C11"] = "Valor (R$)"
    ws["B11"].font = ws["C11"].font = Font(bold=True)
    r = 12
    for cat, valor in despesas_cat.items():
        ws[f"B{r}"] = cat
        ws[f"C{r}"] = valor
        ws[f"C{r}"].number_format = "#,##0.00"
        r += 1
    cat_fim = r - 1

    # Tabela auxiliar: saldo por mes
    mensal = df.groupby("Mes")["Valor"].sum().sort_index()
    base_col = 7  # G
    ws.cell(row=10, column=base_col, value="Saldo por mes").font = Font(bold=True, color=NAVY)
    ws.cell(row=11, column=base_col, value="Mes").font = Font(bold=True)
    ws.cell(row=11, column=base_col + 1, value="Saldo (R$)").font = Font(bold=True)
    r = 12
    for mes, valor in mensal.items():
        ws.cell(row=r, column=base_col, value=mes)
        c = ws.cell(row=r, column=base_col + 1, value=valor)
        c.number_format = "#,##0.00"
        r += 1
    mensal_fim = r - 1

    # Grafico de barras: despesas por categoria
    bar = BarChart()
    bar.title = "Despesas por categoria"
    bar.y_axis.title = "R$"
    bar.style = 10
    data_ref = Reference(ws, min_col=3, min_row=11, max_row=cat_fim)
    cats_ref = Reference(ws, min_col=2, min_row=12, max_row=cat_fim)
    bar.add_data(data_ref, titles_from_data=True)
    bar.set_categories(cats_ref)
    bar.width, bar.height = 15, 8
    ws.add_chart(bar, f"B{cat_fim + 3}")

    # Grafico de linha: saldo mensal
    line = LineChart()
    line.title = "Saldo mensal"
    line.y_axis.title = "R$"
    line.style = 12
    data_ref2 = Reference(ws, min_col=base_col + 1, min_row=11, max_row=mensal_fim)
    cats_ref2 = Reference(ws, min_col=base_col, min_row=12, max_row=mensal_fim)
    line.add_data(data_ref2, titles_from_data=True)
    line.set_categories(cats_ref2)
    line.width, line.height = 15, 8
    ws.add_chart(line, f"G{mensal_fim + 3}")

    for col, width in zip("ABCDEFGHI", [2, 22, 14, 14, 14, 14, 14, 14, 14]):
        ws.column_dimensions[col].width = width

    wb.save(caminho_saida)


def main():
    if len(sys.argv) != 3:
        print("Uso: python gerar_dashboard.py <entrada.csv> <saida.xlsx>")
        sys.exit(1)

    entrada, saida = sys.argv[1], sys.argv[2]
    df = carregar_dados(entrada)
    kpis = calcular_kpis(df)
    montar_planilha(df, kpis, saida)

    print(f"Dashboard gerado em: {saida}")
    print(f"Receitas: R$ {kpis['Receitas']:.2f}")
    print(f"Despesas: R$ {kpis['Despesas']:.2f}")
    print(f"Saldo:    R$ {kpis['Saldo']:.2f}")


if __name__ == "__main__":
    main()
