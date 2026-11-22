"""Gera o relatório Excel (openpyxl), o gráfico PNG e o HTML do e-mail a partir dos dados extraídos."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import BarChart, PieChart, Reference  # noqa: E402
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
from openpyxl.worksheet.table import Table, TableStyleInfo  # noqa: E402

AMARELO, ESCURO, CINZA = "F2C811", "252423", "605E5C"


def resumo_por_categoria(df: pd.DataFrame) -> pd.DataFrame:
    r = (df.groupby("categoria")
         .agg(livros=("titulo", "count"), preco_medio=("preco_gbp", "mean"), nota_media=("nota", "mean"),
              em_estoque=("em_estoque", "mean"))
         .sort_values(["livros", "preco_medio"], ascending=[False, False])
         .reset_index())
    r["preco_medio"] = r["preco_medio"].round(2)
    r["nota_media"] = r["nota_media"].round(2)
    return r


def kpis(df: pd.DataFrame) -> dict:
    return {
        "Livros": len(df),
        "Categorias": df["categoria"].nunique(),
        "Preço médio (£)": round(df["preco_gbp"].mean(), 2),
        "Nota média": round(df["nota"].mean(), 2),
        "Em estoque": f"{df['em_estoque'].mean():.0%}",
    }


def _cabecalho(ws, linha, colunas):
    for c, titulo in enumerate(colunas, 1):
        cel = ws.cell(linha, c, titulo)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor=ESCURO)
        cel.alignment = Alignment(horizontal="center", vertical="center")


def gerar_excel(df: pd.DataFrame, destino: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Relatório do catálogo"
    ws["A1"].font = Font(size=18, bold=True, color=ESCURO)
    ws["A2"] = f"Gerado automaticamente em {datetime.now():%d/%m/%Y %H:%M} · fonte: books.toscrape.com"
    ws["A2"].font = Font(italic=True, color=CINZA)

    borda = Border(top=Side(style="thick", color=AMARELO))
    for i, (nome, valor) in enumerate(kpis(df).items()):
        col = 1 + i * 2
        ws.cell(4, col, nome).font = Font(color=CINZA)
        v = ws.cell(5, col, valor)
        v.font = Font(size=16, bold=True)
        ws.cell(4, col).border = borda

    r = resumo_por_categoria(df)
    inicio = 8
    _cabecalho(ws, inicio, ["Categoria", "Livros", "Preço médio (£)", "Nota média", "% em estoque"])
    for i, row in enumerate(r.itertuples(index=False), inicio + 1):
        ws.cell(i, 1, row.categoria)
        ws.cell(i, 2, row.livros)
        ws.cell(i, 3, row.preco_medio).number_format = "£#,##0.00"
        ws.cell(i, 4, row.nota_media).number_format = "0.00"
        ws.cell(i, 5, row.em_estoque).number_format = "0%"
    fim = inicio + len(r)
    ws.conditional_formatting.add(f"B{inicio + 1}:B{fim}", DataBarRule(start_type="min", end_type="max", color=AMARELO))
    ws.conditional_formatting.add(f"D{inicio + 1}:D{fim}", ColorScaleRule(
        start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84",
        end_type="max", end_color="63BE7B"))
    for col, w in zip("ABCDEFGHIJ", [26, 10, 16, 12, 14, 4, 4, 4, 4, 4]):
        ws.column_dimensions[col].width = w

    top = BarChart()
    top.type = "bar"
    top.title = "Top 10 categorias por quantidade"
    top.y_axis.title = "livros"
    top.add_data(Reference(ws, min_col=2, min_row=inicio, max_row=inicio + 10), titles_from_data=True)
    top.set_categories(Reference(ws, min_col=1, min_row=inicio + 1, max_row=inicio + 10))
    top.height, top.width = 9, 16
    top.legend = None
    ws.add_chart(top, "G8")

    notas = df["nota"].value_counts().sort_index()
    base = fim + 3
    ws.cell(base, 1, "Nota").font = Font(bold=True)
    ws.cell(base, 2, "Livros").font = Font(bold=True)
    for i, (n, q) in enumerate(notas.items(), base + 1):
        ws.cell(i, 1, f"{n} estrela{'s' if n > 1 else ''}")
        ws.cell(i, 2, int(q))
    pizza = PieChart()
    pizza.title = "Distribuição das notas"
    pizza.add_data(Reference(ws, min_col=2, min_row=base, max_row=base + len(notas)), titles_from_data=True)
    pizza.set_categories(Reference(ws, min_col=1, min_row=base + 1, max_row=base + len(notas)))
    pizza.height, pizza.width = 9, 12
    ws.add_chart(pizza, "G27")

    wl = wb.create_sheet("Livros")
    colunas = ["Título", "Categoria", "Preço (£)", "Nota", "Estrelas", "Em estoque", "Link"]
    wl.append(colunas)
    for row in df.sort_values(["categoria", "titulo"]).itertuples(index=False):
        wl.append([row.titulo, row.categoria, row.preco_gbp, row.nota, "★" * row.nota + "☆" * (5 - row.nota),
                   "sim" if row.em_estoque else "não", row.url])
    n = len(df) + 1
    tabela = Table(displayName="Livros", ref=f"A1:G{n}")
    tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium15", showRowStripes=True)
    wl.add_table(tabela)
    for i in range(2, n + 1):
        wl.cell(i, 3).number_format = "£#,##0.00"
        wl.cell(i, 7).hyperlink = wl.cell(i, 7).value
    wl.conditional_formatting.add(f"C2:C{n}", ColorScaleRule(start_type="min", start_color="FFFFFF",
                                                             end_type="max", end_color=AMARELO))
    wl.freeze_panes = "A2"
    for i, w in enumerate([52, 20, 11, 7, 10, 11, 60], 1):
        wl.column_dimensions[get_column_letter(i)].width = w

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)
    return destino


def grafico_png(df: pd.DataFrame, destino: Path) -> Path:
    r = resumo_por_categoria(df).head(10).iloc[::-1]
    plt.rcParams["font.family"] = ["Segoe UI", "DejaVu Sans"]
    fig, ax = plt.subplots(figsize=(8, 4.2), dpi=150)
    fig.patch.set_facecolor("#1F1E1D")
    ax.set_facecolor("#1F1E1D")
    barras = ax.barh(r["categoria"], r["livros"], color="#F2C811", height=.62)
    ax.bar_label(barras, padding=4, color="#F3F2F1", fontsize=9)
    ax.set_title("Top 10 categorias por quantidade de livros", color="#F3F2F1", loc="left", fontsize=12, weight="bold")
    ax.tick_params(colors="#D2D0CE", labelsize=9)
    for s in ax.spines.values():
        s.set_visible(False)
    ax.xaxis.set_visible(False)
    fig.tight_layout()
    fig.savefig(destino, facecolor=fig.get_facecolor())
    plt.close(fig)
    return destino


def gerar_email(df: pd.DataFrame, src_grafico: str, inicio: datetime) -> str:
    k = kpis(df)
    cards = "".join(
        f'<td style="padding:14px 18px;background:#252423;border-top:3px solid #F2C811;border-radius:6px">'
        f'<div style="font-size:24px;font-weight:800;color:#F3F2F1">{v}</div>'
        f'<div style="font-size:12px;color:#A19F9D">{n}</div></td><td width="10"></td>' for n, v in k.items())
    r = resumo_por_categoria(df).head(5)
    linhas = "".join(
        f'<tr><td style="padding:6px 10px;border-bottom:1px solid #3B3A39">{x.categoria}</td>'
        f'<td style="padding:6px 10px;border-bottom:1px solid #3B3A39;text-align:right">{x.livros}</td>'
        f'<td style="padding:6px 10px;border-bottom:1px solid #3B3A39;text-align:right">£{x.preco_medio:.2f}</td></tr>'
        for x in r.itertuples())
    dur = (datetime.now() - inicio).total_seconds()
    return f"""<!doctype html><html><body style="margin:0;background:#161514;font-family:Segoe UI,Arial,sans-serif;color:#F3F2F1">
<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center" style="padding:28px">
<table width="760" cellpadding="0" cellspacing="0" style="background:#1F1E1D;border:1px solid #3B3A39;border-radius:10px">
<tr><td style="padding:26px 28px 6px">
  <div style="font-size:12px;letter-spacing:2px;color:#F2C811;font-weight:700">RELATÓRIO AUTOMÁTICO</div>
  <div style="font-size:24px;font-weight:800;margin-top:6px">Catálogo de livros · {datetime.now():%d/%m/%Y}</div>
  <div style="font-size:13px;color:#A19F9D;margin-top:4px">Extraído por robô em {dur:.0f} s · planilha completa em anexo</div>
</td></tr>
<tr><td style="padding:18px 28px"><table cellpadding="0" cellspacing="0"><tr>{cards}</tr></table></td></tr>
<tr><td style="padding:0 28px"><img src="{src_grafico}" width="704" style="display:block;border-radius:6px" alt="gráfico"></td></tr>
<tr><td style="padding:18px 28px 28px">
  <div style="font-size:14px;font-weight:700;margin-bottom:8px">Maiores categorias</div>
  <table width="100%" cellpadding="0" cellspacing="0" style="font-size:13px;color:#D2D0CE">
    <tr style="color:#A19F9D"><td style="padding:6px 10px">Categoria</td><td style="padding:6px 10px;text-align:right">Livros</td>
    <td style="padding:6px 10px;text-align:right">Preço médio</td></tr>{linhas}</table>
</td></tr></table></td></tr></table></body></html>"""
