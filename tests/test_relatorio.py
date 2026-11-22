import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import relatorio  # noqa: E402
import robo  # noqa: E402


@pytest.mark.parametrize("texto,valor", [("£51.77", 51.77), ("Â£13.99", 13.99), ("£1,000.00", 1000.0)])
def test_preco(texto, valor):
    assert robo.preco(texto) == valor


def test_preco_invalido():
    with pytest.raises(ValueError):
        robo.preco("grátis")


def test_nota():
    assert robo.nota("star-rating Three") == 3
    assert robo.nota("star-rating") == 0


@pytest.fixture
def df():
    return pd.DataFrame([
        {"titulo": "A", "categoria": "Poesia", "preco_gbp": 10.0, "nota": 5, "em_estoque": True, "url": "http://x/a"},
        {"titulo": "B", "categoria": "Poesia", "preco_gbp": 20.0, "nota": 3, "em_estoque": True, "url": "http://x/b"},
        {"titulo": "C", "categoria": "Viagem", "preco_gbp": 30.0, "nota": 1, "em_estoque": False, "url": "http://x/c"},
    ])


def test_resumo_por_categoria(df):
    r = relatorio.resumo_por_categoria(df)
    assert list(r["categoria"]) == ["Poesia", "Viagem"]
    assert r.loc[0, "livros"] == 2 and r.loc[0, "preco_medio"] == 15.0 and r.loc[0, "nota_media"] == 4.0


def test_excel_tem_abas_graficos_e_linhas(df, tmp_path):
    destino = relatorio.gerar_excel(df, tmp_path / "c.xlsx")
    wb = load_workbook(destino)
    assert wb.sheetnames == ["Resumo", "Livros"]
    assert len(wb["Resumo"]._charts) == 2
    assert wb["Livros"].max_row == len(df) + 1
    assert "Livros" in wb["Livros"].tables


def test_email_tem_kpis(df):
    html = relatorio.gerar_email(df, "cid:grafico", datetime.now())
    assert "cid:grafico" in html and "Poesia" in html and ">3<" in html
