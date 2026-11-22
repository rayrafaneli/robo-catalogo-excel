"""
Robô de catálogo -> relatório Excel + e-mail.

Navega o catálogo público https://books.toscrape.com (site feito para treinar extração de dados)
categoria por categoria, seguindo a paginação como um usuário faria, e gera:

  saida/catalogo.xlsx   planilha com aba de resumo (KPIs, tabela por categoria, gráficos nativos do Excel)
                        e aba com os 1.000 livros (tabela com filtro, links e formatação condicional)
  saida/email.html      resumo pronto para envio, com gráfico embutido
  saida/livros.csv      dados brutos

Uso:
    python robo.py                     # extrai e gera os arquivos
    python robo.py --demo              # também salva quadros da navegação (GIF do README)
    python robo.py --enviar            # envia o e-mail via SMTP (variáveis SMTP_* no ambiente)
"""
from __future__ import annotations

import argparse
import logging
import os
import re
import smtplib
from dataclasses import asdict, dataclass
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import pandas as pd
from playwright.sync_api import Page, sync_playwright

from relatorio import gerar_email, gerar_excel, grafico_png

BASE = "https://books.toscrape.com/"
SAIDA = Path(__file__).resolve().parent / "saida"
NOTAS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

log = logging.getLogger("robo")


@dataclass
class Livro:
    titulo: str
    categoria: str
    preco_gbp: float
    nota: int
    em_estoque: bool
    url: str


def preco(texto: str) -> float:
    """'£51.77' -> 51.77 (aceita o 'Â£' que aparece com encoding errado)."""
    m = re.search(r"(\d+(?:\.\d+)?)", texto.replace(",", ""))
    if not m:
        raise ValueError(f"preço inválido: {texto!r}")
    return float(m.group(1))


def nota(classes: str) -> int:
    for nome, valor in NOTAS.items():
        if nome in classes.split():
            return valor
    return 0


def categorias(page: Page) -> list[tuple[str, str]]:
    page.goto(BASE, wait_until="domcontentloaded")
    links = page.locator(".side_categories ul li ul li a")
    return [(links.nth(i).inner_text().strip(), links.nth(i).get_attribute("href")) for i in range(links.count())]


def livros_da_pagina(page: Page, categoria: str) -> list[Livro]:
    itens = page.eval_on_selector_all(
        "article.product_pod",
        """els => els.map(e => ({
            titulo: e.querySelector('h3 a').getAttribute('title'),
            href: e.querySelector('h3 a').href,
            preco: e.querySelector('.price_color').textContent,
            nota: e.querySelector('.star-rating').className,
            estoque: e.querySelector('.availability').textContent.trim(),
        }))""")
    return [Livro(i["titulo"], categoria, preco(i["preco"]), nota(i["nota"]), "In stock" in i["estoque"], i["href"])
            for i in itens]


def extrair(page: Page, foto=None) -> list[Livro]:
    todos: list[Livro] = []
    cats = categorias(page)
    log.info("%d categorias encontradas", len(cats))
    for n, (nome, href) in enumerate(cats, 1):
        page.goto(BASE + href, wait_until="domcontentloaded")
        while True:
            todos.extend(livros_da_pagina(page, nome))
            if foto:
                page.wait_for_load_state("load")  # no modo demo, espera as capas para o quadro sair completo
                foto(f"categoria {n}/{len(cats)} · {len(todos)} livros")
            proxima = page.locator("li.next a")
            if not proxima.count():
                break
            proxima.click()
            page.wait_for_load_state("domcontentloaded")
        log.info("%-20s %4d livros no total", nome[:20], len(todos))
    return todos


class Quadros:
    def __init__(self, page: Page, pasta: Path):
        self.page, self.pasta, self.selos = page, pasta, []
        pasta.mkdir(parents=True, exist_ok=True)
        for f in pasta.glob("*.jpg"):
            f.unlink()

    def __call__(self, selo: str) -> None:
        self.selos.append(selo)
        self.page.screenshot(path=str(self.pasta / f"{len(self.selos):04d}.jpg"), type="jpeg", quality=85)
        (self.pasta / "selos.txt").write_text("\n".join(self.selos), encoding="utf-8")


def enviar_email(html: str, png: Path) -> None:
    msg = EmailMessage()
    msg["Subject"] = f"Relatório do catálogo · {datetime.now():%d/%m/%Y}"
    msg["From"] = os.environ["SMTP_USER"]
    msg["To"] = os.environ["EMAIL_PARA"]
    msg.set_content("Seu cliente de e-mail não exibe HTML. Veja o anexo catalogo.xlsx.")
    msg.add_alternative(html, subtype="html")
    msg.get_payload()[1].add_related(png.read_bytes(), "image", "png", cid="<grafico>")
    msg.add_attachment((SAIDA / "catalogo.xlsx").read_bytes(), maintype="application",
                       subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="catalogo.xlsx")
    with smtplib.SMTP(os.environ.get("SMTP_HOST", "smtp.gmail.com"), int(os.environ.get("SMTP_PORT", "587"))) as s:
        s.starttls()
        s.login(os.environ["SMTP_USER"], os.environ["SMTP_SENHA"])
        s.send_message(msg)
    log.info("e-mail enviado para %s", msg["To"])


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--demo", action="store_true", help="salva quadros da navegação em saida/quadros")
    ap.add_argument("--enviar", action="store_true", help="envia o e-mail via SMTP")
    ap.add_argument("--canal", default=None, help="canal do navegador (ex.: msedge, chrome)")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", datefmt="%H:%M:%S")
    SAIDA.mkdir(exist_ok=True)

    inicio = datetime.now()
    with sync_playwright() as p:
        browser = p.chromium.launch(channel=args.canal, headless=True)
        page = browser.new_page(viewport={"width": 1280, "height": 800})
        foto = Quadros(page, SAIDA / "quadros") if args.demo else None
        livros = extrair(page, foto)
        browser.close()

    df = pd.DataFrame([asdict(l) for l in livros]).drop_duplicates(subset="url")
    df.to_csv(SAIDA / "livros.csv", index=False, encoding="utf-8")
    gerar_excel(df, SAIDA / "catalogo.xlsx")
    png = grafico_png(df, SAIDA / "grafico.png")
    html = gerar_email(df, "cid:grafico", inicio)
    (SAIDA / "email.html").write_text(html.replace("cid:grafico", "grafico.png"), encoding="utf-8")
    log.info("%d livros, %d categorias em %.0f s", len(df), df["categoria"].nunique(),
             (datetime.now() - inicio).total_seconds())
    if args.enviar:
        enviar_email(html, png)


if __name__ == "__main__":
    main()
