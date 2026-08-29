"""
AGROMV Semanal - atualizacao do painel de mercado
-------------------------------------------------
Busca indicadores em fontes publicas e grava data/mercado.json.
Roda pelo GitHub Actions; nao precisa de servidor.

Fontes (todas com acesso publico e uso permitido com credito):
  - Banco Central do Brasil, API SGS  -> dolar PTAX, Selic, IPCA
  - IBGE, API SIDRA (LSPA)            -> producao nacional de soja e milho

Regra de seguranca: se uma fonte falhar, o valor anterior e mantido
e marcado como desatualizado, em vez de sumir do site.
"""

import json
import os
import ssl
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAIDA = os.path.join(RAIZ, "data", "mercado.json")
FUSO = timezone(timedelta(hours=-3))  # horario de Brasilia

CTX = ssl.create_default_context()
CABECALHO = {"User-Agent": "AGROMV-Semanal/1.0 (+https://mvagro.net)"}


def buscar_json(url, timeout=25):
    req = urllib.request.Request(url, headers=CABECALHO)
    with urllib.request.urlopen(req, timeout=timeout, context=CTX) as r:
        return json.loads(r.read().decode("utf-8"))


# ----------------------------------------------------------------------
# BANCO CENTRAL - Sistema Gerenciador de Series Temporais
# ----------------------------------------------------------------------
SGS = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.{cod}/dados/ultimos/{n}?formato=json"


def serie_bcb(codigo, n=1):
    dados = buscar_json(SGS.format(cod=codigo, n=n))
    if not dados:
        raise ValueError("serie vazia")
    return dados


def dolar():
    """Dolar PTAX venda (serie 1), com variacao frente ao dia anterior."""
    d = serie_bcb(1, n=2)
    atual = float(d[-1]["valor"])
    item = {
        "rotulo": "Dólar PTAX",
        "valor": f"R$ {atual:.4f}".replace(".", ","),
        "detalhe": f"Cotação de venda · {d[-1]['data']}",
    }
    if len(d) > 1:
        anterior = float(d[-2]["valor"])
        if anterior:
            var = (atual - anterior) / anterior * 100
            item["variacao"] = f"{var:+.2f}%".replace(".", ",")
            item["sentido"] = "alta" if var > 0 else ("baixa" if var < 0 else "estavel")
    return item


def selic():
    """Meta Selic definida pelo Copom (serie 432)."""
    d = serie_bcb(432, n=1)[-1]
    return {
        "rotulo": "Selic (meta)",
        "valor": f"{float(d['valor']):.2f}%".replace(".", ",") + " a.a.",
        "detalhe": f"Vigente em {d['data']}",
    }


def ipca_12m():
    """IPCA acumulado em 12 meses, somado a partir da serie mensal (433)."""
    d = serie_bcb(433, n=12)
    acumulado = 1.0
    for m in d:
        acumulado *= 1 + float(m["valor"]) / 100
    return {
        "rotulo": "IPCA 12 meses",
        "valor": f"{(acumulado - 1) * 100:.2f}%".replace(".", ","),
        "detalhe": f"Até {d[-1]['data']}",
    }


# ----------------------------------------------------------------------
# IBGE - Levantamento Sistematico da Producao Agricola (LSPA)
# ----------------------------------------------------------------------
SIDRA = ("https://apisidra.ibge.gov.br/values/t/6588/n1/all/"
         "v/35/p/last%201/c48/{cultura}")

CULTURAS = {
    "soja":  ("39443", "Soja (produção nacional)"),
    "milho": ("39441", "Milho (produção nacional)"),
}


def producao_ibge(chave):
    codigo, rotulo = CULTURAS[chave]
    dados = buscar_json(SIDRA.format(cultura=codigo))
    linha = dados[1]  # a primeira linha traz os cabecalhos
    valor = float(linha["V"])
    periodo = linha.get("D3N", "")
    milhoes = valor / 1_000_000
    return {
        "rotulo": rotulo,
        "valor": f"{milhoes:,.1f} mi t".replace(",", "X").replace(".", ",").replace("X", "."),
        "detalhe": f"Estimativa LSPA · {periodo}",
    }


# ----------------------------------------------------------------------
COLETORES = [
    ("dolar", dolar),
    ("selic", selic),
    ("ipca", ipca_12m),
    ("soja", lambda: producao_ibge("soja")),
    ("milho", lambda: producao_ibge("milho")),
]


def carregar_anterior():
    try:
        with open(SAIDA, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"indicadores": {}}


def main():
    anterior = carregar_anterior()
    indicadores = {}
    falhas = []

    for chave, coletor in COLETORES:
        try:
            item = coletor()
            item["atualizado"] = True
            indicadores[chave] = item
            print(f"[ok]    {chave}: {item['valor']}")
        except Exception as e:
            falhas.append(chave)
            print(f"[falha] {chave}: {e}")
            antigo = anterior.get("indicadores", {}).get(chave)
            if antigo:                      # mantem o ultimo valor conhecido
                antigo["atualizado"] = False
                indicadores[chave] = antigo
                print(f"        mantido valor anterior: {antigo['valor']}")

    if not indicadores:
        raise SystemExit("nenhum indicador obtido; arquivo nao foi alterado")

    saida = {
        "atualizado_em": datetime.now(FUSO).strftime("%d/%m/%Y às %H:%M"),
        "fontes": "Banco Central do Brasil (SGS) e IBGE (LSPA/SIDRA)",
        "falhas": falhas,
        "indicadores": indicadores,
    }

    os.makedirs(os.path.dirname(SAIDA), exist_ok=True)
    with open(SAIDA, "w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, indent=2)
    print(f"\ngravado: {SAIDA}")
    if falhas:
        print("fontes que falharam nesta execucao:", ", ".join(falhas))


if __name__ == "__main__":
    main()
