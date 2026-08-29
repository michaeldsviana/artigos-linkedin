"""
AGROMV Semanal - atualizacao do painel de mercado
--------------------------------------------------
Gera data/mercado.json com tres grupos:

  precos  -> dolar (BCB) + commodities e fertilizantes (Banco Mundial)
  macro   -> Selic e IPCA (BCB)
  safra   -> projecao nacional de producao por cultura (IBGE / LSPA)

Fontes e licencas:
  Banco Central do Brasil - dados publicos
  IBGE, Levantamento Sistematico da Producao Agricola - dados publicos
  Banco Mundial, Commodity Markets (Pink Sheet) - CC BY 4.0, exige citacao

Roda pelo GitHub Actions, sem servidor e sem dependencia externa:
o arquivo .xlsx do Banco Mundial e lido com a biblioteca padrao do Python.

Se uma fonte falhar, o valor anterior e mantido e marcado como defasado.
"""

import io
import json
import os
import re
import ssl
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAIDA = os.path.join(RAIZ, "data", "mercado.json")
FUSO = timezone(timedelta(hours=-3))

CTX = ssl.create_default_context()
CABECALHO = {"User-Agent": "AGROMV-Semanal/1.0 (+https://mvagro.net)"}


def abrir(url, timeout=60):
    req = urllib.request.Request(url, headers=CABECALHO)
    return urllib.request.urlopen(req, timeout=timeout, context=CTX)


def buscar_json(url, timeout=30):
    with abrir(url, timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def br(valor, casas=2):
    """Formata numero no padrao brasileiro."""
    txt = f"{valor:,.{casas}f}"
    return txt.replace(",", "X").replace(".", ",").replace("X", ".")


# ======================================================================
# 1. BANCO CENTRAL
# ======================================================================
SGS = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.{cod}/dados/ultimos/{n}?formato=json"


def serie_bcb(codigo, n=1):
    dados = buscar_json(SGS.format(cod=codigo, n=n))
    if not dados:
        raise ValueError("serie vazia")
    return dados


def dolar_olinda():
    hoje = datetime.now(FUSO)
    inicio = (hoje - timedelta(days=12)).strftime("%m-%d-%Y")
    fim = hoje.strftime("%m-%d-%Y")
    url = (
        "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
        "CotacaoDolarPeriodo(dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)"
        f"?@dataInicial='{inicio}'&@dataFinalCotacao='{fim}'"
        "&$top=10&$orderby=dataHoraCotacao%20desc&$format=json"
    )
    dados = buscar_json(url).get("value", [])
    if not dados:
        raise ValueError("Olinda retornou vazio")
    atual = float(dados[0]["cotacaoVenda"])
    anterior = float(dados[1]["cotacaoVenda"]) if len(dados) > 1 else None
    a, m, d = dados[0]["dataHoraCotacao"][:10].split("-")
    return atual, anterior, f"{d}/{m}/{a}"


def dolar():
    try:
        d = serie_bcb(1, n=2)
        atual = float(d[-1]["valor"])
        data_ref = d[-1]["data"]
        anterior = float(d[-2]["valor"]) if len(d) > 1 else None
    except Exception as e:
        print(f"        SGS indisponivel ({type(e).__name__}); usando Olinda")
        atual, anterior, data_ref = dolar_olinda()

    item = {
        "rotulo": "Dólar",
        "valor": "R$ " + br(atual, 4),
        "detalhe": f"PTAX venda · {data_ref}",
    }
    if anterior:
        var = (atual - anterior) / anterior * 100
        item["variacao"] = f"{var:+.2f}%".replace(".", ",")
        item["sentido"] = "alta" if var > 0 else ("baixa" if var < 0 else "estavel")
    return item


def selic():
    d = serie_bcb(432, n=1)[-1]
    return {"rotulo": "Selic (meta)", "valor": br(float(d["valor"])) + "% a.a.",
            "detalhe": f"Vigente em {d['data']}"}


def ipca_12m():
    d = serie_bcb(433, n=12)
    acum = 1.0
    for m in d:
        acum *= 1 + float(m["valor"]) / 100
    return {"rotulo": "IPCA 12 meses", "valor": br((acum - 1) * 100) + "%",
            "detalhe": f"Até {d[-1]['data']}"}


# ======================================================================
# 2. BANCO MUNDIAL - commodities e fertilizantes
#    O arquivo e um .xlsx; lemos com zipfile + XML, sem instalar nada.
# ======================================================================
URLS_BM = [
    # edicao corrente (2026)
    ("2026", "https://thedocs.worldbank.org/en/doc/"
             "74e8be41ceb20fa0da750cda2f6b9e4e-0050012026/related/"
             "CMO-Historical-Data-Monthly.xlsx"),
    # edicao anterior (2025)
    ("2025", "https://thedocs.worldbank.org/en/doc/"
             "18675f1d1639c7a34d463f59263ba0a2-0050012025/related/"
             "CMO-Historical-Data-Monthly.xlsx"),
    # endereco antigo, mantido pelo Banco Mundial por compatibilidade
    ("pubdocs", "https://pubdocs.worldbank.org/en/561011486076393416/"
                "CMO-Historical-Data-Monthly.xlsx"),
]


def baixar_planilha():
    """Tenta os enderecos conhecidos ate um responder com um .xlsx valido.

    O caminho do arquivo muda a cada edicao anual, entao a lista precisa
    ter alternativas. Um .xlsx sempre comeca com a assinatura PK.
    """
    erros = []
    for nome, url in URLS_BM:
        try:
            with abrir(url, timeout=90) as r:
                conteudo = r.read()
            if conteudo[:2] != b"PK":
                erros.append(f"{nome}: resposta nao e xlsx ({conteudo[:16]!r})")
                print(f"        [{nome}] resposta invalida")
                continue
            print(f"        [{nome}] planilha obtida: {len(conteudo)} bytes")
            return conteudo
        except Exception as e:
            erros.append(f"{nome}: {type(e).__name__}: {e}")
            print(f"        [{nome}] {type(e).__name__}: {e}")
    raise ValueError("nenhum endereco respondeu -> " + " | ".join(erros))

# rotulo exibido -> termos que identificam a coluna na planilha
COMMODITIES = {
    "Soja":            ["soybean"],
    "Milho":           ["maize"],
    "Algodão":         ["cotton"],
    "Ureia":           ["urea"],
    "DAP":             ["dap"],
    "Fosfato (rocha)": ["phosphate rock"],
}

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _letras(ref):
    return re.match(r"[A-Z]+", ref).group(0)


def _indice(letras):
    n = 0
    for c in letras:
        n = n * 26 + (ord(c) - 64)
    return n - 1


def ler_planilha(conteudo):
    """Devolve a primeira aba do .xlsx como matriz de strings."""
    z = zipfile.ZipFile(io.BytesIO(conteudo))

    compartilhadas = []
    if "xl/sharedStrings.xml" in z.namelist():
        raiz = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in raiz.findall(f"{NS}si"):
            compartilhadas.append("".join(t.text or "" for t in si.iter(f"{NS}t")))

    nomes = [n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
    raiz = ET.fromstring(z.read(sorted(nomes)[0]))

    linhas = []
    for row in raiz.iter(f"{NS}row"):
        celulas = {}
        for c in row.findall(f"{NS}c"):
            texto = None
            if c.get("t") == "inlineStr":          # texto embutido na celula
                bloco = c.find(f"{NS}is")
                if bloco is not None:
                    texto = "".join(t.text or "" for t in bloco.iter(f"{NS}t"))
            else:
                v = c.find(f"{NS}v")
                if v is not None and v.text is not None:
                    texto = v.text
                    if c.get("t") == "s":          # texto na tabela compartilhada
                        i = int(texto)
                        texto = compartilhadas[i] if i < len(compartilhadas) else ""
            if texto is None:
                continue
            celulas[_indice(_letras(c.get("r")))] = texto
        if celulas:
            largura = max(celulas) + 1
            linhas.append([celulas.get(i, "") for i in range(largura)])
    return linhas


def commodities_banco_mundial():
    """Ultimo valor mensal de cada commodity, com variacao frente ao mes anterior."""
    conteudo = baixar_planilha()
    linhas = ler_planilha(conteudo)
    print(f"        linhas lidas: {len(linhas)}")

    alvos = [t for lista in COMMODITIES.values() for t in lista]
    melhor, pontos = None, 0
    for i, linha in enumerate(linhas[:25]):
        texto = " | ".join(linha).lower()
        p = sum(1 for t in alvos if t in texto)
        if p > pontos:
            melhor, pontos = i, p
    if melhor is None or pontos == 0:
        amostra = [" | ".join(l)[:120] for l in linhas[:6]]
        raise ValueError("cabecalho nao encontrado; primeiras linhas: " + str(amostra))
    print(f"        cabecalho na linha {melhor + 1} ({pontos} termos reconhecidos)")

    cabecalho = [c.lower() for c in linhas[melhor]]

    def coluna_de(termos):
        for j, nome in enumerate(cabecalho):
            if all(t in nome for t in termos):
                return j
        return None

    dados = [l for l in linhas[melhor + 1:]
             if l and re.match(r"^\d{4}M\d{2}", str(l[0]).strip())]
    if len(dados) < 2:
        raise ValueError("nenhuma linha de periodo encontrada")

    ultima, penultima = dados[-1], dados[-2]
    ano, mes = str(ultima[0]).strip().split("M")
    periodo = f"{mes}/{ano}"

    saida = {}
    for rotulo, termos in COMMODITIES.items():
        j = coluna_de(termos)
        if j is None or j >= len(ultima):
            print(f"        [pulado] {rotulo}: coluna nao localizada")
            continue
        try:
            atual = float(ultima[j])
        except (ValueError, IndexError):
            print(f"        [pulado] {rotulo}: valor invalido")
            continue

        item = {"rotulo": rotulo, "valor": "US$ " + br(atual),
                "detalhe": f"Média de {periodo} · Banco Mundial"}
        try:
            ant = float(penultima[j])
            if ant:
                var = (atual - ant) / ant * 100
                item["variacao"] = f"{var:+.1f}%".replace(".", ",")
                item["sentido"] = "alta" if var > 0 else ("baixa" if var < 0 else "estavel")
        except (ValueError, IndexError):
            pass
        saida[rotulo] = item

    if not saida:
        raise ValueError("nenhuma commodity extraida")
    return saida


# ======================================================================
# 3. IBGE - projecao de safra nacional (LSPA)
# ======================================================================
URL_SIDRA = "https://apisidra.ibge.gov.br/values/t/6588/n1/all/v/35/p/last%201/c48/all"

CULTURAS = ["soja", "milho", "algod", "arroz", "feij", "trigo", "café", "cafe"]


def safra_ibge():
    dados = buscar_json(URL_SIDRA, timeout=60)
    if len(dados) < 2:
        raise ValueError("SIDRA retornou vazio")

    saida = {}
    for linha in dados[1:]:
        nome = (linha.get("D4N") or linha.get("D3N") or "").strip()
        chave = nome.lower()
        if not any(c in chave for c in CULTURAS) or "total" in chave:
            continue
        try:
            valor = float(linha.get("V"))
        except (TypeError, ValueError):
            continue
        periodo = linha.get("D3N") or linha.get("D2N") or ""
        limpo = re.sub(r"\s*\(.*?\)", "", nome).strip()
        saida[limpo] = {
            "rotulo": limpo,
            "valor": br(valor / 1_000_000, 1) + " mi t",
            "detalhe": f"Estimativa LSPA · {periodo}",
        }
    if not saida:
        raise ValueError("nenhuma cultura reconhecida na resposta")
    return saida


# ======================================================================
COLETA = {
    "precos": [("dolar", dolar), ("commodities", commodities_banco_mundial)],
    "macro":  [("selic", selic), ("ipca", ipca_12m)],
    "safra":  [("culturas", safra_ibge)],
}


def carregar_anterior():
    try:
        with open(SAIDA, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"grupos": {}}


def main():
    anterior = carregar_anterior().get("grupos", {})
    grupos = {"precos": {}, "macro": {}, "safra": {}}
    falhas = []

    for grupo, coletores in COLETA.items():
        for chave, funcao in coletores:
            try:
                r = funcao()
                itens = r if "rotulo" not in r else {chave: r}
                for k, item in itens.items():
                    item["atualizado"] = True
                    grupos[grupo][k] = item
                    print(f"[ok]    {grupo}/{k}: {item['valor']}")
            except Exception as e:
                falhas.append(f"{grupo}/{chave}")
                print(f"[falha] {grupo}/{chave}: {type(e).__name__}: {e}")
                for k, item in anterior.get(grupo, {}).items():
                    if k not in grupos[grupo]:
                        item["atualizado"] = False
                        grupos[grupo][k] = item
                        print(f"        mantido anterior: {k} = {item['valor']}")

    if not any(grupos.values()):
        raise SystemExit("nenhum dado obtido; arquivo preservado")

    saida = {
        "atualizado_em": datetime.now(FUSO).strftime("%d/%m/%Y às %H:%M"),
        "fontes": ("Banco Central do Brasil · IBGE (LSPA) · "
                   "Banco Mundial, Commodity Markets (CC BY 4.0)"),
        "falhas": falhas,
        "grupos": grupos,
    }

    os.makedirs(os.path.dirname(SAIDA), exist_ok=True)
    with open(SAIDA, "w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, indent=2)

    total = sum(len(v) for v in grupos.values())
    print(f"\ngravado: {SAIDA} ({total} indicadores)")
    if falhas:
        print("falharam nesta execucao:", ", ".join(falhas))


if __name__ == "__main__":
    main()
