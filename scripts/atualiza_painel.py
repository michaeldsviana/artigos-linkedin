"""
AGROMV Semanal - atualizacao do painel de mercado
--------------------------------------------------
Gera data/mercado.json com tres grupos:

  precos  -> dolar (BCB) + culturas do Cerrado (CONAB) + fertilizantes (Banco Mundial)
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
import time
import time
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


def buscar_json(url, timeout=30, tentativas=3):
    """Le JSON com nova tentativa apenas em falha passageira.

    502, 503 e timeout costumam ser instabilidade momentanea e vale repetir.
    Ja 404 ou 403 nao mudam com insistencia, entao abortamos na hora para
    nao gastar minutos de execucao a toa.
    """
    PASSAGEIROS = (429, 500, 502, 503, 504)
    ultimo = None
    for n in range(tentativas):
        try:
            with abrir(url, timeout) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception as e:
            ultimo = e
            codigo = getattr(e, "code", None)
            if codigo is not None and codigo not in PASSAGEIROS:
                raise                      # erro permanente: nao adianta repetir
            if n < tentativas - 1:
                espera = 3 * (n + 1)
                print(f"        tentativa {n+1} falhou ({type(e).__name__}"
                      f"{f' {codigo}' if codigo else ''}); repetindo em {espera}s")
                time.sleep(espera)
    raise ultimo


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
        "fonte": "bcb",
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

# A tabela do Banco Mundial passa a cobrir apenas fertilizantes: as culturas
# vem da CONAB, com preco nacional em reais e recorte do Cerrado.
COMMODITIES = {
    "Soja (int.)":     ["soybean"],
    "Milho (int.)":    ["maize"],
    "Algodão (int.)":  ["cotton"],
    "Ureia":           ["urea"],
    "DAP":             ["dap"],
    "Fosfato (rocha)": ["phosphate rock"],
    "Potássio (KCl)":  ["potassium chloride"],
}

def _letras(ref):
    return re.match(r"[A-Z]+", ref).group(0)


def _indice(letras):
    n = 0
    for c in letras:
        n = n * 26 + (ord(c) - 64)
    return n - 1


def _ns(raiz):
    """Extrai o namespace do proprio arquivo.

    Arquivos .xlsx podem usar o namespace padrao (schemas.openxmlformats.org)
    ou o do formato estrito (purl.oclc.org). Fixar um deles quebra o outro.
    """
    if raiz.tag.startswith("{"):
        return raiz.tag.split("}")[0] + "}"
    return ""


def _escolher_aba(z):
    """Devolve o caminho da aba de precos mensais.

    A pasta de trabalho lista as abas por nome; preferimos a que fala de
    precos mensais e evitamos as de indices. Sem essa informacao, cai na
    primeira aba do arquivo.
    """
    caminhos = sorted(n for n in z.namelist()
                      if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
    try:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        nsw = _ns(wb)
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        nsr = _ns(rels)
        destino = {r.get("Id"): r.get("Target") for r in rels.iter(f"{nsr}Relationship")}

        abas = []
        for s in wb.iter(f"{nsw}sheet"):
            rid = next((v for k, v in s.attrib.items() if k.endswith("}id")), None)
            alvo = destino.get(rid, "")
            if alvo:
                alvo = "xl/" + alvo.lstrip("/").replace("worksheets/", "worksheets/")
                if not alvo.startswith("xl/worksheets"):
                    alvo = "xl/" + alvo.split("xl/")[-1]
            abas.append((s.get("name", ""), alvo))
        print("        abas: " + ", ".join(n for n, _ in abas))

        # 1a escolha: aba de precos mensais
        for nome, alvo in abas:
            baixo = nome.lower()
            if "price" in baixo and "month" in baixo and alvo in z.namelist():
                print(f"        usando aba '{nome}'")
                return alvo
        # 2a escolha: qualquer aba de precos que nao seja de indices
        for nome, alvo in abas:
            baixo = nome.lower()
            if "price" in baixo and "ind" not in baixo and alvo in z.namelist():
                print(f"        usando aba '{nome}'")
                return alvo
        # 3a escolha: mensal que nao seja indice
        for nome, alvo in abas:
            baixo = nome.lower()
            if "month" in baixo and "ind" not in baixo and alvo in z.namelist():
                print(f"        usando aba '{nome}'")
                return alvo
        for nome, alvo in abas:
            if alvo in z.namelist():
                print(f"        usando aba '{nome}' (primeira disponivel)")
                return alvo
    except Exception as e:
        print(f"        nao foi possivel ler a lista de abas ({type(e).__name__}); usando a primeira")
    return caminhos[0]


def ler_planilha(conteudo):
    """Devolve a aba de precos mensais do .xlsx como matriz de strings."""
    z = zipfile.ZipFile(io.BytesIO(conteudo))

    compartilhadas = []
    if "xl/sharedStrings.xml" in z.namelist():
        raiz = ET.fromstring(z.read("xl/sharedStrings.xml"))
        ns = _ns(raiz)
        for si in raiz.findall(f"{ns}si"):
            compartilhadas.append("".join(t.text or "" for t in si.iter(f"{ns}t")))

    caminho = _escolher_aba(z)
    raiz = ET.fromstring(z.read(caminho))
    ns = _ns(raiz)

    linhas = []
    for row in raiz.iter(f"{ns}row"):
        celulas = {}
        for c in row.findall(f"{ns}c"):
            texto = None
            if c.get("t") == "inlineStr":
                bloco = c.find(f"{ns}is")
                if bloco is not None:
                    texto = "".join(t.text or "" for t in bloco.iter(f"{ns}t"))
            else:
                v = c.find(f"{ns}v")
                if v is not None and v.text is not None:
                    texto = v.text
                    if c.get("t") == "s":
                        i = int(texto)
                        texto = compartilhadas[i] if i < len(compartilhadas) else ""
            if texto is None:
                continue
            ref = c.get("r")
            if not ref:
                continue
            celulas[_indice(_letras(ref))] = texto
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
    for i, linha in enumerate(linhas[:40]):
        texto = " | ".join(linha).lower()
        p = sum(1 for t in alvos if t in texto)
        if p > pontos:
            melhor, pontos = i, p
    if melhor is None or pontos == 0:
        amostra = [" | ".join(l)[:140] for l in linhas[:8]]
        raise ValueError("cabecalho nao encontrado; primeiras linhas: " + str(amostra))
    print(f"        cabecalho na linha {melhor + 1} ({pontos} termos reconhecidos)")

    cabecalho = [c.lower() for c in linhas[melhor]]

    # Logo abaixo do cabecalho a planilha traz a unidade de cada coluna,
    # do tipo ($/mt) ou ($/kg). Sem isso, soja em tonelada e algodao em
    # quilo apareceriam lado a lado como se fossem comparaveis.
    unidades = {}
    for linha in linhas[melhor + 1: melhor + 4]:
        for j, celula in enumerate(linha):
            if j in unidades or not celula:
                continue
            m = re.search(r"\$\s*/\s*([A-Za-z]+)", str(celula))
            if m:
                unidades[j] = m.group(1).lower()
    if unidades:
        print(f"        unidades reconhecidas em {len(unidades)} colunas")
    else:
        print("        [aviso] unidades nao encontradas na planilha")

    NOMES_UNIDADE = {
        "mt": "tonelada", "kg": "quilo", "bbl": "barril",
        "mmbtu": "milhão de BTU", "dmtu": "tonelada seca", "oz": "onça troy",
    }
    SIMBOLO = {"mt": "/t", "kg": "/kg", "bbl": "/bbl",
               "mmbtu": "/mmBTU", "dmtu": "/dmtu", "oz": "/oz"}

    # A planilha vem em tonelada (ou quilo). Convertemos as culturas para a
    # unidade usada na comercializacao no Brasil e deixamos os fertilizantes
    # em tonelada, que e como sao negociados.
    # culturas convertidas para a unidade de comercializacao;
    # fertilizantes permanecem em tonelada
    CONVERSAO = {
        "Soja (int.)":    ("saca de 60 kg", 60, "/sc"),
        "Milho (int.)":   ("saca de 60 kg", 60, "/sc"),
        "Algodão (int.)": ("arroba de 15 kg", 15, "/@"),
    }

    def para_kg(valor, unidade):
        if unidade == "mt":
            return valor / 1000.0
        if unidade == "kg":
            return valor
        return None

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

        unidade = unidades.get(j, "")
        sufixo = SIMBOLO.get(unidade, "")
        por_extenso = NOMES_UNIDADE.get(unidade)
        exibido = atual

        conversao = CONVERSAO.get(rotulo)
        preco_kg = para_kg(atual, unidade) if conversao else None

        if conversao and preco_kg is not None:
            nome_alvo, kg_alvo, simbolo_alvo = conversao
            exibido = preco_kg * kg_alvo
            sufixo = simbolo_alvo
            original = "US$ " + br(atual) + SIMBOLO.get(unidade, "")
            detalhe = (f"Por {nome_alvo} · média de {periodo} · "
                       f"Banco Mundial ({original})")
        else:
            detalhe = f"Média de {periodo} · Banco Mundial"
            if por_extenso:
                detalhe = f"Por {por_extenso} · média de {periodo} · Banco Mundial"

        item = {"rotulo": rotulo, "valor": "US$ " + br(exibido) + sufixo,
                "detalhe": detalhe, "fonte": "banco-mundial"}
        if unidade:
            item["unidade"] = unidade
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

    # Cada linha do IBGE e mantida como veio: 1a e 2a safra de milho, as tres
    # safras de feijao e os tipos de cafe tem comportamento de mercado
    # diferente, entao somar mascararia a leitura.
    itens = []
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
        limpo = re.sub(r"^[\d.]+\s*", "", nome)               # tira "1.15 "
        limpo = re.sub(r"\s*\((em grão|em casca|em caroço)\)", "", limpo, flags=re.I)
        limpo = re.sub(r"\s{2,}", " ", limpo).strip()
        limpo = limpo[:1].upper() + limpo[1:]
        itens.append((limpo, valor, periodo))

    if not itens:
        raise ValueError("nenhuma cultura reconhecida na resposta")

    itens.sort(key=lambda x: -x[1])
    saida = {}
    for nome, valor, periodo in itens:
        saida[nome] = {
            "rotulo": nome,
            "valor": br(valor / 1_000_000, 1) + " mi t",
            "detalhe": f"Estimativa LSPA · {periodo}",
        }
    return saida



# ======================================================================
# 4. CONAB - precos das culturas no Cerrado
#    A CONAB autoriza reproducao sem fins lucrativos, citada a fonte e
#    mantida a integridade das informacoes.
# ======================================================================
BASE_CONAB = "https://portaldeinformacoes.conab.gov.br/downloads/arquivos/"

# o portal e feito em JavaScript, entao testamos os nomes provaveis do
# arquivo que alimenta o painel de precos semanais por UF
ARQUIVOS_CONAB = [
    "PrecosAgropecuariosSemanalUF.csv",
    "PrecosAgropecuariosSemanalUf.csv",
    "precos_agropecuarios_semanal_uf.csv",
    "PrecosSemanalUF.csv",
    "PrecosAgropecuariosMensalUF.csv",
    "SerieHistoricaPrecos.csv",
    "PrecosAgropecuarios.csv",
]

# estados do bioma Cerrado
UF_CERRADO = ["MT", "GO", "MS", "MG", "BA", "TO", "MA", "PI", "DF"]

CULTURAS_CONAB = ["soja", "milho", "algod", "sorgo", "feij"]


def baixar_conab():
    """Percorre os nomes candidatos ate encontrar um CSV de verdade."""
    erros = []
    for nome in ARQUIVOS_CONAB:
        url = BASE_CONAB + nome
        try:
            with abrir(url, timeout=90) as r:
                bruto = r.read()
            if b"<!doctype" in bruto[:200].lower() or b"<html" in bruto[:200].lower():
                erros.append(f"{nome}: veio HTML")
                print(f"        [{nome}] pagina do app, nao e dado")
                continue
            print(f"        [{nome}] arquivo obtido: {len(bruto)} bytes")
            return nome, bruto
        except Exception as e:
            erros.append(f"{nome}: {type(e).__name__}")
            print(f"        [{nome}] {type(e).__name__}")
    raise ValueError("nenhum arquivo da CONAB respondeu -> " + " | ".join(erros))


def _decodificar(bruto):
    for cod in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return bruto.decode(cod)
        except UnicodeDecodeError:
            continue
    return bruto.decode("utf-8", "ignore")


def _achar(colunas, termos):
    for i, nome in enumerate(colunas):
        baixo = nome.lower().strip()
        if any(t in baixo for t in termos):
            return i
    return None


def precos_conab():
    """Ultimo preco por cultura e UF do Cerrado."""
    import csv as _csv

    nome_arq, bruto = baixar_conab()
    texto = _decodificar(bruto)

    # o separador varia entre ; e , nos arquivos do governo
    amostra = texto[:2000]
    separador = ";" if amostra.count(";") > amostra.count(",") else ","
    leitor = _csv.reader(io.StringIO(texto), delimiter=separador)

    linhas = [l for l in leitor if l]
    if len(linhas) < 2:
        raise ValueError("arquivo sem linhas de dados")

    colunas = linhas[0]
    print(f"        colunas: {colunas[:10]}")

    i_prod = _achar(colunas, ["produto", "cultura"])
    i_uf   = _achar(colunas, ["uf", "estado", "sigla"])
    i_data = _achar(colunas, ["data", "periodo", "semana", "mes", "ano"])
    i_val  = _achar(colunas, ["preco", "preço", "valor"])
    i_uni  = _achar(colunas, ["unidade", "embalagem"])

    if i_prod is None or i_val is None:
        raise ValueError(f"colunas essenciais nao identificadas em {colunas[:12]}")

    registros = {}
    for linha in linhas[1:]:
        if len(linha) <= max(x for x in [i_prod, i_val, i_uf, i_data] if x is not None):
            continue
        produto = linha[i_prod].strip()
        baixo = produto.lower()
        if not any(c in baixo for c in CULTURAS_CONAB):
            continue

        uf = linha[i_uf].strip().upper() if i_uf is not None else ""
        if uf and uf not in UF_CERRADO:
            continue

        bruto_valor = linha[i_val].strip().replace(".", "").replace(",", ".")
        try:
            valor = float(bruto_valor)
        except ValueError:
            continue
        if valor <= 0:
            continue

        data = linha[i_data].strip() if i_data is not None else ""
        unidade = linha[i_uni].strip() if i_uni is not None else ""

        chave = f"{produto} {uf}".strip()
        anterior = registros.get(chave)
        if anterior is None or data > anterior["data"]:
            registros[chave] = {"produto": produto, "uf": uf, "valor": valor,
                                "data": data, "unidade": unidade}

    if not registros:
        raise ValueError("nenhum registro de cultura no Cerrado encontrado")

    saida = {}
    for chave in sorted(registros, key=lambda k: (registros[k]["produto"], registros[k]["uf"])):
        r = registros[chave]
        rotulo = f"{r['produto']} {r['uf']}".strip()
        unidade = r["unidade"] or "unidade da fonte"
        saida[rotulo] = {
            "rotulo": rotulo,
            "valor": "R$ " + br(r["valor"]),
            "detalhe": f"Por {unidade} · {r['data']} · CONAB",
            "fonte": "conab",
        }
    print(f"        {len(saida)} precos do Cerrado ({nome_arq})")
    return saida


# ======================================================================
# cada coletor declara a etiqueta de fonte que produz, para que uma falha
# restaure somente os proprios valores antigos, sem misturar bases
COLETA = {
    "precos": [("dolar", dolar, "bcb"),
               ("culturas_conab", precos_conab, "conab"),
               ("fertilizantes", commodities_banco_mundial, "banco-mundial")],
    "macro":  [("selic", selic, None), ("ipca", ipca_12m, None)],
    "safra":  [("culturas", safra_ibge, "ibge")],
}


# ----------------------------------------------------------------------
# SONDAGEM: localizar a API da CONAB
# O JavaScript do portal aponta para "localhost:8001/api/v1", ou seja, o
# endereco real e injetado em tempo de execucao. Procuramos o arquivo de
# configuracao e os caminhos de rota dentro do codigo. Apenas diagnostico.
# ----------------------------------------------------------------------
PORTAL_CONAB = "https://portaldeinformacoes.conab.gov.br/"

CONFIGS = [
    "assets/config.json", "assets/config/config.json",
    "assets/environment.json", "assets/env.json",
    "assets/data/config.json", "config.json", "assets/appsettings.json",
]

BASES_API = [
    "https://portaldeinformacoes.conab.gov.br/api/v1/",
    "https://apiportaldeinformacoes.conab.gov.br/api/v1/",
    "https://portaldeinformacoes.conab.gov.br/backend/api/v1/",
    "https://sisdep.conab.gov.br/api/v1/",
]


LIMITE_LEITURA = 1_500_000   # 1,5 MB por arquivo ja cobre os pacotes do portal


def _texto(url, timeout=20, limite=LIMITE_LEITURA):
    """Le no maximo `limite` bytes: os pacotes do portal sao grandes e a
    sondagem nao pode arrastar a execucao do painel."""
    with abrir(url, timeout) as r:
        return r.read(limite).decode("utf-8", "ignore"), r.headers.get("Content-Type", "?")


def sondar(orcamento=90):
    """Diagnostico com tempo limitado: o painel ja foi gravado antes daqui."""
    inicio = time.time()

    def esgotou():
        return time.time() - inicio > orcamento

    print("\n--- sondagem: procurando a API da CONAB ---")

    # 1. arquivos de configuracao do aplicativo
    print("\n[1] arquivos de configuracao")
    for caminho in CONFIGS:
        if esgotou():
            print("    (tempo esgotado)")
            break
        url = PORTAL_CONAB + caminho
        try:
            corpo, tipo = _texto(url, 12)
            eh_html = "<!doctype" in corpo[:200].lower() or "<html" in corpo[:200].lower()
            if eh_html:
                print(f"    [-] {caminho}: pagina do app")
            else:
                print(f"    [OK] {caminho} ({tipo})")
                print(f"         {corpo[:400]}")
        except Exception as e:
            print(f"    [x] {caminho}: {type(e).__name__}")

    # 2. rotas citadas no codigo do aplicativo
    print("\n[2] rotas encontradas no JavaScript")
    try:
        html, _ = _texto(PORTAL_CONAB, 20)
        scripts = re.findall(r'src="([^"]+\.js)"', html)
        chaves = ("preco", "safra", "prohort", "arquivo", "download",
                  "serie", "produto", "cultura", "api/")
        achados = set()
        for src in scripts:
            if "polyfill" in src.lower():
                continue
            url = src if src.startswith("http") else PORTAL_CONAB + src.lstrip("/")
            try:
                codigo, _ = _texto(url, 25)
            except Exception:
                continue
            for texto in re.findall(r'"([^"\\]{4,120})"', codigo):
                baixo = texto.lower()
                if any(k in baixo for k in chaves) and " " not in texto:
                    achados.add(texto)
        for a in sorted(achados)[:70]:
            print("    ", a)
        if not achados:
            print("     nenhuma rota reconhecida")
    except Exception as e:
        print(f"    [x] {type(e).__name__}: {e}")

    # 3. bases de API candidatas
    print("\n[3] bases de API")
    for base in BASES_API:
        if esgotou():
            print("    (tempo esgotado)")
            break
        try:
            corpo, tipo = _texto(base, 12)
            eh_html = "<!doctype" in corpo[:200].lower()
            print(f"    [{'-' if eh_html else 'OK'}] {base} ({tipo})")
            if not eh_html:
                print(f"         {corpo[:200]}")
        except Exception as e:
            print(f"    [x] {base}: {type(e).__name__}")

    print("\n--- fim da sondagem ---\n")


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
        for chave, funcao, etiqueta in coletores:
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
                # so recupera o que veio da mesma fonte; assim um preco antigo
                # em dolar nao reaparece no lugar de um preco em reais
                for k, item in anterior.get(grupo, {}).items():
                    if k in grupos[grupo]:
                        continue
                    if etiqueta and item.get("fonte") != etiqueta:
                        continue
                    if not etiqueta and chave != k:
                        continue
                    item["atualizado"] = False
                    grupos[grupo][k] = item
                    print(f"        mantido anterior: {k} = {item['valor']}")

    if not any(grupos.values()):
        raise SystemExit("nenhum dado obtido; arquivo preservado")

    saida = {
        "atualizado_em": datetime.now(FUSO).strftime("%d/%m/%Y às %H:%M"),
        "fontes": ("CONAB (preços das culturas) · Banco Central do Brasil (câmbio) · "
                   "IBGE/LSPA (safra) · Banco Mundial, Commodity Markets, CC BY 4.0 "
                   "(fertilizantes)"),
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

    # diagnostico opcional, com tempo limitado; o painel ja esta salvo
    try:
        sondar()
    except Exception as e:
        print("sondagem nao concluida:", e)


if __name__ == "__main__":
    main()
