"""
Descoberta da API real do portal de preços da CONAB
====================================================
https://consultaprecosdemercado.conab.gov.br/

Abre a pagina em um navegador de verdade, interage com os filtros e registra
TODAS as requisicoes XHR/fetch disparadas, com URL, metodo, cabecalhos, corpo
e resposta. E o equivalente automatizado ao painel Network do DevTools.

Nao faz scraping de HTML e nao contorna nenhum mecanismo de seguranca:
apenas observa as chamadas que a propria aplicacao realiza.

Roda no GitHub Actions, onde a rede e aberta.
Uso: python scripts/descobrir_api_conab.py
"""

import json
import re
import sys

from playwright.sync_api import sync_playwright

SITE = "https://consultaprecosdemercado.conab.gov.br/"

# requisicoes de infraestrutura que nao interessam ao diagnostico
IGNORAR = re.compile(
    r"\.(js|css|png|jpe?g|gif|svg|woff2?|ttf|eot|ico|map)(\?|$)", re.I
)

capturadas = []


def interessa(url):
    return not IGNORAR.search(url)


def resumir_corpo(texto, limite=1800):
    if texto is None:
        return None
    if len(texto) <= limite:
        return texto
    return texto[:limite] + f"\n... [+{len(texto) - limite} caracteres]"


def registrar(resposta):
    """Guarda cada resposta de dados, com o maximo de contexto util."""
    req = resposta.request
    if req.resource_type not in ("xhr", "fetch") or not interessa(req.url):
        return
    try:
        corpo = resposta.text()
    except Exception as e:
        corpo = f"[corpo nao lido: {type(e).__name__}]"

    capturadas.append({
        "url": req.url,
        "metodo": req.method,
        "status": resposta.status,
        "content_type": resposta.headers.get("content-type", "?"),
        "cabecalhos_requisicao": {
            k: v for k, v in req.headers.items()
            if k.lower() in ("accept", "content-type", "referer", "origin",
                             "authorization", "x-requested-with")
        },
        "corpo_requisicao": resumir_corpo(req.post_data, 900),
        "resposta": resumir_corpo(corpo),
    })


def imprimir(titulo, itens):
    print("\n" + "=" * 74)
    print(titulo)
    print("=" * 74)
    if not itens:
        print("(nenhuma requisicao de dados capturada)")
        return
    for i, c in enumerate(itens, 1):
        print(f"\n--- [{i}] {c['metodo']} {c['status']} ---")
        print(f"URL: {c['url']}")
        print(f"Content-Type: {c['content_type']}")
        if c["cabecalhos_requisicao"]:
            print(f"Headers: {json.dumps(c['cabecalhos_requisicao'], ensure_ascii=False)}")
        if c["corpo_requisicao"]:
            print(f"Body enviado: {c['corpo_requisicao']}")
        print(f"Resposta:\n{c['resposta']}")


def descrever_controles(pagina):
    """Lista os campos do formulario, para sabermos o que da para acionar."""
    print("\n" + "=" * 74)
    print("CONTROLES DISPONIVEIS NA PAGINA")
    print("=" * 74)
    for seletor in ("select", "input", "button", "mat-select", "[role=combobox]"):
        try:
            elementos = pagina.query_selector_all(seletor)
        except Exception:
            continue
        if not elementos:
            continue
        print(f"\n<{seletor}> encontrados: {len(elementos)}")
        for el in elementos[:15]:
            try:
                texto = (el.inner_text() or "").strip().replace("\n", " ")[:60]
                nome = el.get_attribute("name") or el.get_attribute("formcontrolname") or ""
                idt = el.get_attribute("id") or ""
                ph = el.get_attribute("placeholder") or ""
                aria = el.get_attribute("aria-label") or ""
                print(f"   texto='{texto}' name='{nome}' id='{idt}' "
                      f"placeholder='{ph}' aria='{aria}'")
            except Exception:
                continue


def main():
    with sync_playwright() as p:
        navegador = p.chromium.launch(args=["--no-sandbox"])
        contexto = navegador.new_context(
            locale="pt-BR",
            user_agent=("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/125.0 Safari/537.36"),
        )
        pagina = contexto.new_page()
        pagina.on("response", registrar)

        print(f"abrindo {SITE}")
        pagina.goto(SITE, wait_until="networkidle", timeout=90000)
        pagina.wait_for_timeout(4000)

        imprimir("FASE 1 - REQUISICOES DO CARREGAMENTO INICIAL", list(capturadas))
        descrever_controles(pagina)

        # ---- Fase 2: acionar os filtros e observar o que muda ----
        capturadas.clear()
        print("\n" + "=" * 74)
        print("FASE 2 - INTERAGINDO COM OS FILTROS")
        print("=" * 74)

        # tenta preencher combos nativos; se o portal usa componentes proprios,
        # a fase 1 ja mostra as chamadas de carga dos filtros
        try:
            selects = pagina.query_selector_all("select")
            print(f"combos nativos: {len(selects)}")
            for idx, s in enumerate(selects[:4]):
                opcoes = s.query_selector_all("option")
                rotulos = [(o.get_attribute("value"), (o.inner_text() or "").strip())
                           for o in opcoes[:12]]
                print(f"   combo {idx}: {rotulos}")
                alvo = None
                for valor, texto in rotulos:
                    baixo = texto.lower()
                    if "soja" in baixo or baixo == "mt" or "mato grosso" in baixo:
                        alvo = valor
                        break
                if alvo:
                    s.select_option(alvo)
                    pagina.wait_for_timeout(2500)
                    print(f"   -> selecionado '{alvo}' no combo {idx}")
        except Exception as e:
            print(f"combos nativos indisponiveis: {type(e).__name__}: {e}")

        # aciona qualquer botao que pareca "consultar"
        for texto in ("Consultar", "Pesquisar", "Buscar", "Filtrar", "Aplicar"):
            try:
                botao = pagina.get_by_role("button", name=re.compile(texto, re.I)).first
                if botao and botao.is_visible():
                    botao.click()
                    print(f"clicou em '{texto}'")
                    pagina.wait_for_timeout(5000)
                    break
            except Exception:
                continue

        pagina.wait_for_timeout(3000)
        imprimir("FASE 2 - REQUISICOES APOS INTERACAO", list(capturadas))

        navegador.close()

    print("\nfim da investigacao")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERRO: {type(e).__name__}: {e}")
        sys.exit(1)
