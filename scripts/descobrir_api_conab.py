"""
Descoberta da API do portal de precos da CONAB - versao 2
==========================================================
https://consultaprecosdemercado.conab.gov.br/

Mudancas em relacao a v1, que nao capturou nada:
  - registra TODAS as requisicoes, nao apenas xhr/fetch (o filtro anterior
    era estreito demais e pode ter descartado a chamada de dados)
  - clica nas abas "Precos medios semanais" e "mensais", que provavelmente
    e o que dispara a carga
  - preenche o intervalo de datas antes de clicar em Pesquisar
  - registra falhas de requisicao e erros de console

Apenas observa o que a propria aplicacao faz. Nao contorna seguranca.
"""

import re
import sys
from datetime import date, timedelta

from playwright.sync_api import sync_playwright

SITE = "https://consultaprecosdemercado.conab.gov.br/"

ESTATICO = re.compile(
    r"\.(js|css|png|jpe?g|gif|svg|webp|woff2?|ttf|eot|ico|map)(\?|$)", re.I
)

eventos = []      # requisicoes com resposta
falhas = []       # requisicoes que nao completaram
console = []      # mensagens de erro do aplicativo


def util(url):
    return not ESTATICO.search(url) and not url.startswith("data:")


def ao_responder(resposta):
    req = resposta.request
    if not util(req.url):
        return
    registro = {
        "url": req.url,
        "metodo": req.method,
        "tipo": req.resource_type,
        "status": resposta.status,
        "content_type": resposta.headers.get("content-type", "?"),
        "corpo_envio": req.post_data,
        "resposta": None,
    }
    # le o corpo sempre: servidores costumam devolver JSON com content-type
    # generico, e descartar por causa disso ja nos custou uma rodada
    try:
        texto = resposta.text()
        registro["resposta"] = texto[:2500]
        if len(texto) > 2500:
            registro["resposta"] += f"\n... [+{len(texto)-2500} caracteres]"
    except Exception as e:
        registro["resposta"] = f"[nao lido: {type(e).__name__}]"
    eventos.append(registro)


def ao_falhar(req):
    if util(req.url):
        falhas.append(f"{req.method} {req.url} -> {req.failure}")


def despejar(titulo):
    print("\n" + "=" * 74)
    print(titulo)
    print("=" * 74)
    if not eventos:
        print("(nenhuma requisicao)")
    for i, e in enumerate(eventos, 1):
        print(f"\n--- [{i}] {e['metodo']} {e['status']} ({e['tipo']}) ---")
        print(f"URL: {e['url']}")
        print(f"Content-Type: {e['content_type']}")
        if e["corpo_envio"]:
            print(f"Body enviado: {e['corpo_envio'][:600]}")
        if e["resposta"]:
            print(f"Resposta:\n{e['resposta']}")
    if falhas:
        print("\nrequisicoes que falharam:")
        for f in falhas:
            print("   ", f)
    eventos.clear()
    falhas.clear()


def clicar(pagina, texto, espera=6000):
    """Tenta clicar por varias estrategias e sempre informa o resultado."""
    padrao = re.compile(texto, re.I)
    tentativas = [
        ("papel de botao", lambda: pagina.get_by_role("button", name=padrao).first),
        ("papel de aba",   lambda: pagina.get_by_role("tab", name=padrao).first),
        ("texto visivel",  lambda: pagina.get_by_text(padrao).first),
        ("seletor css",    lambda: pagina.locator(f"text={texto}").first),
    ]
    for nome, obter in tentativas:
        try:
            alvo = obter()
            if alvo.count() == 0:
                continue
            alvo.click(timeout=6000)
            print(f">>> clicou em '{texto}' (via {nome})")
            pagina.wait_for_timeout(espera)
            return True
        except Exception as e:
            print(f">>> '{texto}' via {nome}: {type(e).__name__}")
    print(f">>> NAO foi possivel clicar em '{texto}'")
    return False


def main():
    with sync_playwright() as p:
        navegador = p.chromium.launch(args=["--no-sandbox"])
        ctx = navegador.new_context(locale="pt-BR")
        pagina = ctx.new_page()
        pagina.on("response", ao_responder)
        pagina.on("requestfailed", ao_falhar)
        pagina.on("console", lambda m: console.append(f"{m.type}: {m.text[:200]}")
                  if m.type in ("error", "warning") else None)

        print(f"abrindo {SITE}")
        pagina.goto(SITE, wait_until="networkidle", timeout=90000)
        pagina.wait_for_timeout(5000)
        despejar("FASE 1 - CARREGAMENTO INICIAL (todas as requisicoes)")

        # ---- Fase 2: aba de precos semanais ----
        clicar(pagina, r"semanais")
        despejar("FASE 2 - APOS ABRIR A ABA SEMANAL")

        # ---- Fase 3: informar periodo e pesquisar ----
        fim = date.today()
        inicio = fim - timedelta(days=30)
        periodo = f"{inicio.strftime('%d/%m/%Y')} até {fim.strftime('%d/%m/%Y')}"
        try:
            campo = pagina.locator("#range-input")
            campo.click(timeout=8000)
            campo.fill(periodo, timeout=8000)
            pagina.keyboard.press("Escape")
            print(f">>> periodo preenchido: {periodo}")
        except Exception as e:
            print(f">>> nao preencheu o periodo: {type(e).__name__}: {e}")

        pagina.wait_for_timeout(1500)
        try:
            pagina.get_by_role("button", name=re.compile("Pesquisar", re.I)).first.click(timeout=8000)
            print(">>> clicou em Pesquisar")
        except Exception as e:
            print(f">>> nao clicou em Pesquisar: {type(e).__name__}")
        pagina.wait_for_timeout(9000)
        despejar("FASE 3 - APOS PESQUISAR (semanais)")

        # ---- Fase 4: aba mensal ----
        clicar(pagina, r"mensais")
        try:
            pagina.get_by_role("button", name=re.compile("Pesquisar", re.I)).first.click(timeout=8000)
            pagina.wait_for_timeout(9000)
        except Exception:
            pass
        despejar("FASE 4 - APOS ABRIR A ABA MENSAL")

        # ---- diagnostico extra ----
        print("\n" + "=" * 74)
        print("MENSAGENS DO APLICATIVO")
        print("=" * 74)
        for c in console[:25]:
            print("   ", c)
        if not console:
            print("(nenhuma)")

        print("\n" + "=" * 74)
        print("TEXTO VISIVEL DA PAGINA (inicio)")
        print("=" * 74)
        try:
            print(pagina.inner_text("body")[:1500])
        except Exception as e:
            print(f"[nao lido: {type(e).__name__}]")

        navegador.close()

    print("\nfim da investigacao")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERRO: {type(e).__name__}: {e}")
        sys.exit(1)
