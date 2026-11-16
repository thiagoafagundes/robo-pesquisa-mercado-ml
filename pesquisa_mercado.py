# -*- coding: utf-8 -*-
"""
ROBÔ DE PESQUISA DE MERCADO - MERCADO LIVRE  (versão 2)
=======================================================
O que ele faz:
  1. Escolhe sozinho as subcategorias do seu NICHO (elétrico, ferramentas,
     iluminação, energia...) ou as categorias que você indicar.
  2. Pega os 20 produtos MAIS VENDIDOS de cada uma.
  3. Para cada produto verifica: preço, quantos vendedores concorrem, se a loja
     oficial domina, se tem Full e se pode exigir certificação do Inmetro.
  4. Guarda um HISTÓRICO: rodando toda semana, ele mostra quantas vezes cada
     produto apareceu no topo (o que fica no topo é tendência; o resto é moda).
  5. Dá NOTA (0-100) e VEREDITO: VENDER / ANALISAR / EVITAR, com o motivo.
  6. Gera planilha Excel com CALCULADORA DE MARGEM: você preenche só o custo
     do fornecedor e ela calcula taxa do ML, frete, imposto, lucro e margem.

Como usar no computador:
  pip install requests openpyxl
  Preencha APP_ID e SECRET_KEY abaixo e rode:  python pesquisa_mercado.py
  Para ver um exemplo SEM credenciais (dados falsos):
        python pesquisa_mercado.py --simular
"""

import base64
import csv
import hashlib
import json
import secrets
import os
import random
import re
import sys
import time
import unicodedata
from datetime import datetime
from urllib.parse import urlencode, urlparse, parse_qs

import requests
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================
# 1) SUAS CONFIGURAÇÕES  (edite aqui)
# ============================================================
APP_ID = ""          # "Client ID" / "App ID" do seu aplicativo no Mercado Livre
SECRET_KEY = ""      # "Client Secret" / "Secret Key"
REDIRECT_URI = "https://www.google.com/"   # a MESMA cadastrada no aplicativo

# Categorias específicas (IDs). Vazio = o robô escolhe pelo NICHO abaixo.
CATEGORIAS = []
INCLUIR_SUBCATEGORIAS = False   # só vale quando você preenche CATEGORIAS

# Palavras do seu nicho. O robô procura subcategorias com esses nomes.
# Deixe a lista vazia para pesquisar TODAS as categorias principais.
NICHO = ["eletric", "ferrament", "iluminac", "lampada", "energia", "solar",
         "tomada", "interruptor", "disjuntor", "cabo", "fios", "carregador",
         "automac", "led", "extens", "gerador", "pilhas"]

# Categorias que batem com o NICHO por acaso e devem ser ignoradas
# (ex.: "bateria" de instrumento musical, telefone "sem fio").
EXCLUIR_CATEGORIAS = ["instrumentos musicais", "percuss", "telefonia",
                      "cameras", "games", "brinquedos", "beleza"]

# Regras de seleção (as que combinamos)
PRECO_MIN = 50.0        # abaixo disso a taxa fixa do ML come a margem
PRECO_MAX = 300.0       # acima disso prende muito capital em estoque
MAX_VENDEDORES = 20     # concorrência aceitável no catálogo

# Valores usados na calculadora de margem (você pode mudar também na planilha,
# na aba "Parâmetros"). CONFIRA OS VALORES ATUAIS NO MERCADO LIVRE.
TAXA_ML = 0.12                 # anúncio Clássico: 10 a 14% conforme a categoria
CUSTO_FIXO_ABAIXO_79 = 6.75    # custo por unidade em vendas abaixo de R$ 79 (aprox.)
FRETE_ACIMA_79 = 20.0          # frete que você paga acima de R$ 79 (estimativa)
IMPOSTO = 0.04                 # Simples Nacional, comércio, faixa inicial (MEI: 0)
MARGEM_MINIMA = 0.20           # 20%

# Produtos que costumam exigir certificação do Inmetro
PALAVRAS_INMETRO = ["tomada", "plugue", "adaptador", "extensao", "filtro de linha",
                    "disjuntor", "interruptor", "carregador", "fonte", "lampada",
                    "cabo", "fio", "benjamim", "chuveiro", "reator", "luminaria"]

ARQUIVO_TOKEN = "token_ml.json"
ARQUIVO_HISTORICO = "historico_mais_vendidos.csv"
PAUSA_ENTRE_CHAMADAS = 0.15   # segundos (para não ser bloqueado)

API = "https://api.mercadolibre.com"
AUTH = "https://auth.mercadolivre.com.br/authorization"
SITE = "MLB"


def norm(texto):
    """minúsculas e sem acento, para comparar palavras."""
    t = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return t.lower()


def contem_palavra(texto, palavras):
    t = norm(texto)
    return any(re.search(r"(^|[^a-z])" + re.escape(norm(p)), t) for p in palavras)


# ============================================================
# 2) LOGIN NO MERCADO LIVRE (token de acesso)
# ============================================================
ARQUIVO_PKCE = "pkce_verifier.txt"


def _novo_pkce():
    """Gera o par PKCE (verifier secreto + challenge enviado no link)."""
    verifier = secrets.token_urlsafe(64)[:96]
    challenge = base64.urlsafe_b64encode(
        hashlib.sha256(verifier.encode()).digest()).decode().rstrip("=")
    with open(ARQUIVO_PKCE, "w", encoding="utf-8") as f:
        f.write(verifier)
    return challenge


def _ler_pkce():
    if os.path.exists(ARQUIVO_PKCE):
        with open(ARQUIVO_PKCE, encoding="utf-8") as f:
            return f.read().strip()
    return None


def link_autorizacao(app_id=None, redirect_uri=None):
    """Link que você abre no navegador para autorizar o robô."""
    params = {"response_type": "code", "client_id": app_id or APP_ID,
              "redirect_uri": redirect_uri or REDIRECT_URI,
              "code_challenge": _novo_pkce(), "code_challenge_method": "S256"}
    return AUTH + "?" + urlencode(params)


def _extrair_code(texto):
    """Aceita o code puro OU o endereço inteiro colado (…?code=TG-…)."""
    texto = texto.strip()
    if "code=" in texto:
        return parse_qs(urlparse(texto).query).get("code", [texto])[0]
    return texto


def _salvar_token(dados, arquivo=ARQUIVO_TOKEN):
    dados["expira_em"] = time.time() + int(dados.get("expires_in", 21600)) - 300
    with open(arquivo, "w", encoding="utf-8") as f:
        json.dump(dados, f)
    return dados


def trocar_code_por_token(code, app_id=None, secret=None, redirect_uri=None,
                          arquivo=ARQUIVO_TOKEN):
    dados = {
        "grant_type": "authorization_code",
        "client_id": app_id or APP_ID,
        "client_secret": secret or SECRET_KEY,
        "code": _extrair_code(code),
        "redirect_uri": redirect_uri or REDIRECT_URI,
    }
    verifier = _ler_pkce()
    if verifier:
        dados["code_verifier"] = verifier
    r = requests.post(f"{API}/oauth/token", timeout=30, data=dados,
                      headers={"accept": "application/json",
                               "content-type": "application/x-www-form-urlencoded"})
    if r.status_code != 200:
        raise RuntimeError(f"Não consegui o token ({r.status_code}): {r.text}\n"
                           "Dica: o code só vale UMA vez e por poucos minutos. "
                           "Rode de novo a célula do link, abra o link NOVO "
                           "e cole o endereço logo em seguida.")
    return _salvar_token(r.json(), arquivo)


def _renovar_token(dados, app_id=None, secret=None, arquivo=ARQUIVO_TOKEN):
    r = requests.post(f"{API}/oauth/token", timeout=30, data={
        "grant_type": "refresh_token",
        "client_id": app_id or APP_ID,
        "client_secret": secret or SECRET_KEY,
        "refresh_token": dados.get("refresh_token", ""),
    }, headers={"accept": "application/json",
                "content-type": "application/x-www-form-urlencoded"})
    if r.status_code != 200:
        return None
    return _salvar_token(r.json(), arquivo)


def obter_token(app_id=None, secret=None, redirect_uri=None,
                arquivo=ARQUIVO_TOKEN, perguntar=True):
    """Usa o token salvo; renova se venceu; se não tiver, pede autorização."""
    if os.path.exists(arquivo):
        with open(arquivo, encoding="utf-8") as f:
            dados = json.load(f)
        if time.time() < dados.get("expira_em", 0):
            return dados["access_token"]
        novo = _renovar_token(dados, app_id, secret, arquivo)
        if novo:
            return novo["access_token"]
    if not perguntar:
        raise RuntimeError("Sem token válido. Rode a etapa de autorização.")
    print("\n1) Abra este link no navegador e clique em 'Permitir':\n")
    print("   " + link_autorizacao(app_id, redirect_uri))
    print("\n2) Você vai cair numa página (ex.: Google). Copie o endereço inteiro"
          " da barra do navegador (tem '?code=TG-...').")
    code = input("\n3) Cole aqui e aperte Enter: ")
    return trocar_code_por_token(code, app_id, secret, redirect_uri, arquivo)["access_token"]


# ============================================================
# 3) CONVERSA COM A API
# ============================================================
class ClienteML:
    def __init__(self, token):
        self.s = requests.Session()
        self.s.headers.update({"Authorization": f"Bearer {token}"})
        self.chamadas = 0

    def get(self, caminho, params=None, tentativas=3):
        for t in range(tentativas):
            time.sleep(PAUSA_ENTRE_CHAMADAS)
            self.chamadas += 1
            try:
                r = self.s.get(API + caminho, params=params, timeout=30)
            except requests.RequestException:
                time.sleep(2 * (t + 1))
                continue
            if r.status_code == 200:
                return r.json()
            if r.status_code == 429:          # muitas chamadas: espera
                time.sleep(5 * (t + 1))
                continue
            if r.status_code == 401:
                raise RuntimeError("Token inválido ou vencido. Apague o arquivo "
                                   f"{ARQUIVO_TOKEN} e autorize de novo.")
            return None                       # 403/404 etc.: sem dados
        return None


# ============================================================
# 4) COLETA
# ============================================================
def listar_categorias(cli, categorias=None, subcategorias=False, nicho=None):
    """Devolve [(id, nome_completo)]."""
    nicho = NICHO if nicho is None else nicho
    if categorias:
        saida = []
        for cid in categorias:
            info = cli.get(f"/categories/{cid}") or {}
            nome = info.get("name", cid)
            saida.append((cid, nome))
            if subcategorias:
                for f in info.get("children_categories", []):
                    saida.append((f["id"], f"{nome} > {f['name']}"))
        return saida

    tops = cli.get(f"/sites/{SITE}/categories") or []
    if not nicho:
        return [(c["id"], c["name"]) for c in tops]

    saida = []
    for c in tops:
        info = cli.get(f"/categories/{c['id']}") or {}
        top_ok = contem_palavra(c["name"], nicho)
        for f in info.get("children_categories", []):
            nome = f"{c['name']} > {f['name']}"
            if contem_palavra(nome, EXCLUIR_CATEGORIAS):
                continue
            if top_ok or contem_palavra(f["name"], nicho):
                saida.append((f["id"], nome))
    if not saida:
        print("Nenhuma subcategoria bateu com o NICHO; pesquisando as principais.")
        return [(c["id"], c["name"]) for c in tops]
    return saida


def _concorrencia(cli, product_id):
    dados = cli.get(f"/products/{product_id}/items", {"limit": 50})
    if not dados:
        return None, None
    res = dados.get("results", []) or []
    total = (dados.get("paging") or {}).get("total", len(res))
    vendedores = {x.get("seller_id") for x in res if x.get("seller_id")}
    n = max(total, len(vendedores)) if res else total
    precos = [x.get("price") for x in res if isinstance(x.get("price"), (int, float))]
    return n, (min(precos) if precos else None)


def detalhar(cli, tipo, id_):
    """Junta as informações de um produto (catálogo) ou anúncio."""
    d = {"id": id_, "tipo": tipo, "titulo": None, "preco": None, "vendedores": None,
         "loja_oficial_ganha": None, "frete_gratis": None, "full": None, "link": None}
    if tipo == "USER_PRODUCT" or str(id_).startswith("MLBU"):
        # Anúncio no formato novo do ML ("produto do vendedor").
        up = cli.get(f"/user-products/{id_}") or {}
        d["titulo"] = up.get("name") or up.get("title")
        d["link"] = f"https://www.mercadolivre.com.br/up/{id_}"
        if up.get("catalog_product_id"):
            d["vendedores"], d["preco"] = _concorrencia(cli, up["catalog_product_id"])
        if not d["titulo"]:
            d["titulo"] = f"{id_} (sem detalhes na API: abra o link)"
        return d
    if tipo == "PRODUCT":
        p = cli.get(f"/products/{id_}") or {}
        d["titulo"] = p.get("name")
        d["link"] = p.get("permalink") or f"https://www.mercadolivre.com.br/p/{id_}"
        bbw = p.get("buy_box_winner") or {}
        d["preco"] = bbw.get("price")
        if "official_store_id" in bbw:
            d["loja_oficial_ganha"] = bool(bbw.get("official_store_id"))
        ship = bbw.get("shipping") or {}
        d["frete_gratis"] = ship.get("free_shipping")
        d["full"] = ship.get("logistic_type") == "fulfillment" if ship else None
        n, menor = _concorrencia(cli, id_)
        d["vendedores"] = n
        if d["preco"] is None:
            d["preco"] = menor
    else:  # ITEM (anúncio comum) ou USER_PRODUCT
        it = cli.get(f"/items/{id_}") or {}
        d["titulo"] = it.get("title")
        d["preco"] = it.get("price")
        d["link"] = it.get("permalink")
        if "official_store_id" in it:
            d["loja_oficial_ganha"] = bool(it.get("official_store_id"))
        if not d["titulo"]:
            d["titulo"] = f"{id_} (sem detalhes na API: abra o link)"
            d["link"] = d["link"] or f"https://produto.mercadolivre.com.br/{id_[:3]}-{id_[3:]}"
        ship = it.get("shipping") or {}
        d["frete_gratis"] = ship.get("free_shipping")
        d["full"] = ship.get("logistic_type") == "fulfillment" if ship else None
        if it.get("catalog_product_id"):
            d["vendedores"], _ = _concorrencia(cli, it["catalog_product_id"])
    return d


def buscas_em_alta(cli, categoria=None):
    caminho = f"/trends/{SITE}" + (f"/{categoria}" if categoria else "")
    dados = cli.get(caminho) or []
    if isinstance(dados, dict):
        dados = dados.get("results", [])
    return [x.get("keyword") for x in dados if isinstance(x, dict) and x.get("keyword")]


# ============================================================
# 5) HISTÓRICO (para saber o que fica no topo semana após semana)
# ============================================================
CAMPOS_HIST = ["data", "id", "categoria", "posicao", "titulo", "preco"]


def carregar_historico(arquivo):
    if not os.path.exists(arquivo):
        return []
    with open(arquivo, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def salvar_historico(arquivo, linhas_antigas, novas):
    vistos = {(l["data"], l["id"]) for l in linhas_antigas}
    todas = list(linhas_antigas) + [l for l in novas if (l["data"], l["id"]) not in vistos]
    with open(arquivo, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS_HIST)
        w.writeheader()
        for l in todas:
            w.writerow({k: l.get(k, "") for k in CAMPOS_HIST})


# ============================================================
# 6) A "INTELIGÊNCIA": nota e veredito
# ============================================================
def avaliar(d, posicao, vezes_no_top, categoria_nome):
    preco, n = d.get("preco"), d.get("vendedores")
    bons, ruins = [], []

    demanda = (21 - posicao) / 20                      # 1º = 1.0 ... 20º = 0.05
    if posicao <= 5:
        bons.append(f"top {posicao} da categoria")

    if n is None:
        conc = 0.5
        ruins.append("concorrência não informada: confira no anúncio")
    else:
        conc = max(0.0, 1 - min(n, 50) / 50)
        if n <= 2:
            bons.append(f"poucos concorrentes no catálogo ({n})")
            ruins.append("confira a concorrência no site: a API conta só quem está no catálogo")
        elif n <= 8:
            bons.append(f"poucos concorrentes ({n})")
        elif n > MAX_VENDEDORES:
            ruins.append(f"muitos concorrentes ({n} vendedores)")

    if preco is None:
        tick = 0.5
    elif PRECO_MIN <= preco <= PRECO_MAX:
        tick = 1.0
        bons.append("preço na faixa ideal")
    elif 30 <= preco < PRECO_MIN or PRECO_MAX < preco <= 1000:
        tick = 0.5
        ruins.append(f"preço fora da faixa ideal (R$ {preco:.0f})")
    else:
        tick = 0.1
    if preco is not None and preco < 79:
        ruins.append("abaixo de R$ 79: o ML cobra custo fixo por unidade")

    oficial = 0.0 if d.get("loja_oficial_ganha") else 1.0
    if d.get("loja_oficial_ganha"):
        ruins.append("loja oficial da marca domina a venda")

    constancia = min(vezes_no_top, 4) / 4
    if vezes_no_top >= 3:
        bons.append(f"no topo em {vezes_no_top} pesquisas: tendência")
    elif vezes_no_top == 1:
        ruins.append("1ª vez no topo: confirme nas próximas semanas")

    nota = round(100 * (0.40 * demanda + 0.30 * conc + 0.15 * tick
                        + 0.05 * oficial + 0.10 * constancia))

    evitar = False
    if preco is not None and (preco < 30 or preco > 1000):
        evitar = True
        ruins.append("preço inviável para começar (taxa come a margem ou prende capital)")
    if n is not None and n > 40:
        evitar = True
    if d.get("loja_oficial_ganha") and (n or 0) > 15:
        evitar = True

    if evitar:
        veredito = "EVITAR"
    elif (nota >= 60 and n is not None and n <= MAX_VENDEDORES
          and not d.get("loja_oficial_ganha") and preco is not None and PRECO_MIN <= preco <= PRECO_MAX):
        veredito = "VENDER"
    else:
        veredito = "ANALISAR"

    inmetro = "Verificar" if contem_palavra(f"{d.get('titulo')} {categoria_nome}",
                                            PALAVRAS_INMETRO) else ""
    if inmetro:
        ruins.append("pode exigir certificação do Inmetro")
    motivo = "; ".join(bons + ruins) or "sem destaque"
    return nota, veredito, motivo, inmetro


# ============================================================
# 7) PLANILHA
# ============================================================
VERDE = PatternFill("solid", fgColor="C6EFCE")
AMARELO = PatternFill("solid", fgColor="FFEB9C")
VERMELHO = PatternFill("solid", fgColor="FFC7CE")
ENTRADA = PatternFill("solid", fgColor="FFF2CC")
CABEC = PatternFill("solid", fgColor="1F4E79")
P = "'Parâmetros'!"


def _cabecalho(ws, colunas, larguras):
    ws.append(colunas)
    for i, larg in enumerate(larguras, 1):
        c = ws.cell(row=1, column=i)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = CABEC
        c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = larg
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "D2"


def salvar_excel(linhas, tendencias, arquivo, simulado=False):
    wb = Workbook()

    # --- Parâmetros (usados pelas fórmulas)
    wp = wb.active
    wp.title = "Parâmetros"
    _cabecalho(wp, ["Parâmetro", "Valor", "Observação"], [42, 10, 70])
    wp.freeze_panes = None
    params = [
        ("Taxa do Mercado Livre (anúncio Clássico)", TAXA_ML, "0.0%", "Clássico 10 a 14%, Premium 15 a 19%, conforme a categoria. Confira no ML."),
        ("Custo fixo por unidade abaixo de R$ 79", CUSTO_FIXO_ABAIXO_79, '"R$" #,##0.00', "Valor aproximado. Confira a tabela atual do ML."),
        ("Frete pago por você acima de R$ 79", FRETE_ACIMA_79, '"R$" #,##0.00', "Estimativa. Depende do peso, do tamanho e da sua reputação."),
        ("Imposto sobre a venda", IMPOSTO, "0.0%", "Simples Nacional (comércio), faixa inicial. Se for MEI, use 0."),
        ("Margem mínima para vender", MARGEM_MINIMA, "0%", "Regra combinada: só vende se sobrar pelo menos isso do preço."),
    ]
    for nome, val, fmt, obs in params:
        wp.append([nome, val, obs])
        wp.cell(wp.max_row, 2).number_format = fmt
        wp.cell(wp.max_row, 2).fill = ENTRADA

    # --- Ranking
    ws = wb.create_sheet("Ranking", 0)
    cols = ["Veredito", "Nota", "Produto", "Categoria", "Posição", "Vezes no topo",
            "Preço (R$)", "Vendedores", "Loja oficial domina", "Full", "Inmetro",
            "Custo unitário (PREENCHA)", "Taxa ML (R$)", "Frete (R$)", "Imposto (R$)",
            "Lucro por unidade (R$)", "Margem", "Passa nos 20%?", "Por quê", "Link"]
    larg = [10, 6, 46, 30, 8, 8, 10, 10, 9, 6, 9, 12, 10, 9, 9, 11, 8, 9, 60, 8]
    _cabecalho(ws, cols, larg)
    sn = lambda v: "" if v is None else ("Sim" if v else "Não")
    ordem = {"VENDER": 0, "ANALISAR": 1, "EVITAR": 2}
    for l in sorted(linhas, key=lambda x: (ordem[x["veredito"]], -x["nota"])):
        ws.append([l["veredito"], l["nota"], l["titulo"] or l["id"], l["categoria"],
                   l["posicao"], l["vezes"], l["preco"], l["vendedores"],
                   sn(l["loja_oficial_ganha"]), sn(l["full"]), l["inmetro"],
                   None, None, None, None, None, None, None, l["motivo"], "abrir"])
        r = ws.max_row
        G, L = f"G{r}", f"L{r}"
        ws[f"M{r}"] = f'=IF({G}="","",{G}*{P}$B$2+IF({G}<79,{P}$B$3,0))'
        ws[f"N{r}"] = f'=IF({G}="","",IF({G}>=79,{P}$B$4,0))'
        ws[f"O{r}"] = f'=IF({G}="","",{G}*{P}$B$5)'
        ws[f"P{r}"] = f'=IF(OR({L}="",{G}=""),"",{G}-{L}-M{r}-N{r}-O{r})'
        ws[f"Q{r}"] = f'=IF(P{r}="","",P{r}/{G})'
        ws[f"R{r}"] = f'=IF(Q{r}="","",IF(Q{r}>={P}$B$6,"SIM","NÃO"))'
        ws.cell(r, 1).fill = {"VENDER": VERDE, "ANALISAR": AMARELO}.get(l["veredito"], VERMELHO)
        ws.cell(r, 1).font = Font(bold=True)
        ws.cell(r, 12).fill = ENTRADA
        for c in (7, 12, 13, 14, 15, 16):
            ws.cell(r, c).number_format = '#,##0.00'
        ws.cell(r, 17).number_format = "0%"
        ws.cell(r, 19).alignment = Alignment(wrap_text=True, vertical="top")
        if l.get("link"):
            ws.cell(r, 20).hyperlink = l["link"]
            ws.cell(r, 20).font = Font(color="0563C1", underline="single")
    ultima = max(ws.max_row, 2)
    ws.conditional_formatting.add(f"R2:R{ultima}", CellIsRule(operator="equal", formula=['"SIM"'], fill=VERDE))
    ws.conditional_formatting.add(f"R2:R{ultima}", CellIsRule(operator="equal", formula=['"NÃO"'], fill=VERMELHO))
    ws.auto_filter.ref = f"A1:T{ultima}"

    # --- Resumo por categoria
    wr = wb.create_sheet("Resumo por categoria", 1)
    _cabecalho(wr, ["Categoria", "Nota média", "VENDER", "ANALISAR", "EVITAR",
                    "Preço médio (R$)", "Vendedores (média)"], [46, 10, 9, 10, 9, 14, 14])
    wr.freeze_panes = "A2"
    grupos = {}
    for l in linhas:
        grupos.setdefault(l["categoria"], []).append(l)
    resumo = []
    for cat, ls in grupos.items():
        precos = [x["preco"] for x in ls if x["preco"]]
        vend = [x["vendedores"] for x in ls if x["vendedores"] is not None]
        resumo.append([cat, round(sum(x["nota"] for x in ls) / len(ls), 1),
                       sum(x["veredito"] == "VENDER" for x in ls),
                       sum(x["veredito"] == "ANALISAR" for x in ls),
                       sum(x["veredito"] == "EVITAR" for x in ls),
                       round(sum(precos) / len(precos), 2) if precos else None,
                       round(sum(vend) / len(vend), 1) if vend else None])
    for linha in sorted(resumo, key=lambda x: (-x[2], -x[1])):
        wr.append(linha)
        wr.cell(wr.max_row, 6).number_format = '#,##0.00'

    # --- Buscas em alta
    wt = wb.create_sheet("Buscas em alta", 2)
    _cabecalho(wt, ["Onde", "O que as pessoas estão buscando"], [46, 60])
    wt.freeze_panes = "A2"
    for onde, palavras in tendencias.items():
        for p in palavras:
            wt.append([onde, p])

    # --- Como usar
    wl = wb.create_sheet("Como usar", 0)
    wl.column_dimensions["A"].width = 115
    texto = [
        "COMO USAR ESTA PESQUISA",
        f"Gerada em {datetime.now():%d/%m/%Y %H:%M}" + ("  —  DADOS SIMULADOS (só exemplo, NÃO são reais)" if simulado else ""),
        "",
        "1. Abra a aba Ranking. Os produtos VENDER estão no topo: vendem muito, têm até "
        f"{MAX_VENDEDORES} concorrentes, preço entre R$ {PRECO_MIN:.0f} e R$ {PRECO_MAX:.0f} e nenhuma loja oficial dominando.",
        "2. Para cada candidato, pesquise o fornecedor (atacado no Brasil, com nota fiscal) e escreva o custo "
        "por unidade na coluna amarela 'Custo unitário'.",
        "3. A planilha calcula sozinha a taxa do ML, o frete, o imposto, o lucro e a margem. "
        "'Passa nos 20%?' = SIM quer dizer que a conta fecha.",
        "4. Confira a coluna Inmetro: produtos elétricos como tomadas, disjuntores, carregadores e extensões "
        "precisam de certificação. Compre só de fornecedor com produto certificado.",
        "5. Rode o robô toda semana. A coluna 'Vezes no topo' mostra em quantas pesquisas o produto apareceu: "
        "3 ou mais = tendência; 1 = pode ser moda passageira.",
        "6. Teste com 5 a 10 unidades de 2 ou 3 produtos. Depois de 30 a 45 dias, aumente o que deu lucro e girou.",
        "",
        "LIMITES DOS DADOS: a coluna Vendedores conta só quem vende pelo catálogo do ML; pode haver outros anúncios "
        "do mesmo produto. Antes de comprar, busque o produto no site e conte os concorrentes. "
        "Células vazias em Full e Loja oficial = a API não informou.",
        "",
        "Os valores de taxa, custo fixo, frete e imposto estão na aba Parâmetros (células amarelas). "
        "Confira os valores atuais no Mercado Livre e ajuste: todas as contas se atualizam.",
        "NOTA (0 a 100) = 40% demanda (posição) + 30% pouca concorrência + 15% preço + 10% constância no topo + 5% sem loja oficial.",
    ]
    for i, t in enumerate(texto, 1):
        c = wl.cell(i, 1, t)
        c.alignment = Alignment(wrap_text=True)
        if i == 1:
            c.font = Font(bold=True, size=14)
    wb.active = 1
    wb.save(arquivo)
    return arquivo


# ============================================================
# 8) EXECUÇÃO
# ============================================================
def rodar_pesquisa(cli, categorias=None, subcategorias=False, arquivo=None,
                   simulado=False, nicho=None, arquivo_historico=None):
    hoje = datetime.now().strftime("%Y-%m-%d")
    arquivo = arquivo or f"pesquisa_mercado_{hoje}.xlsx"
    arquivo_historico = arquivo_historico or (
        "historico_SIMULADO.csv" if simulado else ARQUIVO_HISTORICO)
    historico = carregar_historico(arquivo_historico)
    datas_por_id = {}
    for h in historico:
        datas_por_id.setdefault(h["id"], set()).add(h["data"])

    cats = listar_categorias(cli, categorias, subcategorias, nicho)
    print(f"Pesquisando {len(cats)} categorias… (pode levar alguns minutos)")

    linhas, tendencias, novas_hist, vistos = [], {}, [], set()
    geral = buscas_em_alta(cli)
    if geral:
        tendencias["Mercado Livre (geral)"] = geral

    for i, (cid, nome) in enumerate(cats, 1):
        print(f"  [{i}/{len(cats)}] {nome}")
        dados = cli.get(f"/highlights/{SITE}/category/{cid}")
        conteudo = (dados or {}).get("content", [])
        if not conteudo:
            print("      (sem ranking de mais vendidos nesta categoria)")
        for item in conteudo:
            if item["id"] in vistos:          # mesmo produto em outra categoria
                continue
            vistos.add(item["id"])
            d = detalhar(cli, item.get("type", "ITEM"), item["id"])
            pos = item.get("position", 20)
            vezes = len(datas_por_id.get(d["id"], set()) | {hoje})
            nota, veredito, motivo, inmetro = avaliar(d, pos, vezes, nome)
            d.update(categoria=nome, posicao=pos, vezes=vezes, nota=nota,
                     veredito=veredito, motivo=motivo, inmetro=inmetro)
            linhas.append(d)
            novas_hist.append({"data": hoje, "id": d["id"], "categoria": nome,
                               "posicao": pos, "titulo": d["titulo"] or "",
                               "preco": d["preco"] if d["preco"] is not None else ""})
        alta = buscas_em_alta(cli, cid)
        if alta:
            tendencias[nome] = alta

    if not linhas:
        print("\nNenhum dado coletado. Verifique o token e as categorias.")
        return None
    salvar_historico(arquivo_historico, historico, novas_hist)
    salvar_excel(linhas, tendencias, arquivo, simulado)
    bons = sorted([l for l in linhas if l["veredito"] == "VENDER"], key=lambda x: -x["nota"])
    print(f"\nPronto! {len(linhas)} produtos analisados ({cli.chamadas} consultas).")
    print(f"VENDER: {len(bons)} | ANALISAR: {sum(l['veredito']=='ANALISAR' for l in linhas)}"
          f" | EVITAR: {sum(l['veredito']=='EVITAR' for l in linhas)}")
    for l in bons[:10]:
        print(f"  {l['nota']:>3}  {(l['titulo'] or '')[:60]}  (R$ {l['preco'] or 0:.2f})")
    print(f"\nPlanilha: {os.path.abspath(arquivo)}")
    print(f"Histórico: {os.path.abspath(arquivo_historico)} (guarde este arquivo para a próxima semana)")
    return arquivo


# ------------------------------------------------------------
# Modo simulado: imita a API com dados FALSOS, só para testar
# ------------------------------------------------------------
class ClienteSimulado(ClienteML):
    def __init__(self, semente=7):
        self.chamadas = 0
        self.rng = random.Random(semente)
        self.tops = [("MLB263532", "Ferramentas"), ("MLB1500", "Construção"),
                     ("MLB1574", "Casa, Móveis e Decoração"), ("MLB1051", "Celulares e Telefones")]
        self.filhos = {"MLB263532": ["Ferramentas Elétricas", "Ferramentas Manuais"],
                       "MLB1500": ["Materiais Elétricos", "Pisos e Revestimentos"],
                       "MLB1574": ["Iluminação", "Cozinha"],
                       "MLB1051": ["Carregadores", "Capinhas"]}
        self.nomes = ["Tomada inteligente Wi-Fi", "Kit de ferramentas 129 peças", "Disjuntor DIN 20A",
                      "Fita LED 5 m", "Carregador turbo USB-C", "Furadeira de impacto 650 W",
                      "Kit 10 lâmpadas LED 9 W", "Extensão elétrica 5 m", "Alicate amperímetro digital",
                      "Luminária de emergência", "Multímetro digital", "Sensor de presença"]

    def get(self, caminho, params=None, tentativas=3):
        self.chamadas += 1
        r = self.rng
        if caminho == f"/sites/{SITE}/categories":
            return [{"id": i, "name": n} for i, n in self.tops]
        if caminho.startswith("/categories/"):
            cid = caminho.split("/")[-1]
            return {"id": cid, "name": dict(self.tops).get(cid, cid),
                    "children_categories": [{"id": f"{cid}-{k}", "name": n}
                                            for k, n in enumerate(self.filhos.get(cid, []))]}
        if caminho.startswith(f"/highlights/{SITE}/category/"):
            cid = caminho.split("/")[-1]
            return {"content": [{"id": f"{cid}P{k}", "position": k,
                                 "type": r.choice(["PRODUCT", "PRODUCT", "ITEM"])} for k in range(1, 21)]}
        if caminho.startswith("/products/") and caminho.endswith("/items"):
            n = r.choice([3, 5, 8, 12, 18, 25, 35, 48])
            return {"paging": {"total": n},
                    "results": [{"seller_id": s, "price": round(r.uniform(40, 400), 2),
                                 "official_store_id": None} for s in range(min(n, 50))]}
        if caminho.startswith("/products/"):
            pid = caminho.split("/")[-1]
            return {"id": pid, "name": r.choice(self.nomes) + " (exemplo)",
                    "buy_box_winner": {"price": r.choice([24.9, 59.9, 89.9, 129.0, 189.0, 249.0, 459.0]),
                                       "official_store_id": r.choice([None, None, None, 123]),
                                       "shipping": {"free_shipping": True,
                                                    "logistic_type": r.choice(["fulfillment", "cross_docking"])}}}
        if caminho.startswith("/items/"):
            iid = caminho.split("/")[-1]
            return {"id": iid, "title": r.choice(self.nomes) + " (exemplo)",
                    "price": round(r.uniform(30, 350), 2), "official_store_id": None,
                    "permalink": "https://www.mercadolivre.com.br/",
                    "shipping": {"free_shipping": True, "logistic_type": "fulfillment"},
                    "catalog_product_id": None}
        if caminho.startswith(f"/trends/{SITE}"):
            return [{"keyword": k, "url": ""} for k in
                    r.sample(["tomada inteligente", "fita led", "carregador veicular", "kit ferramentas",
                              "lâmpada smart", "sensor de presença", "painel solar", "extensão"], 5)]
        return None


def main():
    if "--simular" in sys.argv:
        print("MODO SIMULADO: dados FALSOS só para você ver a planilha.\n")
        rodar_pesquisa(ClienteSimulado(), arquivo="pesquisa_EXEMPLO_simulada.xlsx", simulado=True)
        return
    if not APP_ID or not SECRET_KEY:
        print("Preencha APP_ID e SECRET_KEY no começo do arquivo (veja o LEIA-ME).")
        print("Para só ver um exemplo da planilha: python pesquisa_mercado.py --simular")
        return
    token = obter_token()
    rodar_pesquisa(ClienteML(token), CATEGORIAS, INCLUIR_SUBCATEGORIAS)


if __name__ == "__main__":
    main()
