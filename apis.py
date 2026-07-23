# =============================================================================
# apis.py — cliente leve para B3 Calculator e FI Analytics
# -----------------------------------------------------------------------------
# Self-contained: usa apenas a biblioteca padrão (urllib), sem httpx e sem
# depender de outros repositórios. Pensado para o add-in do Excel: precifica
# via API (B3 → FI Analytics).
#
# Segredos e proxy: lidos de variáveis de ambiente (banco), com fallback no
# .env para desenvolvimento local. Ver ENV_* abaixo.
#
# Estratégia de leveza:
#   - token B3 obtido uma vez e reaproveitado (renova em 401);
#   - cache de resposta por (origem, ticker, dataIso, taxa/pu) — inclusive
#     resultados None — para não refazer chamadas a cada recálculo do Excel;
#   - timeout curto por chamada, para não travar a planilha.
# =============================================================================

import os
import re
import json
import time
import atexit
import base64
import tempfile
import http.client
import urllib.request
import urllib.error
from urllib.parse import urlparse, unquote

from config import ENV_PATH

# Nomes das variáveis de ambiente (definidas no banco). O add-in lê os segredos
# daqui em produção; o .env serve só como fallback de desenvolvimento local.
ENV_TOKEN_B3 = "token_calc_b3"
ENV_KEY_FI   = "token_fianalytics"
ENV_USER_FI  = "user_fianalytics"   # e-mail p/ getuserbonds do bondbuilder (fallback: FIANALYTICS_USER)

# Proxy — env vars com a URL completa, incluindo usuário/senha, no formato
# http://USUARIO:SENHA@HOST:PORTA. Ausentes (dev) → conexão direta.
ENV_PROXY_HTTP  = "proxy_http"
ENV_PROXY_HTTPS = "proxy_https"

B3_BASE      = "https://api.calculadorarendafixa.com.br"
FI_BASE      = "https://endpoint.fi-analytics.com.br"
FI_DEB_PATH  = "/deb/debenturecalculator"
FI_CR_PATH   = "/cr/cricracalculator"   # CRI/CRA (mesmos campos da resposta)
FI_BB_PATH       = "/bb/bondbuildercalculator"             # fallback: papéis fora de /deb e /cr
FI_BB_BONDS_PATH = "/bb/bondbuildercalculator/getuserbonds"
TIMEOUT_SEG  = 6  # timeout curto por chamada para não travar o Excel

# Cache em disco (persiste entre reaberturas do Excel): só resultados VÁLIDOS
# (nunca None/erro), com validade de CACHE_TTL_SEG. Fora do TTL, refaz a chamada.
CACHE_TTL_SEG = 600  # 10 min
_CACHE_FILE = os.path.join(tempfile.gettempdir(), "calcrf_cache.json")

# Gravação do cache em disco é "debounced": num recálculo em massa (F9), cada
# célula nova geraria uma reescrita do arquivo inteiro. Em vez disso mantemos o
# registro em MEMÓRIA e gravamos no máximo a cada _MIN_FLUSH_SEG; o atexit
# garante o flush final ao fechar o Excel. O disco só serve p/ aquecer a próxima
# sessão — perder alguns segundos do fim de um burst é irrelevante.
_MIN_FLUSH_SEG = 3.0

_CACHE_AUSENTE = object()


def _Norm(valor):
    """Remove ruído de float (ex: 6.461800000000001 → 6.4618) em taxas e PUs."""
    return float(f"{float(valor):.10g}")


def CarregarCredenciais():
    """Lê chaves do .env (formato CHAVE=valor). Falha silenciosa se ausente."""
    creds = {}
    try:
        with open(ENV_PATH, encoding="utf-8") as arquivo:
            for linha in arquivo:
                linha = linha.strip()
                if "=" in linha and not linha.startswith("#"):
                    chave, valor = linha.split("=", 1)
                    creds[chave.strip()] = valor.strip()
    except OSError:
        pass
    return creds


CREDENCIAIS = CarregarCredenciais()


def _Cred(nomeEnv, chaveArquivo):
    """Resolve um segredo: variável de ambiente primeiro (banco), .env como fallback (dev)."""
    return os.getenv(nomeEnv) or CREDENCIAIS.get(chaveArquivo)


def _ConstruirOpener():
    """Opener urllib com o proxy do banco (env vars). Sem env vars → conexão direta.

    O urllib aceita proxy autenticado embutido na URL (http://user:senha@host:porta)
    e, em HTTPS, move o Proxy-Authorization para o tunnel CONNECT automaticamente.
    """
    proxies = {}
    http  = os.getenv(ENV_PROXY_HTTP)
    https = os.getenv(ENV_PROXY_HTTPS)
    if http:
        proxies["http"] = http
    if https:
        proxies["https"] = https
    return urllib.request.build_opener(urllib.request.ProxyHandler(proxies))


# Opener montado uma vez no import (as env vars do banco já existem no processo).
_opener = _ConstruirOpener()


_tokenB3 = None
_cacheRespostas: dict = {}
_fonteTicker: dict = {}   # ticker(upper) -> "b3"|"fi"|"bb": qual fonte respondeu (memo)
_userBonds = None         # cache 1x/processo do getuserbonds: lista de {bond_name,_id} ou None
_userBondsBuscado = False

# Espelho em memória do cache de disco: chaveStr -> {"k":list, "v":valor, "t":ts}.
# Carregado 1× no import; sincronizado com o arquivo só nos flushes (não a cada set).
_discoRegistro: dict = {}
_discoSujo = False
_discoUltimoFlush = 0.0


# ─── cache em disco (TTL) ────────────────────────────────────────────────────

def _LerArquivoCache() -> dict:
    try:
        with open(_CACHE_FILE, encoding="utf-8") as f:
            dados = json.load(f)
        return dados if isinstance(dados, dict) else {}
    except Exception:
        return {}


def _CarregarCacheDisco():
    """Carrega, no import, as entradas do arquivo ainda dentro do TTL — tanto no
    cache em memória (respostas) quanto no espelho de disco (p/ os flushes)."""
    agora = time.time()
    for chaveStr, entrada in _LerArquivoCache().items():
        try:
            if isinstance(entrada, dict) and agora - entrada["t"] < CACHE_TTL_SEG:
                _cacheRespostas[tuple(entrada["k"])] = entrada["v"]
                _discoRegistro[chaveStr] = entrada
        except Exception:
            pass


def _FlushDisco(forcado=False):
    """Grava o espelho no arquivo, no máximo a cada _MIN_FLUSH_SEG (ou já, se
    `forcado`). Antes de gravar, mescla o que outros processos tenham escrito e
    poda as entradas expiradas. Escrita atômica; falha de I/O é ignorada."""
    global _discoSujo, _discoUltimoFlush
    if not _discoSujo:
        return
    agora = time.time()
    if not forcado and agora - _discoUltimoFlush < _MIN_FLUSH_SEG:
        return
    try:
        # Mescla entradas de outros processos (outra instância do Excel usa o mesmo
        # arquivo em %TEMP%); em conflito, vence o timestamp mais novo.
        for chaveStr, entrada in _LerArquivoCache().items():
            if isinstance(entrada, dict):
                atual = _discoRegistro.get(chaveStr)
                if atual is None or entrada.get("t", 0) > atual.get("t", 0):
                    _discoRegistro[chaveStr] = entrada
        vivos = {k: v for k, v in _discoRegistro.items()
                 if isinstance(v, dict) and agora - v.get("t", 0) < CACHE_TTL_SEG}
        _discoRegistro.clear()
        _discoRegistro.update(vivos)
        tmp = _CACHE_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(_discoRegistro, f)
        os.replace(tmp, _CACHE_FILE)
        _discoSujo = False
        _discoUltimoFlush = agora
    except Exception:
        pass


atexit.register(_FlushDisco, forcado=True)   # garante o flush final ao fechar o Excel


def _CacheGet(chave):
    return _cacheRespostas.get(chave, _CACHE_AUSENTE)


def _CacheSet(chave, valor):
    global _discoSujo
    _cacheRespostas[chave] = valor
    if valor is not None:            # só persiste resultado válido (nunca None/erro)
        _discoRegistro["|".join(map(str, chave))] = {
            "k": list(chave), "v": valor, "t": time.time(),
        }
        _discoSujo = True
        _FlushDisco()


def _Abrir(req, base=None):
    """Abre a requisição (timeout padrão). `base` mantido só p/ compat. da chamada."""
    return _opener.open(req, timeout=TIMEOUT_SEG)


def LimparCache():
    """Esvazia o cache (memória+disco), o token, o memo e os bonds."""
    global _tokenB3, _userBonds, _userBondsBuscado, _discoSujo
    _cacheRespostas.clear()
    _fonteTicker.clear()
    _discoRegistro.clear()
    _discoSujo = False
    _tokenB3 = None
    _userBonds = None
    _userBondsBuscado = False
    try:
        os.remove(_CACHE_FILE)
    except OSError:
        pass


# -----------------------------------------------------------------------------
# B3 Calculator
# -----------------------------------------------------------------------------

def _LoginB3():
    """Obtém o token de sessão via POST /login. Retorna o token ou None."""
    token = _Cred(ENV_TOKEN_B3, "B3_CALC_TOKEN")
    if not token:
        return None
    try:
        corpo = json.dumps({"token": token}).encode()
        req = urllib.request.Request(
            f"{B3_BASE}/login", data=corpo,
            headers={"Content-Type": "application/json"}, method="POST",
        )
        with _Abrir(req, B3_BASE) as resp:
            dados = json.loads(resp.read().decode())
        return dados.get("Authorization")
    except Exception:
        return None


def _ObterTokenB3():
    global _tokenB3
    if _tokenB3 is None:
        _tokenB3 = _LoginB3()
    return _tokenB3


def _GetB3(url):
    """GET autenticado. Em 401 renova o token e tenta uma vez mais. dict ou None."""
    global _tokenB3
    for tentativa in range(2):
        token = _ObterTokenB3()
        if token is None:
            return None
        try:
            req = urllib.request.Request(url, headers={"Authorization": token}, method="GET")
            with _Abrir(req, B3_BASE) as resp:
                return json.loads(resp.read().decode())
        except urllib.error.HTTPError as erro:
            if erro.code == 401 and tentativa == 0:
                _tokenB3 = None
                continue
            return None
        except Exception:
            return None
    return None


_RE_NTNB = re.compile(r"^NTN-?B\s*(\d{2})$")
_RE_NTNF = re.compile(r"^NTN-?F\s*(\d{2})$")


def _NormalizarTicker(ticker):
    """Converte nome amigável de título público no código CETIP da B3.
    NTN-B (Tesouro IPCA+): vence dia 15; ano par→ago(08), ímpar→mai(05) → 760199{ano}{mmdd}.
    NTN-F (Tesouro Prefixado): vence 01/01 → 950199{ano}0101.
    Aceita 'NTNB35', 'NTN-B 35', 'NTN-B35'. Demais tickers passam inalterados (ex.: código
    cetip já pronto, debênture, CRA)."""
    t = str(ticker).upper().strip()
    m = _RE_NTNB.match(t)
    if m:
        ano = 2000 + int(m.group(1))
        return f"760199{ano}{'0815' if ano % 2 == 0 else '0515'}"
    m = _RE_NTNF.match(t)
    if m:
        return f"950199{2000 + int(m.group(1))}0101"
    return t


def PrecoB3(ticker, dataIso, taxa):
    """calcPU → {'pu': float, 'duration': anos|None} ou None. taxa em % a.a."""
    ticker = _NormalizarTicker(ticker)
    taxa = _Norm(taxa)
    chaveCache = ("b3pu", ticker, dataIso, taxa)
    cacheado = _CacheGet(chaveCache)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    dados = _GetB3(f"{B3_BASE}/calcPU/{ticker}/{dataIso}/{taxa}")
    resultado = None
    if isinstance(dados, dict):
        pu    = dados.get("PU")
        dur   = dados.get("duration")
        pupar = dados.get("PUPar")
        vna   = dados.get("VNA")
        if pu is not None and float(pu) > 0:
            resultado = {
                "pu": float(pu),
                "duration": float(dur) if dur is not None else None,
                "pupar": float(pupar) if pupar is not None else None,
                "vna": float(vna) if vna is not None else None,
            }
    _CacheSet(chaveCache, resultado)
    return resultado


def TaxaB3(ticker, dataIso, pu):
    """calcYield → yield em % a.a. ou None."""
    ticker = _NormalizarTicker(ticker)
    pu = _Norm(pu)
    chaveCache = ("b3yield", ticker, dataIso, pu)
    cacheado = _CacheGet(chaveCache)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    dados = _GetB3(f"{B3_BASE}/calcYield/{ticker}/{dataIso}/{pu}")
    resultado = None
    if isinstance(dados, dict):
        taxa = dados.get("yield")
        if taxa is not None and float(taxa) > 0:
            resultado = float(taxa)
    _CacheSet(chaveCache, resultado)
    return resultado


def FatorDi(dataInicioIso, dataFimIso, percentual=100.0):
    """Fator de CDI acumulado entre duas datas (endpoint público B3 /di/calculo).

    Retorna o `fator` (ex.: 1.10775126 para 100% do CDI em 2024) como float, ou None.
    - dataInicioIso/dataFimIso em 'YYYY-MM-DD';
    - `percentual` = % do CDI (100 = 100% do CDI; 110 = 110% do CDI).
    Endpoint SEM autenticação (não usa token B3), mas."""
    percentual = _Norm(percentual)
    chaveCache = ("b3di", dataInicioIso, dataFimIso, percentual)
    cacheado = _CacheGet(chaveCache)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    resultado = None
    try:
        url = (f"{B3_BASE}/di/calculo?dataInicio={dataInicioIso}&dataFim={dataFimIso}"
               f"&percentual={percentual}&valor=1")
        req = urllib.request.Request(url, method="GET")
        with _Abrir(req, B3_BASE) as resp:
            dados = json.loads(resp.read().decode())
        fator = dados.get("fator") if isinstance(dados, dict) else None
        if fator is not None and float(fator) > 0:
            resultado = float(fator)
    except Exception:
        resultado = None
    _CacheSet(chaveCache, resultado)
    return resultado


# -----------------------------------------------------------------------------
# FI Analytics
# -----------------------------------------------------------------------------

def _PostFi(corpo, path=FI_DEB_PATH):
    """POST num endpoint FI. Resposta é double-encoded. dict ou None."""
    chave = _Cred(ENV_KEY_FI, "FIANALYTICS_API_KEY")
    if not chave:
        return None
    try:
        req = urllib.request.Request(
            f"{FI_BASE}{path}",
            data=json.dumps(corpo).encode(),
            headers={"Content-Type": "application/json; charset=utf-8", "x-api-key": chave},
            method="POST",
        )
        with _Abrir(req, FI_BASE) as resp:
            externo = json.loads(resp.read().decode())
        return json.loads(externo) if isinstance(externo, str) else externo
    except Exception:
        return None


def _PostFiAuto(corpo):
    """Tenta debênture; se não for (ex.: CRI/CRA), tenta o endpoint de CRI/CRA.
    Os dois endpoints devolvem os mesmos campos (m2m, m2mRate, ...)."""
    for path in (FI_DEB_PATH, FI_CR_PATH):
        dados = _PostFi(corpo, path)
        if isinstance(dados, dict) and (
            dados.get("m2m") is not None or dados.get("m2mRate") is not None
        ):
            return dados
    return None


def PrecoFi(ticker, dataIso, taxa):
    """Modo rate → {'pu': m2m, 'duration': maculayDuration anos|None} ou None."""
    taxa = _Norm(taxa)
    chaveCache = ("fipu", ticker, dataIso, taxa)
    cacheado = _CacheGet(chaveCache)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    dados = _PostFiAuto({"ticker": ticker, "date": dataIso, "rate": float(taxa)})
    resultado = None
    if isinstance(dados, dict):
        pu    = dados.get("m2m")
        dur   = dados.get("maculayDuration")
        pupar = dados.get("currentNotionalPlusAccruedInterest")
        vna   = dados.get("adjustedFaceValue")
        if pu is not None and float(pu) > 0:
            resultado = {
                "pu": float(pu),
                "duration": float(dur) if dur is not None else None,
                "pupar": float(pupar) if pupar is not None else None,
                "vna": float(vna) if vna is not None else None,
            }
    _CacheSet(chaveCache, resultado)
    return resultado


def TaxaFi(ticker, dataIso, pu):
    """Modo pu → yield em % a.a. (m2mRate × 100) ou None."""
    pu = _Norm(pu)
    chaveCache = ("fiyield", ticker, dataIso, pu)
    cacheado = _CacheGet(chaveCache)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    dados = _PostFiAuto({"ticker": ticker, "date": dataIso, "pu": float(pu)})
    resultado = None
    if isinstance(dados, dict):
        taxa = dados.get("m2mRate")
        if taxa is not None and float(taxa) > 0:
            resultado = float(taxa) * 100
    _CacheSet(chaveCache, resultado)
    return resultado


# -----------------------------------------------------------------------------
# FI Analytics — bondbuilder (fallback p/ papéis fora de /deb e /cr)
# -----------------------------------------------------------------------------
# Mesmo motor da FI, mas o papel é endereçado por doc_id (de getuserbonds), não
# por ticker. É SIMÉTRICO como o /deb: manda `rate` → volta `m2m` (PU) + duration;
# manda `pu` → volta `m2mRate`. Exige o e-mail do usuário (ENV_USER_FI) e uma
# chamada getuserbonds (cacheada 1x/processo). Só dispara depois de B3 e FI falharem.

def _PostFiRaw(path, corpo):
    """POST num endpoint FI via http.client (NÃO via urllib).

    Motivo: o WAF do /bb exige o header 'x-api-key' em MINÚSCULO; o urllib.request
    title-caseia todo header ('X-Api-Key') e o gateway devolve 502. O http.client
    preserva a caixa. Suporta o proxy do banco (CONNECT tunnel c/ Proxy-Authorization).
    Retorna dict/list (double-encoded resolvido) ou None."""
    chave = _Cred(ENV_KEY_FI, "FIANALYTICS_API_KEY")
    if not chave:
        return None
    corpoBytes = json.dumps(corpo).encode()
    headers = {"content-type": "application/json; charset=utf-8", "x-api-key": chave}
    host = urlparse(FI_BASE).hostname
    proxy = os.getenv(ENV_PROXY_HTTPS)
    conn = None
    try:
        if proxy:
            p = urlparse(proxy)
            conn = http.client.HTTPSConnection(p.hostname, p.port or 8080, timeout=TIMEOUT_SEG)
            tunel = {}
            if p.username:
                cred = f"{unquote(p.username)}:{unquote(p.password or '')}"
                tunel["Proxy-Authorization"] = "Basic " + base64.b64encode(cred.encode()).decode()
            conn.set_tunnel(host, 443, tunel)
        else:
            conn = http.client.HTTPSConnection(host, 443, timeout=TIMEOUT_SEG)
        conn.request("POST", path, body=corpoBytes, headers=headers)
        resp = conn.getresponse()
        raw = resp.read().decode()
        status = resp.status
    except Exception:
        return None
    finally:
        if conn is not None:
            conn.close()
    if status >= 400:
        return None
    try:
        externo = json.loads(raw)
        return json.loads(externo) if isinstance(externo, str) else externo
    except Exception:
        return None


def _ObterBonds():
    """getuserbonds (cacheado 1x/processo). Lista de {bond_name,_id} ou None."""
    global _userBonds, _userBondsBuscado
    if _userBondsBuscado:
        return _userBonds
    _userBondsBuscado = True
    email = _Cred(ENV_USER_FI, "FIANALYTICS_USER")
    if email:
        dados = _PostFiRaw(FI_BB_BONDS_PATH, {"user_email": email, "get_company_bonds": "1"})
        if isinstance(dados, dict):
            dados = dados.get("data") or dados.get("bonds")
        if isinstance(dados, list):
            _userBonds = dados
    return _userBonds


def _AcharDocId(ticker):
    """doc_id do bond cujo bond_name == ticker (case-insensitive), ou None."""
    bonds = _ObterBonds()
    if not bonds:
        return None
    alvo = str(ticker).upper().strip()
    for b in bonds:
        if str(b.get("bond_name", "")).upper().strip() == alvo:
            return b.get("_id")
    return None


def _BbOk(dados):
    """Resposta do bondbuilder sem erro (statusCode 200 ou ausente)."""
    return isinstance(dados, dict) and dados.get("statusCode") in (200, None)


def PrecoBb(ticker, dataIso, taxa):
    """bondbuilder modo rate → {'pu': m2m, 'duration': maculayDuration, 'pupar'} ou None."""
    taxa = _Norm(taxa)
    chaveCache = ("bbpu", ticker, dataIso, taxa)
    cacheado = _CacheGet(chaveCache)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    resultado = None
    docId = _AcharDocId(ticker)
    if docId:
        dados = _PostFiRaw(FI_BB_PATH, {"doc_id": docId, "date": dataIso, "rate": float(taxa)})
        if _BbOk(dados):
            pu    = dados.get("m2m")
            dur   = dados.get("maculayDuration")
            pupar = dados.get("currentNotionalPlusAccruedInterest")
            vna   = dados.get("adjustedFaceValue")
            if pu is not None and float(pu) > 0:
                resultado = {
                    "pu": float(pu),
                    "duration": float(dur) if dur is not None else None,
                    "pupar": float(pupar) if pupar is not None else None,
                    "vna": float(vna) if vna is not None else None,
                }
    _CacheSet(chaveCache, resultado)
    return resultado


def TaxaBb(ticker, dataIso, pu):
    """bondbuilder modo pu → yield em % a.a. (m2mRate × 100) ou None."""
    pu = _Norm(pu)
    chaveCache = ("bbyield", ticker, dataIso, pu)
    cacheado = _CacheGet(chaveCache)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    resultado = None
    docId = _AcharDocId(ticker)
    if docId:
        dados = _PostFiRaw(FI_BB_PATH, {"doc_id": docId, "date": dataIso, "pu": float(pu)})
        if _BbOk(dados):
            taxa = dados.get("m2mRate")
            if taxa is not None and float(taxa) > 0:
                resultado = float(taxa) * 100
    _CacheSet(chaveCache, resultado)
    return resultado


# ─── roteamento B3 → FI → bondbuilder com memo de fonte por ticker ───────────
# Tenta as fontes em ordem; a que respondeu (só em sucesso) é memorizada e passa
# a ser tentada 1º nas próximas — mas o fallback é SEMPRE mantido (o memo só
# reordena, nunca bloqueia). bondbuilder fica por último (custa 1 getuserbonds).

def _OrdemFontes(tk, todas):
    memo = _fonteTicker.get(tk)
    return [memo] + [s for s in todas if s != memo] if memo in todas else list(todas)


def Preco(ticker, dataIso, taxa):
    """PU via B3→FI→bondbuilder com memo de fonte. dict {pu,duration,pupar} ou None."""
    tk = str(ticker).upper().strip()
    fontes = {"b3": PrecoB3, "fi": PrecoFi, "bb": PrecoBb}
    for src in _OrdemFontes(tk, ["b3", "fi", "bb"]):
        r = fontes[src](tk, dataIso, taxa)
        if r:
            _fonteTicker[tk] = src
            return r
    return None


def TaxaOp(ticker, dataIso, pu):
    """Taxa (% a.a.) via B3→FI→bondbuilder com memo de fonte. float ou None."""
    tk = str(ticker).upper().strip()
    fontes = {"b3": TaxaB3, "fi": TaxaFi, "bb": TaxaBb}
    for src in _OrdemFontes(tk, ["b3", "fi", "bb"]):
        r = fontes[src](tk, dataIso, pu)
        if r is not None:
            _fonteTicker[tk] = src
            return r
    return None


# =============================================================================
# DETALHES COMPLETOS DO PAPEL — pupar, VNA, fluxo, datas, gross up, etc.
# -----------------------------------------------------------------------------
# A FI Analytics devolve TUDO num call (a fonte mais rica); a B3 getBondDetails
# traz as datas/taxa de emissão em unidade nativa. As funções abaixo expõem esses
# campos de forma normalizada (source-agnostic) para as UDFs do add-in.
# =============================================================================

def BondDetailsB3(ticker):
    """getBondDetails da B3 (estático, sem data): dict cru ou None.
    Campos: expiredate, issuedate, startingdate, vne, yield, anniversaryday, events..."""
    tk = _NormalizarTicker(ticker)
    chave = ("b3bond", tk)
    c = _CacheGet(chave)
    if c is not _CACHE_AUSENTE:
        return c
    dados = _GetB3(f"{B3_BASE}/getBondDetails/{tk}")
    res = dados if isinstance(dados, dict) and dados.get("codbond") else None
    _CacheSet(chave, res)
    return res


def FluxoCadastrado(ticker):
    """Agenda de eventos cadastrada (getBondDetails.events), AGREGADA por data: lista ordenada
    [{data(iso), tipo, amort(%), incorp(%)}], ou None se não houver cadastro na B3.

    - UMA linha por data: a B3 manda os eventos SEPARADOS, então uma data de amortização vem 2x
      (o 'A' e o 'J' do cupom). Aqui vira uma linha só (ex.: FGEN13 15/06/2027 = 'A' 8,025 + 'J' 0).
    - TODA data da agenda entra, inclusive as de cupom puro (saem com amort=0 e incorp=0) — é a
      agenda de eventos do papel, não só de amort/incorp. Num papel bullet com cupom (ex.: ISAEC2,
      method=IPCA) as datas de cupom são justamente o que interessa.
    - `tipo`: composto pelo que existe na data — 'J' (juros), 'J+A' (juros e amortização),
      'J+I' (juros e incorporação).
    Regras (validação de fluxos):
      - 'A' e 'V' -> %amortização (yield).
      - 'J': só quando method == 'IPCA-I' e yield>0 é incorporação (yield=%incorp). Nos demais
        estilos (IPCA simples, DI-PERC, DI-SPREAD, PRE) o yield do 'J' é a TAXA do cupom PAGO,
        NÃO incorporação -> não entra.
    Datas vêm no dia cru (ex.: dia 15) — o ajuste p/ dia útil é feito na UDF."""
    b = BondDetailsB3(ticker)
    if not isinstance(b, dict) or not b.get("events"):
        return None
    ipcaI = (b.get("method") == "IPCA-I")
    porData = {}
    for e in b["events"]:
        d = e.get("date")
        if not d:
            continue
        t = e.get("eventType")
        y = e.get("yield") or 0.0
        acc = porData.setdefault(d, [0.0, 0.0, False])   # [amort, incorp, temJuros]
        if t in ("A", "V"):
            acc[0] += y
        elif t == "J":
            acc[2] = True
            if ipcaI and y > 0:
                acc[1] += y
            # 'J' cupom (não-IPCA-I ou yield 0) não soma %: é juros pago, não incorporação
    if porData:
        _CompletarResiduo(porData, b.get("expiredate"))
    return [{"data": d, "tipo": _TipoEvento(j, a, i), "amort": a, "incorp": i}
            for d, (a, i, j) in sorted(porData.items())]


def _CompletarResiduo(porData, vencimento):
    """Fecha o principal na linha do VENCIMENTO quando os eventos 'A' cadastrados somam < 100%.

    A B3 não cadastra o principal residual quitado no vencimento (FGEN13: os 'A' somam 91,975 e o
    vencimento vem só com o 'J' → faltam 8,025%). Aqui o vencimento recebe 100 − Σ%amort. Valor
    DERIVADO, não cadastrado. Se os 'A' já somam 100 (ex.: ISAEC2, bullet), nada muda."""
    venc = vencimento if vencimento in porData else max(porData)
    residuo = round(100.0 - sum(v[0] for v in porData.values()), 6)
    if residuo > 0:
        porData[venc][0] = round(porData[venc][0] + residuo, 6)


def _TipoEvento(temJuros, amort, incorp):
    """Rótulo do evento a partir do que existe na data: 'J', 'J+A', 'J+I' (ou 'A'/'I' se, num papel
    atípico, a data não tiver o evento de juros)."""
    partes = []
    if temJuros:
        partes.append("J")
    if amort > 0:
        partes.append("A")
    if incorp > 0:
        partes.append("I")
    return "+".join(partes) if partes else "J"


def _CalcPuB3Full(ticker, dataIso, taxa):
    """Resposta COMPLETA do calcPU da B3 (com VNA, cashFlowList) ou None."""
    tk = _NormalizarTicker(ticker)
    taxa = _Norm(taxa)
    chave = ("b3full", tk, dataIso, taxa)
    c = _CacheGet(chave)
    if c is not _CACHE_AUSENTE:
        return c
    dados = _GetB3(f"{B3_BASE}/calcPU/{tk}/{dataIso}/{taxa}")
    res = dados if isinstance(dados, dict) and dados.get("PU") is not None else None
    _CacheSet(chave, res)
    return res


def _FiFull(ticker, dataIso, taxa):
    """Resposta COMPLETA da FI (modo rate), via /deb, /cr ou bondbuilder, ou None.

    Mesma cadeia de fontes do `Preco`/`TaxaOp`: papel que só existe no bondbuilder (LCD, LF, CDB...)
    não é conhecido pelo /deb nem pelo /cr, mas o /bb devolve os MESMOS campos (m2m, taxedM2MRate,
    taxedType, dv01...) — é o que faz cpGrossUp/cpGrossUpTipo/cpDv01 funcionarem nesses papéis."""
    taxa = _Norm(taxa)
    chave = ("fifull", str(ticker).upper().strip(), dataIso, taxa)
    c = _CacheGet(chave)
    if c is not _CACHE_AUSENTE:
        return c
    dados = _PostFiAuto({"ticker": ticker, "date": dataIso, "rate": float(taxa)})
    if not (isinstance(dados, dict) and dados.get("m2m") is not None):
        dados = _BbFull(ticker, dataIso, taxa)
    res = dados if isinstance(dados, dict) and dados.get("m2m") is not None else None
    _CacheSet(chave, res)
    return res


def _BbFull(ticker, dataIso, taxa):
    """Resposta COMPLETA do bondbuilder (modo rate) ou None."""
    docId = _AcharDocId(ticker)
    if not docId:
        return None
    dados = _PostFiRaw(FI_BB_PATH, {"doc_id": docId, "date": dataIso, "rate": float(taxa)})
    return dados if _BbOk(dados) else None


def _FluxoFi(f):
    return [{"data": e.get("date"), "tipo": e.get("eventType"), "prazo": e.get("term"),
             "vf": e.get("futureValue"), "vp": e.get("presentValue")}
            for e in (f.get("cashFlowEvents") or [])]


def _FluxoB3(full):
    fluxo = []
    for e in (full.get("cashFlowList") or []):
        # a B3 usa nomes fat*Vf/fat*Vp variáveis; pega o 1º *Vf/*Vp preenchido.
        vf = next((v for k, v in e.items() if k.lower().endswith("vf") and v is not None), None)
        vp = next((v for k, v in e.items() if k.lower().endswith("vp") and v is not None), None)
        fluxo.append({"data": e.get("date"), "tipo": e.get("eventType"),
                      "prazo": e.get("term"), "vf": vf, "vp": vp})
    return fluxo


def Detalhes(ticker, dataIso, taxa):
    """
    Dict normalizado com tudo do papel (FI-first pela riqueza; B3 como fallback).
    Chaves: fonte, pu, pupar, vna, duration, durationMod, convexidade, dv01,
    vencimento, emissao, inicio, taxaEmissao, vne, grossup, grossupTipo,
    diPercent, spreadDI, jurosAcum, fluxo=[{data,tipo,prazo,vf,vp}], raw.
    Datas/taxaEmissao preferem a B3 getBondDetails (unidade nativa) quando houver.
    """
    tk = str(ticker).upper().strip()
    f = _FiFull(tk, dataIso, taxa)
    bond = BondDetailsB3(tk)
    if f:
        d = {
            "fonte": "fi", "pu": f.get("m2m"),
            "pupar": f.get("currentNotionalPlusAccruedInterest"),
            "vna": f.get("adjustedFaceValue"), "duration": f.get("maculayDuration"),
            "durationMod": f.get("modifiedDuration"), "convexidade": f.get("convexity"),
            "dv01": f.get("dv01"), "vencimento": f.get("maturityDate"),
            "emissao": f.get("issueDate"), "inicio": None, "taxaEmissao": f.get("issueRate"),
            "vne": None, "grossup": f.get("taxedM2MRate"), "grossupTipo": f.get("taxedType"),
            "diPercent": f.get("diPercentage"), "spreadDI": f.get("spreadOverDI"),
            "jurosAcum": f.get("accruedInterest"), "fluxo": _FluxoFi(f), "raw": f,
        }
    else:
        full = _CalcPuB3Full(tk, dataIso, taxa)
        if not full:
            return None
        d = {
            "fonte": "b3", "pu": full.get("PU"), "pupar": full.get("PUPar"),
            "vna": full.get("VNA"), "duration": full.get("duration"),
            "durationMod": None, "convexidade": None, "dv01": None,
            "vencimento": None, "emissao": None, "inicio": None, "taxaEmissao": None,
            "vne": None, "grossup": None, "grossupTipo": None, "diPercent": None,
            "spreadDI": None, "jurosAcum": full.get("interest"),
            "fluxo": _FluxoB3(full), "raw": full,
        }
    # Preferir datas/taxa/VNE da B3 getBondDetails (unidade nativa, datas limpas).
    if bond:
        d["vencimento"]   = bond.get("expiredate")   or d["vencimento"]
        d["emissao"]      = bond.get("issuedate")    or d["emissao"]
        d["inicio"]       = bond.get("startingdate") or d["inicio"]
        d["taxaEmissao"]  = bond.get("yield")        if bond.get("yield") is not None else d["taxaEmissao"]
        d["vne"]          = bond.get("vne")          if bond.get("vne")   is not None else d["vne"]
        d["aniversario"]  = bond.get("anniversaryday")
    return d


# =============================================================================
# BANCO CENTRAL — Sistema Gerenciador de Séries Temporais (SGS), público
# =============================================================================

BCB_BASE = "https://api.bcb.gov.br"


def BcbSerie(serie, ini=None, fim=None, ultimos=None):
    """Série do SGS/BCB. Lista [{'data':'dd/mm/aaaa','valor':'x'}] ou None.
    - ultimos=N → últimos N pontos (mais eficiente para 'valor atual');
    - senão intervalo ini..fim (strings 'dd/MM/yyyy'); sem nada → série inteira.
    Público (sem auth); respeita o proxy do banco."""
    chave = ("bcb", str(serie), str(ini), str(fim), str(ultimos))
    c = _CacheGet(chave)
    if c is not _CACHE_AUSENTE:
        return c
    if ultimos:
        url = f"{BCB_BASE}/dados/serie/bcdata.sgs.{serie}/dados/ultimos/{int(ultimos)}?formato=json"
    else:
        url = f"{BCB_BASE}/dados/serie/bcdata.sgs.{serie}/dados?formato=json"
        if ini:
            url += f"&dataInicial={ini}"
        if fim:
            url += f"&dataFinal={fim}"
    res = None
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}, method="GET")
        with _Abrir(req, BCB_BASE) as resp:
            dados = json.loads(resp.read().decode())
        if isinstance(dados, list):
            res = dados
    except Exception:
        res = None
    _CacheSet(chave, res)
    return res


def BcbValor(serie, data=None):
    """Valor (float) de uma série SGS na `data` (dd/MM/yyyy), com semântica as-of:
    se o dia exato não tem publicação (fim de semana/feriado, ou série mensal
    consultada no meio do mês), retorna o ÚLTIMO valor publicado ATÉ a data.
    Sem `data` → último disponível. None se ausente."""
    if data:
        d = BcbSerie(serie, ini=data, fim=data)
        if not d:  # dia sem publicação → último publicado até a data (as-of correto)
            d = BcbSerie(serie, fim=data)
    else:
        d = BcbSerie(serie, ultimos=1)
    if d:
        try:
            return float(str(d[-1]["valor"]).replace(",", "."))
        except Exception:
            return None
    return None


# ─── IBGE SIDRA: número-índice do IPCA (correção monetária entre datas) ─────────
SIDRA_BASE = "https://apisidra.ibge.gov.br"


def IpcaIndice(data=None):
    """IPCA número-índice (base dez/1993 = 100) do mês de referência de `data`
    (dd/MM/yyyy); sem `data` = último mês publicado. Fonte: IBGE SIDRA
    (tabela 1737, variável 2266). None se ausente.
    A razão de dois números-índice = fator de correção do IPCA entre as datas."""
    periodo = f"{data[6:10]}{data[3:5]}" if data else "last%201"
    chave = ("sidra_ipca", periodo)
    c = _CacheGet(chave)
    if c is not _CACHE_AUSENTE:
        return c
    res = None
    for p in (periodo, "last%201"):  # mês exato; se não publicado, cai no último
        url = f"{SIDRA_BASE}/values/t/1737/n1/all/v/2266/p/{p}"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}, method="GET")
            with _Abrir(req, SIDRA_BASE) as resp:
                dados = json.loads(resp.read().decode())
            # dados[0] é o cabeçalho; linhas de dados a partir de [1]
            if isinstance(dados, list) and len(dados) > 1:
                res = float(str(dados[-1]["V"]).replace(",", "."))
                break
        except Exception:
            res = None
    _CacheSet(chave, res)
    return res


# ─── acessores genéricos de campo (para UDFs específicas e =cpFi/=cpB3/=cpBond) ─

def CampoFi(ticker, dataIso, taxa, campo):
    """Um campo qualquer da resposta COMPLETA da FI (modo rate). None se ausente.
    Ex.: 'taxedM2MRate' (gross up), 'convexity', 'dv01', 'modifiedDuration', 'spreadOverDI'."""
    f = _FiFull(ticker, dataIso, taxa)
    return f.get(campo) if isinstance(f, dict) else None


def CampoB3(ticker, dataIso, taxa, campo):
    """Um campo qualquer da resposta COMPLETA do calcPU da B3. None se ausente.
    Ex.: 'VNA', 'PUPar', 'interest', 'issuer', 'method'."""
    full = _CalcPuB3Full(ticker, dataIso, taxa)
    return full.get(campo) if isinstance(full, dict) else None


def CampoBond(ticker, campo):
    """Um campo qualquer do getBondDetails da B3 (estático). None se ausente.
    Ex.: 'expiredate', 'issuedate', 'startingdate', 'vne', 'yield', 'anniversaryday'."""
    bond = BondDetailsB3(ticker)
    return bond.get(campo) if isinstance(bond, dict) else None


def FluxoRestante(ticker, dataIso, taxa):
    """Fluxo de caixa remanescente na data: lista [{data,tipo,prazo,vf,vp}] ou None.
    FI (cashFlowEvents, mais limpo) primeiro; B3 (cashFlowList) como fallback."""
    f = _FiFull(ticker, dataIso, taxa)
    if isinstance(f, dict) and f.get("cashFlowEvents"):
        return _FluxoFi(f)
    full = _CalcPuB3Full(ticker, dataIso, taxa)
    if isinstance(full, dict) and full.get("cashFlowList"):
        return _FluxoB3(full)
    return None


# Carrega o cache persistido (entradas dentro do TTL) ao importar o módulo.
_CarregarCacheDisco()
