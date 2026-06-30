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
import json
import urllib.request
import urllib.error

from config import ENV_PATH

# Nomes das variáveis de ambiente (definidas no banco). O add-in lê os segredos
# daqui em produção; o .env serve só como fallback de desenvolvimento local.
ENV_TOKEN_B3 = "token_calc_b3"
ENV_KEY_FI   = "token_fianalytics"

# Proxy — env vars com a URL completa, incluindo usuário/senha, no formato
# http://USUARIO:SENHA@HOST:PORTA. Ausentes (dev) → conexão direta.
ENV_PROXY_HTTP  = "proxy_http"
ENV_PROXY_HTTPS = "proxy_https"

B3_BASE      = "https://api.calculadorarendafixa.com.br"
FI_BASE      = "https://endpoint.fi-analytics.com.br"
FI_DEB_PATH  = "/deb/debenturecalculator"
FI_CR_PATH   = "/cr/cricracalculator"   # CRI/CRA (mesmos campos da resposta)
TIMEOUT_SEG  = 6  # timeout curto por chamada para não travar o Excel

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


def LimparCache():
    """Esvazia o cache de respostas e o token (força novas chamadas)."""
    global _tokenB3
    _cacheRespostas.clear()
    _tokenB3 = None


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
        with _opener.open(req, timeout=TIMEOUT_SEG) as resp:
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
            with _opener.open(req, timeout=TIMEOUT_SEG) as resp:
                return json.loads(resp.read().decode())
        except urllib.error.HTTPError as erro:
            if erro.code == 401 and tentativa == 0:
                _tokenB3 = None
                continue
            return None
        except Exception:
            return None
    return None


def PrecoB3(ticker, dataIso, taxa):
    """calcPU → {'pu': float, 'duration': anos|None} ou None. taxa em % a.a."""
    taxa = _Norm(taxa)
    chaveCache = ("b3pu", ticker, dataIso, taxa)
    cacheado = _cacheRespostas.get(chaveCache, _CACHE_AUSENTE)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    dados = _GetB3(f"{B3_BASE}/calcPU/{ticker}/{dataIso}/{taxa}")
    resultado = None
    if isinstance(dados, dict):
        pu    = dados.get("PU")
        dur   = dados.get("duration")
        pupar = dados.get("PUPar")
        if pu is not None and float(pu) > 0:
            resultado = {
                "pu": float(pu),
                "duration": float(dur) if dur is not None else None,
                "pupar": float(pupar) if pupar is not None else None,
            }
    _cacheRespostas[chaveCache] = resultado
    return resultado


def TaxaB3(ticker, dataIso, pu):
    """calcYield → yield em % a.a. ou None."""
    pu = _Norm(pu)
    chaveCache = ("b3yield", ticker, dataIso, pu)
    cacheado = _cacheRespostas.get(chaveCache, _CACHE_AUSENTE)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    dados = _GetB3(f"{B3_BASE}/calcYield/{ticker}/{dataIso}/{pu}")
    resultado = None
    if isinstance(dados, dict):
        taxa = dados.get("yield")
        if taxa is not None and float(taxa) > 0:
            resultado = float(taxa)
    _cacheRespostas[chaveCache] = resultado
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
        with _opener.open(req, timeout=TIMEOUT_SEG) as resp:
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
    cacheado = _cacheRespostas.get(chaveCache, _CACHE_AUSENTE)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    dados = _PostFiAuto({"ticker": ticker, "date": dataIso, "rate": float(taxa)})
    resultado = None
    if isinstance(dados, dict):
        pu    = dados.get("m2m")
        dur   = dados.get("maculayDuration")
        pupar = dados.get("currentNotionalPlusAccruedInterest")
        if pu is not None and float(pu) > 0:
            resultado = {
                "pu": float(pu),
                "duration": float(dur) if dur is not None else None,
                "pupar": float(pupar) if pupar is not None else None,
            }
    _cacheRespostas[chaveCache] = resultado
    return resultado


def TaxaFi(ticker, dataIso, pu):
    """Modo pu → yield em % a.a. (m2mRate × 100) ou None."""
    pu = _Norm(pu)
    chaveCache = ("fiyield", ticker, dataIso, pu)
    cacheado = _cacheRespostas.get(chaveCache, _CACHE_AUSENTE)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    dados = _PostFiAuto({"ticker": ticker, "date": dataIso, "pu": float(pu)})
    resultado = None
    if isinstance(dados, dict):
        taxa = dados.get("m2mRate")
        if taxa is not None and float(taxa) > 0:
            resultado = float(taxa) * 100
    _cacheRespostas[chaveCache] = resultado
    return resultado
