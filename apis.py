# =============================================================================
# apis.py — cliente leve para BTG OTC Markets, FI Analytics e B3 Calculator
# -----------------------------------------------------------------------------
# Self-contained: usa apenas a biblioteca padrão (urllib), sem httpx e sem
# depender de outros repositórios. Pensado para o add-in do Excel: precifica
# via API (BTG → FI Analytics → B3 → bondbuilder).
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
from datetime import date
from urllib.parse import urlparse, unquote

from config import ENV_PATH

# Nomes das variáveis de ambiente (definidas no banco). O add-in lê os segredos
# daqui em produção; o .env serve só como fallback de desenvolvimento local.
ENV_TOKEN_B3 = "token_calc_b3"
ENV_KEY_BTG  = "token_btg_otc"   # API key do BTG OTC Markets (troca por um JWT de 30 min)
# FI Analytics: uma chave POR PRODUTO, lida SÓ da variável de ambiente do produto (sem .env e
# sem fallback para outra chave). Nomes iguais aos do banco.
ENV_KEY_FI_DEB = "token_fianalytics_debentures"       # Debêntures      (/deb)
ENV_KEY_FI_CR  = "token_fianalytics_cri_cra"          # CRIs/CRAs       (/cr)
ENV_KEY_FI_BB  = "token_fianalytics_bond_builder"     # Bond Builder    (/userdefinedbond)
ENV_KEY_FI_TP  = "token_fianalytics_titulos_publicos" # Títulos públicos (sem endpoint FI no add-in hoje)
ENV_USER_FI  = "user_fianalytics"   # e-mail do usuário FI, vai no corpo do bondbuilder (fallback: FIANALYTICS_USER)

# Proxy — env vars com a URL completa, incluindo usuário/senha, no formato
# http://USUARIO:SENHA@HOST:PORTA. Ausentes (dev) → conexão direta.
ENV_PROXY_HTTP  = "proxy_http"
ENV_PROXY_HTTPS = "proxy_https"

B3_BASE      = "https://api.calculadorarendafixa.com.br"
BTG_BASE     = "https://otcmarkets-api.btgpactualsolutions.com/api/v1/btg-otc-mkts"
BTG_CLIENT   = "calccp-addin"
BTG_TOKEN_SEG = 25 * 60   # o JWT do BTG vale 30 min; renova antes
FI_BASE      = "https://endpoint.fi-analytics.com.br"
FI_DEB_PATH  = "/deb/debenturecalculator"
FI_CR_PATH   = "/cr/cricracalculator"   # CRI/CRA (mesmos campos da resposta)
FI_BB_PATH   = "/userdefinedbond/bondcalculator"   # bondbuilder: papéis fora de /deb e /cr
TIMEOUT_SEG  = 6  # timeout curto por chamada para não travar o Excel

# Cache em disco (persiste entre reaberturas do Excel): só resultados VÁLIDOS
# (nunca None/erro), com validade de CACHE_TTL_SEG. Fora do TTL, refaz a chamada.
CACHE_TTL_SEG = 600  # 10 min
# Preço/taxa de data PASSADA não muda; o cadastro (getBondDetails) e o memo de fonte quase nunca.
# Esses ficam 12 h no disco, para reabrir o Excel sem refazer as chamadas.
CACHE_TTL_LONGO_SEG = 12 * 3600
_CHAVES_PRECO = {"btgpu": "btg", "b3pu": "b3", "fipu": "fi", "bbpu": "bb"}   # resultado de preço → fonte
_CHAVES_DATADAS = set(_CHAVES_PRECO) | {"btgyield", "b3yield", "fiyield", "bbyield", "b3full", "fifull"}
# Memo de fonte no disco. "fonte2" desde a v7.1.0 (ordem nova BTG → FI → B3): os memos antigos
# ("fonte", gravados com a B3 em 1º) são ignorados, senão furariam a ordem nova por até 12 h.
_CHAVE_MEMO = "fonte2"
_CHAVES_ESTATICAS = {"b3bond", _CHAVE_MEMO}
# Ordem das fontes de preço/taxa (o memo do ticker só reordena; o fallback é sempre mantido).
_ORDEM_FONTES = ["btg", "fi", "b3", "bb"]
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
_tokenBtg = None      # (jwt, obtido_em)
_cacheRespostas: dict = {}
_fonteTicker: dict = {}   # ticker(upper) -> "btg"|"fi"|"b3"|"bb": qual fonte respondeu (memo)

# Diagnóstico: motivo real da última falha de cada fonte, por ticker, para as UDFs mostrarem
# no lugar do genérico "APIs sem resposta". As funções de rede gravam em _ultimoErro; a
# cascata (Preco/PuPar/TaxaOp) copia para _errosTicker. Não muda nenhum resultado.
_ultimoErro = None
_errosTicker: dict = {}
# Resposta CRUA de cada chamada de preço, só em memória: cpPu, cpVna, cpGrossUp, cpDv01... com o
# mesmo papel/data/taxa usam a MESMA ida à API (antes cada família tinha o seu cache e repetia).
_cacheBruto: dict = {}
# (ticker, dataIso) -> {fonte: {"pupar", "vna"}}: PU Par e VNA não dependem da taxa, então
# qualquer preço já calculado do papel na data responde =cpPupar/=cpVna sem nova chamada.
_porData: dict = {}   # ticker(upper) -> {"B3": msg, "FI": msg, "bondbuilder": msg}
_ROTULO_DIAG = {"btg": "BTG", "fi": "FI", "b3": "B3", "bb": "bondbuilder"}


def _Resumo(txt, n=160):
    """Mensagem curta de um corpo de erro: o 'errorMessage'/'message' do JSON, senão o texto."""
    dados = txt
    try:
        for _ in range(2):   # a FI devolve JSON double-encoded
            if isinstance(dados, str):
                dados = json.loads(dados)
    except Exception:
        dados = txt
    if isinstance(dados, dict):
        msg = dados.get("errorMessage") or dados.get("message") or dados.get("error")
        if msg:
            cod = dados.get("statusCode")
            return (f"{cod} " if cod and cod != 200 else "") + str(msg)[:n]
    return " ".join(str(txt).split())[:n]


def _MsgExcecao(e):
    """Texto da exceção de rede: 'HTTP 403 <motivo>' ou 'Tipo: detalhe'."""
    if isinstance(e, urllib.error.HTTPError):
        try:
            corpo = e.read().decode(errors="replace")
        except Exception:
            corpo = ""
        return f"HTTP {e.code} {_Resumo(corpo) or e.reason}".strip()
    return f"{type(e).__name__}: {str(e)[:160]}"


def _Falhou(msg):
    global _ultimoErro
    _ultimoErro = msg


def _RegistrarFalha(tk, src):
    """Chamada pela cascata quando a fonte `src` não respondeu. Se a chamada nem foi à rede
    (resultado None vindo do cache), mantém o motivo registrado da 1ª vez."""
    erros = _errosTicker.setdefault(tk, {})
    rotulo = _ROTULO_DIAG.get(src, src)
    if _ultimoErro:
        erros[rotulo] = _ultimoErro
    else:
        erros.setdefault(rotulo, "resposta sem valor")


def ErroDetalhado(ticker):
    """'BTG: ... | FI: ... | B3: ... | bondbuilder: ...' com o motivo de cada fonte que falhou, ou ''."""
    erros = _errosTicker.get(str(ticker).upper().strip()) or {}
    return " | ".join(f"{k}: {erros[k]}" for k in ("BTG", "FI", "B3", "bondbuilder") if k in erros)

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
            chave = tuple(entrada["k"])
            if isinstance(entrada, dict) and agora - entrada["t"] < _Ttl(chave):
                if chave[0] == _CHAVE_MEMO:
                    _fonteTicker[chave[1]] = entrada["v"]
                else:
                    _cacheRespostas[chave] = entrada["v"]
                    _Indexar(chave, entrada["v"])
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
                 if isinstance(v, dict) and agora - v.get("t", 0) < _Ttl(tuple(v.get("k") or ()))}
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


def _Persistir(chave, valor):
    global _discoSujo
    _discoRegistro["|".join(map(str, chave))] = {
        "k": list(chave), "v": valor, "t": time.time(),
    }
    _discoSujo = True
    _FlushDisco()


def _CacheSet(chave, valor):
    _cacheRespostas[chave] = valor
    _Indexar(chave, valor)
    if valor is not None:            # só persiste resultado válido (nunca None/erro)
        _Persistir(chave, valor)


def _Ttl(chave):
    """Validade no disco: 12 h p/ cadastro, memo de fonte e preço/taxa de data passada; 10 min no resto."""
    tipo = chave[0] if chave else None
    if tipo in _CHAVES_ESTATICAS:
        return CACHE_TTL_LONGO_SEG
    if tipo in _CHAVES_DATADAS and len(chave) > 2 and str(chave[2]) < date.today().isoformat():
        return CACHE_TTL_LONGO_SEG
    return CACHE_TTL_SEG


def _Bruto(chave, buscar):
    """Resposta crua de uma chamada, cacheada SÓ em memória (inclusive None)."""
    if chave not in _cacheBruto:
        _cacheBruto[chave] = buscar()
    return _cacheBruto[chave]


def _Indexar(chave, valor):
    """Registra PU Par/VNA de um resultado de preço em _porData (independem da taxa)."""
    src = _CHAVES_PRECO.get(chave[0]) if chave else None
    if src and isinstance(valor, dict) and len(chave) > 2:
        campos = {c: valor[c] for c in ("pupar", "vna") if valor.get(c) is not None}
        if campos:
            _porData.setdefault((chave[1], chave[2]), {})[src] = campos


def _Memo(tk, src):
    """Memoriza a fonte que respondeu pelo ticker — em memória e no disco (vale 12 h), para
    que ao reabrir o Excel um papel só-FI não teste a B3 (nem faça o login dela) de novo."""
    if _fonteTicker.get(tk) != src:
        _fonteTicker[tk] = src
        _Persistir((_CHAVE_MEMO, tk), src)


def DadoSemTaxa(ticker, dataIso, campo):
    """"pupar" ou "vna" do papel na data se algum preço (em QUALQUER taxa) já foi calculado nesta
    sessão ou está no cache em disco — sem rede. Respeita a ordem do memo de fonte. float ou None."""
    tk = str(ticker).upper().strip()
    for t in dict.fromkeys((tk, _NormalizarTicker(tk))):
        porFonte = _porData.get((t, dataIso))
        if porFonte:
            for src in _OrdemFontes(tk, _ORDEM_FONTES):
                v = (porFonte.get(src) or {}).get(campo)
                if v is not None:
                    return v
    return None


def _Abrir(req, base=None):
    """Abre a requisição (timeout padrão). `base` mantido só p/ compat. da chamada."""
    return _opener.open(req, timeout=TIMEOUT_SEG)


def LimparCache():
    """Esvazia o cache (memória+disco), o token e o memo."""
    global _tokenB3, _tokenBtg, _discoSujo
    _cacheRespostas.clear()
    _fonteTicker.clear()
    _errosTicker.clear()
    _cacheBruto.clear()
    _porData.clear()
    _discoRegistro.clear()
    _discoSujo = False
    _tokenB3 = None
    _tokenBtg = None
    try:
        os.remove(_CACHE_FILE)
    except OSError:
        pass



# -----------------------------------------------------------------------------
# BTG OTC Markets (1ª fonte)
# -----------------------------------------------------------------------------
# POST /authenticate {api_key, client_id} → AccessToken (JWT, 30 min) → header
# "Authorization: Bearer". GET /pricing/pu?ticker&rate&date → {pu, duration, pu_par}
# e GET /pricing/rate?ticker&pu&date → {rate}. Taxa em % a.a. na unidade nativa (igual à
# B3). Conferido 05/11/2026: 100 papéis (DEB/CRI/CRA; IPCA+, DI+, %DI, PRÉ) em D0 e D+1 →
# PU, taxa, duration e PU Par IDÊNTICOS aos da B3. NÃO devolve VNA nem fluxo. Título
# público não existe no BTG (404 → segue para a próxima fonte).

def _TextoBtg(raw):
    """O BTG manda as mensagens de erro em latin-1 (acentos), não em UTF-8."""
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1")


def _MsgBtg(e):
    """'HTTP 404 Ticker not found.' a partir do ApiClientError do BTG, senão o texto da exceção."""
    if isinstance(e, urllib.error.HTTPError):
        try:
            corpo = json.loads(_TextoBtg(e.read()))
            msg = corpo.get("ApiClientError") if isinstance(corpo, dict) else None
        except Exception:
            msg = None
        return f"HTTP {e.code} {(msg or e.reason or '')[:160]}".strip()
    return _MsgExcecao(e)


def _LoginBtg():
    """JWT do BTG a partir da API key, ou None."""
    chave = os.getenv(ENV_KEY_BTG)
    if not chave:
        _Falhou(f"variável {ENV_KEY_BTG} não setada")
        return None
    try:
        req = urllib.request.Request(
            f"{BTG_BASE}/authenticate",
            data=json.dumps({"api_key": chave, "client_id": BTG_CLIENT}).encode(),
            headers={"Content-Type": "application/json"}, method="POST",
        )
        with _Abrir(req, BTG_BASE) as resp:
            dados = json.loads(_TextoBtg(resp.read()))
        token = dados.get("AccessToken") if isinstance(dados, dict) else None
        if not token:
            _Falhou(f"login BTG sem token: {_Resumo(json.dumps(dados))}")
        return token
    except Exception as e:
        _Falhou(f"login BTG: {_MsgBtg(e)}")
        return None


def _ObterTokenBtg():
    global _tokenBtg
    if _tokenBtg is None or time.time() - _tokenBtg[1] > BTG_TOKEN_SEG:
        token = _LoginBtg()
        _tokenBtg = (token, time.time()) if token else None
    return _tokenBtg[0] if _tokenBtg else None


def _GetBtg(caminho, params):
    """GET autenticado no BTG. Em 401 renova o token e tenta mais uma vez. dict ou None."""
    global _tokenBtg
    url = f"{BTG_BASE}/{caminho}?" + "&".join(f"{k}={v}" for k, v in params.items())
    for tentativa in range(2):
        token = _ObterTokenBtg()
        if token is None:
            return None
        try:
            req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"}, method="GET")
            with _Abrir(req, BTG_BASE) as resp:
                dados = json.loads(_TextoBtg(resp.read()))
            return dados[0] if isinstance(dados, list) and dados else dados   # a doc diz lista; vem objeto
        except urllib.error.HTTPError as erro:
            if erro.code == 401 and tentativa == 0:
                _tokenBtg = None
                continue
            _Falhou(_MsgBtg(erro))
            return None
        except Exception as e:
            _Falhou(_MsgBtg(e))
            return None
    return None


_RE_CETIP_TP = re.compile(r"^(760199|950199|100000|210100)\d{8}$")   # NTN-B, NTN-F, LTN, LFT


def _EhTituloPublico(tk):
    """Título público (NTN-B/NTN-F pelo nome amigável ou código CETIP de NTN-B/F, LTN, LFT).
    BTG OTC e FI /deb e /cr só têm crédito privado → esses pulam os dois sem ir à rede (vão p/ a B3)."""
    t = str(tk).upper().strip()
    return _NormalizarTicker(t) != t or bool(_RE_CETIP_TP.match(t))


def PrecoBtg(ticker, dataIso, taxa):
    """/pricing/pu → {'pu', 'duration', 'pupar', 'vna': None} ou None. taxa em % a.a."""
    tk = str(ticker).upper().strip()
    if _EhTituloPublico(tk):
        _Falhou("título público (BTG só tem crédito privado)")
        return None
    taxa = _Norm(taxa)
    chaveCache = ("btgpu", tk, dataIso, taxa)
    cacheado = _CacheGet(chaveCache)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    dados = _GetBtg("pricing/pu", {"ticker": tk, "rate": taxa, "date": dataIso})
    resultado = None
    if isinstance(dados, dict):
        pu, dur, pupar = dados.get("pu"), dados.get("duration"), dados.get("pu_par")
        if pu is not None and float(pu) > 0:
            resultado = {
                "pu": float(pu),
                "duration": float(dur) if dur is not None else None,
                "pupar": float(pupar) if pupar is not None else None,
                "vna": None,   # o BTG não devolve VNA → =cpVna cai na FI/B3 (ver Vna())
            }
        else:
            _Falhou(f"resposta sem pu: {_Resumo(json.dumps(dados))}")
    _CacheSet(chaveCache, resultado)
    return resultado


def TaxaBtg(ticker, dataIso, pu):
    """/pricing/rate → taxa em % a.a. ou None."""
    tk = str(ticker).upper().strip()
    if _EhTituloPublico(tk):
        _Falhou("título público (BTG só tem crédito privado)")
        return None
    pu = _Norm(pu)
    chaveCache = ("btgyield", tk, dataIso, pu)
    cacheado = _CacheGet(chaveCache)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    dados = _GetBtg("pricing/rate", {"ticker": tk, "pu": pu, "date": dataIso})
    resultado = None
    if isinstance(dados, dict):
        taxa = dados.get("rate")
        if taxa is not None and float(taxa) > 0:
            resultado = float(taxa)
        else:
            _Falhou(f"resposta sem rate: {_Resumo(json.dumps(dados))}")
    _CacheSet(chaveCache, resultado)
    return resultado

_VAZIOS_BTG = {None, "", "-", "--", "nan", "None"}


def CadastroBtg(ticker):
    """Linha do cadastro do BTG (/mktdata/instrument/all?symbols=) do papel.

    dict com os campos do papel; {} se o BTG não conhece o papel; None se a chamada falhou
    (motivo em UltimoErro()). Os campos de mercado (indicativa ANBIMA, ref NTN-B) são os da
    última divulgação da ANBIMA — o BTG não aceita data. Título público não está lá."""
    tk = str(ticker).upper().strip()
    if _EhTituloPublico(tk):
        return {}
    chave = ("btgcad", tk)
    c = _CacheGet(chave)
    if c is not _CACHE_AUSENTE:
        return c
    _Falhou(None)
    dados = _GetBtg("mktdata/instrument/all", {"pagination": "False", "symbols": tk})
    if dados is None:
        return None                         # falha de rede/auth: não cacheia, tenta de novo depois
    res = dados if isinstance(dados, dict) and str(dados.get("symbol", "")).upper() == tk else {}
    _CacheSet(chave, res)
    return res


def CampoBtg(ticker, campo):
    """Um campo do cadastro do BTG, com os marcadores de vazio ('-', '--', 'nan') como None."""
    cad = CadastroBtg(ticker)
    v = cad.get(campo) if isinstance(cad, dict) else None
    return None if v in _VAZIOS_BTG or (isinstance(v, str) and v.strip() in _VAZIOS_BTG) else v


def IncentivadaFi(ticker):
    """Tipo fiscal pela FI, para o =cpIncentivada quando o BTG não sabe. Devolve:
    'GROSS_UP' (debênture isenta) | 'AFTER_TAX' (debênture tributada) | 'SEM_TIPO' (a FI conhece a
    debênture mas não manda o campo — ex.: TAEEB5) | 'CRI_CRA' (o /cr respondeu) | None (nenhum dos
    dois conhece o papel, ou falha de rede — motivo em UltimoErro()).
    O tipo fiscal não depende da taxa nem da data: vai uma taxa-sonda de 10% em hoje."""
    tk = str(ticker).upper().strip()
    if _EhTituloPublico(tk):
        return None
    chave = ("fiincent", tk)
    c = _CacheGet(chave)
    if c is not _CACHE_AUSENTE:
        return c
    corpo = {"ticker": tk, "date": date.today().isoformat(), "rate": 10.0}
    _Falhou(None)
    res = None
    dados = _PostFi(corpo, FI_DEB_PATH)
    if isinstance(dados, dict) and dados.get("m2m") is not None:
        tipo = dados.get("taxedType")
        res = tipo if tipo in ("GROSS_UP", "AFTER_TAX") else "SEM_TIPO"
    else:
        dados = _PostFi(corpo, FI_CR_PATH)
        if isinstance(dados, dict) and dados.get("m2m") is not None:
            res = "CRI_CRA"
    _CacheSet(chave, res)
    return res


def UltimoErro():
    """Motivo da última falha de rede registrada (texto) ou None."""
    return _ultimoErro


# -----------------------------------------------------------------------------
# B3 Calculator
# -----------------------------------------------------------------------------

def _LoginB3():
    """Obtém o token de sessão via POST /login. Retorna o token ou None."""
    token = _Cred(ENV_TOKEN_B3, "B3_CALC_TOKEN")
    if not token:
        _Falhou(f"variável {ENV_TOKEN_B3} não setada")
        return None
    try:
        corpo = json.dumps({"token": token}).encode()
        req = urllib.request.Request(
            f"{B3_BASE}/login", data=corpo,
            headers={"Content-Type": "application/json"}, method="POST",
        )
        with _Abrir(req, B3_BASE) as resp:
            dados = json.loads(resp.read().decode())
        if not dados.get("Authorization"):
            _Falhou(f"login B3 sem token: {_Resumo(json.dumps(dados))}")
        return dados.get("Authorization")
    except Exception as e:
        _Falhou(f"login B3: {_MsgExcecao(e)}")
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
            _Falhou(_MsgExcecao(erro))
            return None
        except Exception as e:
            _Falhou(_MsgExcecao(e))
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


def _CalcPuB3Bruto(tk, dataIso, taxa):
    """calcPU cru da B3 (ticker já normalizado, taxa já _Norm), 1 chamada por sessão."""
    return _Bruto(("b3calc", tk, dataIso, taxa),
                  lambda: _GetB3(f"{B3_BASE}/calcPU/{tk}/{dataIso}/{taxa}"))


def PrecoB3(ticker, dataIso, taxa):
    """calcPU → {'pu': float, 'duration': anos|None} ou None. taxa em % a.a."""
    ticker = _NormalizarTicker(ticker)
    taxa = _Norm(taxa)
    chaveCache = ("b3pu", ticker, dataIso, taxa)
    cacheado = _CacheGet(chaveCache)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    dados = _CalcPuB3Bruto(ticker, dataIso, taxa)
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

def _ChaveFi(path):
    """API key do produto FI dono do endpoint (1ª parte do path), só da variável de ambiente."""
    nome = _VarChaveFi(path)
    return os.getenv(nome) if nome else None


def _VarChaveFi(path):
    """Nome da variável de ambiente da chave do produto FI dono do endpoint."""
    return {
        "deb": ENV_KEY_FI_DEB,
        "cr": ENV_KEY_FI_CR,
        "userdefinedbond": ENV_KEY_FI_BB,
    }.get(path.strip("/").split("/")[0])


def _PostFi(corpo, path=FI_DEB_PATH):
    """POST num endpoint FI. Resposta é double-encoded. dict ou None."""
    chave = _ChaveFi(path)
    if not chave:
        _Falhou(f"variável {_VarChaveFi(path)} não setada")
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
    except Exception as e:
        _Falhou(_MsgExcecao(e))
        return None


def _PostFiAuto(corpo):
    """Tenta debênture; se não for (ex.: CRI/CRA), tenta o endpoint de CRI/CRA.
    Os dois endpoints devolvem os mesmos campos (m2m, m2mRate, ...).
    Título público pula a FI sem ir à rede: /deb e /cr não têm, e a recusa leva ~7 s."""
    if _EhTituloPublico(corpo.get("ticker", "")):
        _Falhou("título público (FI /deb e /cr só têm crédito privado)")
        return None
    motivos = []
    for path in (FI_DEB_PATH, FI_CR_PATH):
        _Falhou(None)
        dados = _PostFi(corpo, path)
        if isinstance(dados, dict) and (
            dados.get("m2m") is not None or dados.get("m2mRate") is not None
        ):
            return dados
        motivo = _ultimoErro or (_Resumo(json.dumps(dados)) if dados is not None else "sem resposta")
        motivos.append(f"/{path.strip('/').split('/')[0]} {motivo}")
    _Falhou("; ".join(motivos))
    return None


def _FiBruto(ticker, dataIso, taxa):
    """/deb→/cr cru em modo rate (taxa já _Norm), 1 chamada por sessão."""
    tk = str(ticker).upper().strip()
    return _Bruto(("fiauto", tk, dataIso, taxa),
                  lambda: _PostFiAuto({"ticker": tk, "date": dataIso, "rate": float(taxa)}))


def PrecoFi(ticker, dataIso, taxa):
    """Modo rate → {'pu': m2m, 'duration': maculayDuration anos|None} ou None."""
    taxa = _Norm(taxa)
    chaveCache = ("fipu", ticker, dataIso, taxa)
    cacheado = _CacheGet(chaveCache)
    if cacheado is not _CACHE_AUSENTE:
        return cacheado

    dados = _FiBruto(ticker, dataIso, taxa)
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
# Rota única POST /userdefinedbond/bondcalculator. O papel vai pelo NOME
# (`bond_name` = ticker) junto com o e-mail do usuário (ENV_USER_FI) — sem passo
# de descoberta de id. Mesmo motor da FI: manda `rate` → volta `m2m` (PU) +
# duration; manda `pu` → volta `m2mRate`. Só dispara depois de B3 e FI falharem.

def _PostFiRaw(path, corpo):
    """POST num endpoint FI via http.client (NÃO via urllib).

    Motivo: o WAF do bondbuilder exige o header 'x-api-key' em MINÚSCULO; o urllib.request
    title-caseia todo header ('X-Api-Key') e o gateway devolve 502. O http.client
    preserva a caixa. Suporta o proxy do banco (CONNECT tunnel c/ Proxy-Authorization).
    Retorna dict/list (double-encoded resolvido) ou None."""
    chave = _ChaveFi(path)
    if not chave:
        _Falhou(f"variável {ENV_KEY_FI_BB} não setada")
        return None
    corpoBytes = json.dumps(corpo).encode()
    headers = {"content-type": "application/json; charset=utf-8", "accept": "*/*", "x-api-key": chave}
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
    except Exception as e:
        _Falhou(_MsgExcecao(e))
        return None
    finally:
        if conn is not None:
            conn.close()
    if status >= 400:
        _Falhou(f"HTTP {status} {_Resumo(raw)}")
        return None
    try:
        externo = json.loads(raw)
        return json.loads(externo) if isinstance(externo, str) else externo
    except Exception:
        _Falhou(f"resposta inválida: {' '.join(raw.split())[:120]}")
        return None


def _CorpoBb(ticker, dataIso, campo, valor):
    """Corpo do bondbuilder, ou None sem o e-mail do usuário (a rota exige)."""
    email = _Cred(ENV_USER_FI, "FIANALYTICS_USER")
    if not email:
        _Falhou(f"variável {ENV_USER_FI} não setada")
        return None
    return {"user_email": email, "bond_name": str(ticker).upper().strip(),
            "date": dataIso, "settlement_date": dataIso,
            "useIntradayCurve": True, campo: float(valor)}


def _PostBb(ticker, dataIso, campo, valor):
    """Chama o bondbuilder (`campo` = "rate" ou "pu"). dict sem erro ou None."""
    corpo = _CorpoBb(ticker, dataIso, campo, valor)
    if corpo is None:
        return None
    dados = _PostFiRaw(FI_BB_PATH, corpo)
    if dados is not None and not _BbOk(dados):
        _Falhou(_Resumo(json.dumps(dados)))
    return dados if _BbOk(dados) else None


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
    dados = _BbFull(ticker, dataIso, taxa)
    if dados:
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
    dados = _PostBb(ticker, dataIso, "pu", pu)
    if dados:
        taxa = dados.get("m2mRate")
        if taxa is not None and float(taxa) > 0:
            resultado = float(taxa) * 100
    _CacheSet(chaveCache, resultado)
    return resultado


# ─── roteamento BTG → FI → B3 → bondbuilder com memo de fonte por ticker ─────
# Tenta as fontes em ordem; a que respondeu (só em sucesso) é memorizada e passa
# a ser tentada 1º nas próximas — mas o fallback é SEMPRE mantido (o memo só
# reordena, nunca bloqueia). bondbuilder fica por último.

def _OrdemFontes(tk, todas):
    memo = _fonteTicker.get(tk)
    return [memo] + [s for s in todas if s != memo] if memo in todas else list(todas)


_FONTES_PRECO = {"btg": PrecoBtg, "fi": PrecoFi, "b3": PrecoB3, "bb": PrecoBb}


def Preco(ticker, dataIso, taxa):
    """PU via BTG→FI→B3→bondbuilder com memo de fonte. dict {pu,duration,pupar,vna} ou None."""
    tk = str(ticker).upper().strip()
    fontes = _FONTES_PRECO
    for src in _OrdemFontes(tk, _ORDEM_FONTES):
        _Falhou(None)
        r = fontes[src](tk, dataIso, taxa)
        if r:
            _Memo(tk, src)
            return r
        _RegistrarFalha(tk, src)
    return None


def _CampoSemTaxa(ticker, dataIso, taxa, campo):
    """`campo` ("pupar"/"vna") via BTG→FI→B3→bondbuilder: a 1ª fonte que devolver o campo.

    Diferente do Preco(), que para na 1ª fonte com PU: aqui uma fonte que precifica mas não
    traz o campo cede a vez à próxima (o BTG não traz VNA). O campo não depende da taxa, então
    `taxa` só monta a chamada. O memo guarda a 1ª fonte que PRECIFICOU o papel, não a que trouxe
    o campo — senão o =cpVna de um papel do BTG empurraria o =cpPu dele para a FI. float ou None."""
    tk = str(ticker).upper().strip()
    conhecido = DadoSemTaxa(tk, dataIso, campo)
    if conhecido is not None:
        return conhecido
    precificou = False
    for src in _OrdemFontes(tk, _ORDEM_FONTES):
        _Falhou(None)
        r = _FONTES_PRECO[src](tk, dataIso, taxa)
        if r:
            if not precificou:
                _Memo(tk, src)
                precificou = True
            if r.get(campo) is not None:
                return r[campo]
        else:
            _RegistrarFalha(tk, src)
    return None


def PuPar(ticker, dataIso, taxa):
    """PU Par via BTG→FI→B3→bondbuilder (a 1ª fonte que devolver o PU Par). float ou None."""
    return _CampoSemTaxa(ticker, dataIso, taxa, "pupar")


def Vna(ticker, dataIso, taxa):
    """VNA via BTG→FI→B3→bondbuilder (o BTG não traz VNA e cede a vez). float ou None."""
    return _CampoSemTaxa(ticker, dataIso, taxa, "vna")


def TaxaOp(ticker, dataIso, pu):
    """Taxa (% a.a.) via BTG→FI→B3→bondbuilder com memo de fonte. float ou None."""
    tk = str(ticker).upper().strip()
    fontes = {"btg": TaxaBtg, "fi": TaxaFi, "b3": TaxaB3, "bb": TaxaBb}
    for src in _OrdemFontes(tk, _ORDEM_FONTES):
        _Falhou(None)
        r = fontes[src](tk, dataIso, pu)
        if r is not None:
            _Memo(tk, src)
            return r
        _RegistrarFalha(tk, src)
    return None


def FonteTicker(ticker):
    """Fonte já conhecida do ticker ("btg"/"fi"/"b3"/"bb"), ou None se ainda não sabemos.
    Só lê o memo — nunca toca a rede, e por isso não precisa de data nem taxa."""
    return _fonteTicker.get(str(ticker).upper().strip())


def FonteCalculo(ticker, dataIso, taxa):
    """Qual fonte precifica este ticker: "btg", "fi", "b3", "bb" ou None (nenhuma responde).

    Se o memo já sabe (algum =cpPu/=cpTaxa/=cpDur já rodou para o papel nesta
    sessão), devolve na hora, sem rede. Senão roda a MESMA cascata do Preco() —
    então a resposta é sempre a fonte que de fato precificaria o papel, e não
    depende da ordem em que o Excel calculou as células. `dataIso`/`taxa` só
    servem para essa sondagem: qual das fontes responde não depende deles."""
    tk = str(ticker).upper().strip()
    memo = FonteTicker(tk)
    if memo:
        return memo
    return _fonteTicker.get(tk) if Preco(tk, dataIso, taxa) else None


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
    dados = _CalcPuB3Bruto(tk, dataIso, taxa)
    res = dados if isinstance(dados, dict) and dados.get("PU") is not None else None
    _CacheSet(chave, res)
    return res


def _FiFull(ticker, dataIso, taxa):
    """Resposta COMPLETA da FI (modo rate), via /deb, /cr ou bondbuilder, ou None.

    Mesma cadeia de fontes do `Preco`/`TaxaOp`: papel que só existe no bondbuilder (LCD, LF, CDB...)
    não é conhecido pelo /deb nem pelo /cr, mas o bondbuilder devolve os MESMOS campos (m2m, taxedM2MRate,
    taxedType, dv01...) — é o que faz cpGrossUp/cpGrossUpTipo/cpDv01 funcionarem nesses papéis."""
    taxa = _Norm(taxa)
    chave = ("fifull", str(ticker).upper().strip(), dataIso, taxa)
    c = _CacheGet(chave)
    if c is not _CACHE_AUSENTE:
        return c
    dados = _FiBruto(ticker, dataIso, taxa)
    if not (isinstance(dados, dict) and dados.get("m2m") is not None):
        dados = _BbFull(ticker, dataIso, taxa)
    res = dados if isinstance(dados, dict) and dados.get("m2m") is not None else None
    _CacheSet(chave, res)
    return res


def _BbFull(ticker, dataIso, taxa):
    """Resposta COMPLETA do bondbuilder (modo rate) ou None. 1 chamada por sessão."""
    tk = str(ticker).upper().strip()
    taxa = _Norm(taxa)
    return _Bruto(("bbrate", tk, dataIso, taxa), lambda: _PostBb(tk, dataIso, "rate", taxa))


def _FluxoFi(f):
    return [{"data": e.get("date"), "tipo": e.get("eventType"), "prazo": e.get("term"),
             "vf": e.get("futureValue"), "vp": e.get("presentValue")}
            for e in (f.get("cashFlowEvents") or [])]


def _FluxoB3(full):
    fluxo = []
    for e in (full.get("cashFlowList") or []):
        # Valor futuro/presente do evento = finalValue/presentValue. Os fat*Vf/fat*Vp são
        # fatores e vêm None nos papéis conferidos (IPCA, DI, PRE) — ficam só de reserva.
        vf = e.get("finalValue")
        if vf is None:
            vf = next((v for k, v in e.items() if k.lower().endswith("vf") and v is not None), None)
        vp = e.get("presentValue")
        if vp is None:
            vp = next((v for k, v in e.items() if k.lower().endswith("vp") and v is not None), None)
        fluxo.append({"data": e.get("date"), "tipo": e.get("eventType"),
                      "prazo": e.get("workingDays"), "vf": vf, "vp": vp})
    return fluxo


def FluxoFuturo(ticker, dataIso, taxa):
    """Eventos ESTRITAMENTE posteriores a `dataIso` com valor futuro: (fonte, [{data,tipo,prazo,vf,vp}])
    ou None. Cascata FI (/deb, /cr, bondbuilder) → B3, na ordem do memo de fonte do ticker
    (o BTG não devolve fluxo).

    O valor futuro não depende da taxa (conferido em IPCA, DI e PRE) — `taxa` só monta a chamada.
    A FI recusa data futura; a B3 aceita. Papel DI: as duas projetam com curvas próprias e diferem."""
    tk = str(ticker).upper().strip()
    for src in _OrdemFontes(tk, ["fi", "b3"]):
        if src == "b3":
            full = _CalcPuB3Full(tk, dataIso, taxa)
            if isinstance(full, dict):
                return "b3", _FluxoB3(full)
        else:
            f = _FiFull(tk, dataIso, taxa)
            if isinstance(f, dict):
                return "fi", _FluxoFi(f)
    return None


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
