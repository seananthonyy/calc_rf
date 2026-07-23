# =============================================================================
# basedados.py — leitura do trades.db (projeto de negociação secundária)
# -----------------------------------------------------------------------------
# Única parte do add-in que lê uma base local. Serve às fórmulas ANBIMA
# (=cpAnbimaRef, =cpAnbimaIndicativo, =cpAnbimaIndicativoHistorico), que
# consultam dados que não existem em nenhuma API.
#
# REGRAS:
#   - SOMENTE LEITURA. A conexão é aberta com a URI `?mode=ro`, então o SQLite
#     recusa qualquer escrita e não cria journal/WAL ao lado do arquivo — o
#     pipeline de negociação segue escrevendo na base sem disputa.
#   - Sem dependência externa: só `sqlite3` da stdlib.
#   - O caminho do arquivo NÃO é embutido aqui: vem de config.ResolverTradesDb().
#   - Cache por processo (CACHE_TTL_SEG) para não bater no arquivo — que pode
#     estar numa share de rede — a cada recálculo do Excel. =cpLimparCache limpa.
#
# Tabelas usadas:
#   InfoAtivos(cdTicker, cdReferencia, ...)                  — vértice de referência
#   AnbimaIndicativos(cdTicker, dtReferencia, vrTaxaAnbima)  — indicativos publicados
# =============================================================================

import time
import sqlite3
import urllib.request

from config import ResolverTradesDb

TIMEOUT_SEG = 5      # espera máxima por lock do SQLite
CACHE_TTL_SEG = 300  # 5 min — a base é atualizada uma vez por dia

_cache = {}


class BancoIndisponivel(Exception):
    """trades.db não encontrado ou ilegível (mensagem já pronta pro Excel)."""


def _Uri(caminho):
    """Caminho do Windows → URI de leitura do SQLite (trata espaço e acento)."""
    return "file:" + urllib.request.pathname2url(caminho) + "?mode=ro"


def CaminhoBanco():
    """Caminho do trades.db em uso. Estoura BancoIndisponivel se não achar."""
    caminho = ResolverTradesDb()
    if not caminho:
        raise BancoIndisponivel(
            "trades.db não encontrado — defina a variável de ambiente TRADES_DB_PATH "
            "com o caminho completo do arquivo"
        )
    return caminho


def _Consultar(sql, parametros=()):
    """Roda uma consulta em modo leitura e devolve a lista de linhas."""
    caminho = CaminhoBanco()
    try:
        conexao = sqlite3.connect(_Uri(caminho), uri=True, timeout=TIMEOUT_SEG)
    except sqlite3.Error as e:
        raise BancoIndisponivel(f"não foi possível abrir o trades.db ({e})")
    try:
        return conexao.execute(sql, parametros).fetchall()
    except sqlite3.Error as e:
        raise BancoIndisponivel(f"consulta falhou no trades.db ({e})")
    finally:
        conexao.close()


def _Memo(chave, funcao):
    """Memoiza o resultado de `funcao` por CACHE_TTL_SEG (inclusive None)."""
    agora = time.time()
    guardado = _cache.get(chave)
    if guardado is not None and agora - guardado[0] < CACHE_TTL_SEG:
        return guardado[1]
    valor = funcao()
    _cache[chave] = (agora, valor)
    return valor


def LimparCache(forcarPath=True):
    """Esvazia o cache das consultas (e refaz a busca do arquivo)."""
    _cache.clear()
    if forcarPath:
        ResolverTradesDb(forcar=True)


def _Ticker(ticker):
    return str(ticker).strip().upper()


# ─── consultas ───────────────────────────────────────────────────────────────
# Nota de performance: `cdTicker` é PK em InfoAtivos e 1ª coluna da PK composta
# (cdTicker, dtReferencia) em AnbimaIndicativos, e a base guarda os tickers SEMPRE
# em maiúsculas. Por isso comparamos a coluna crua (`cdTicker = ?`) com o ticker
# já maiusculado no Python: assim o SQLite usa o índice (SEARCH). Envolver a
# coluna em `upper(cdTicker)` obrigaria um SCAN da tabela inteira a cada célula
# (medido: ~160× mais lento no histórico de 136k linhas). O `ORDER BY
# dtReferencia DESC` também sai de graça — é a 2ª coluna do índice (sem sort).

def ReferenciaAnbima(ticker):
    """cdReferencia do papel (vértice/título público de referência), ou None."""
    ticker = _Ticker(ticker)

    def _Buscar():
        linhas = _Consultar(
            "SELECT cdReferencia FROM InfoAtivos WHERE cdTicker = ? LIMIT 1",
            (ticker,),
        )
        if not linhas:
            return None
        valor = linhas[0][0]
        return valor.strip() if isinstance(valor, str) and valor.strip() else None

    return _Memo(("ref", ticker), _Buscar)


def IndicativoAnbima(ticker):
    """Indicativo ANBIMA mais recente do papel: {'data', 'taxa'} ou None.

    `taxa` vem como está publicada na base — em % a.a. (ex.: 7.5983)."""
    ticker = _Ticker(ticker)

    def _Buscar():
        linhas = _Consultar(
            "SELECT dtReferencia, vrTaxaAnbima FROM AnbimaIndicativos "
            "WHERE cdTicker = ? AND vrTaxaAnbima IS NOT NULL "
            "ORDER BY dtReferencia DESC LIMIT 1",
            (ticker,),
        )
        if not linhas:
            return None
        return {"data": linhas[0][0], "taxa": float(linhas[0][1])}

    return _Memo(("ind", ticker), _Buscar)


def HistoricoIndicativoAnbima(ticker):
    """Histórico de indicativos do papel, do mais recente para o mais antigo:
    lista de {'data', 'ticker', 'ref', 'taxa'} (vazia se não houver).

    A referência (`ref`) é a mesma em todas as linhas de um ticker, então em vez
    de um JOIN linha-a-linha com InfoAtivos, buscamos a `ref` uma única vez
    (reaproveitando o cache de ReferenciaAnbima) e uma consulta de tabela única,
    indexada, no AnbimaIndicativos. Resultado idêntico ao do JOIN, bem mais leve."""
    ticker = _Ticker(ticker)

    def _Buscar():
        linhas = _Consultar(
            "SELECT dtReferencia, vrTaxaAnbima FROM AnbimaIndicativos "
            "WHERE cdTicker = ? AND vrTaxaAnbima IS NOT NULL "
            "ORDER BY dtReferencia DESC",
            (ticker,),
        )
        ref = ReferenciaAnbima(ticker) or ""
        return [
            {"data": data, "ticker": ticker, "ref": ref, "taxa": float(taxa)}
            for data, taxa in linhas
        ]

    return _Memo(("hist", ticker), _Buscar)


def Diagnostico():
    """Uma linha sobre o estado da base, para o =cpTeste()/Sobre."""
    try:
        caminho = CaminhoBanco()
    except BancoIndisponivel as e:
        return f"trades.db: {e}"
    try:
        linhas = _Consultar("SELECT max(dtReferencia) FROM AnbimaIndicativos")
        return f"trades.db: OK (ANBIMA até {linhas[0][0] or 's/ dados'}) — {caminho}"
    except BancoIndisponivel as e:
        return f"trades.db: {e}"
