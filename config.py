# =============================================================================
# config.py — configuração mínima do add-in (paths)
# -----------------------------------------------------------------------------
# Os segredos (token B3, API key FI) e o proxy vêm de VARIÁVEIS DE AMBIENTE.
# Ver apis.py / CLAUDE.md para os nomes das variáveis.
#
# ENV_PATH é apenas um fallback de desenvolvimento: caminho de um .env opcional
# na própria pasta. No banco esse arquivo NÃO existe e todos os valores vêm das
# variáveis de ambiente (apis.py trata a ausência sem erro).
#
# ResolverTradesDb() descobre o trades.db (base do projeto de negociação
# secundária) — usado SÓ pelas fórmulas ANBIMA, em leitura. Ver basedados.py.
# =============================================================================

import os

# Fallback de dev: por padrão procura um .env na própria pasta (não versionado).
ENV_PATH = os.path.join(os.path.dirname(__file__), ".env")

# Override de desenvolvimento LOCAL (fora do versionamento): para apontar a um
# .env FORA da pasta do repo, crie um config_local.py (gitignored) com:
#     ENV_PATH = r"C:\caminho\para\.env"
# No banco isso não existe e os segredos vêm das variáveis de ambiente.
try:
    from config_local import ENV_PATH  # noqa: F811
except Exception:
    pass


# =============================================================================
# trades.db — base do projeto de negociação secundária (SOMENTE LEITURA)
# -----------------------------------------------------------------------------
# O add-in não carrega esse caminho embutido: ele é DESCOBERTO em tempo de
# execução. A pasta do add-in (Z:\CP em produção, a pasta local em dev) tem a
# pasta do projeto de negociação como "irmã" — em algum ponto acima dela:
#
#     <raiz>\<pasta do projeto>\NegociacaoSecundario\code\data\trades.db
#     <raiz>\negociacao-secundario\code\data\trades.db          (layout de dev)
#
# O segmento intermediário varia por ambiente (e é nominal), por isso ele NÃO é
# escrito aqui — o repo é público: descobrimos varrendo UM nível de subpastas da
# raiz. Quem quiser fugir disso seta TRADES_DB_PATH e nada é varrido.
# =============================================================================

# Override explícito: caminho COMPLETO do arquivo .db. Tem prioridade sobre tudo.
ENV_TRADES_DB = "TRADES_DB_PATH"

TRADES_DB_NOME = "trades.db"

# Sufixo relativo, a partir da pasta do projeto de negociação secundária.
_SUFIXO_DB = os.path.join("code", "data", TRADES_DB_NOME)

# Nomes possíveis da pasta do projeto (produção e layout de dev).
_PASTAS_PROJETO = ("NegociacaoSecundario", "negociacao-secundario")

# Teto da varredura: nº de subpastas inspecionadas por raiz (evita travar numa
# share enorme). O resultado é memoizado, então isso roda no máximo uma vez.
_MAX_SUBPASTAS = 200

_cacheTradesDb = None   # None = ainda não resolvido; "" = procurado e não achado


def _RaizesCandidatas():
    """Pastas onde procurar a árvore do projeto de negociação secundária:
    a pasta do add-in, a apontada por CALCCP_DIR, e as pastas-mãe das duas."""
    raizes = []
    for base in (os.environ.get("CALCCP_DIR"), os.path.dirname(os.path.abspath(__file__))):
        if not base:
            continue
        base = os.path.abspath(os.path.expandvars(base.strip()))
        for cand in (base, os.path.dirname(base)):
            if cand and cand not in raizes:
                raizes.append(cand)
    return raizes


def _TentarRaiz(raiz):
    """Procura o trades.db sob `raiz`: primeiro o join direto, depois um nível
    de subpastas. Devolve o caminho ou None."""
    for pasta in _PASTAS_PROJETO:
        caminho = os.path.join(raiz, pasta, _SUFIXO_DB)
        if os.path.isfile(caminho):
            return caminho
    try:
        with os.scandir(raiz) as entradas:
            for i, entrada in enumerate(entradas):
                if i >= _MAX_SUBPASTAS:
                    break
                if not entrada.is_dir():
                    continue
                for pasta in _PASTAS_PROJETO:
                    caminho = os.path.join(entrada.path, pasta, _SUFIXO_DB)
                    if os.path.isfile(caminho):
                        return caminho
    except OSError:
        pass  # raiz inacessível (drive não mapeado, sem permissão) — segue adiante
    return None


def ResolverTradesDb(forcar=False):
    """Caminho do trades.db, ou None se não for encontrado.

    Ordem: variável de ambiente TRADES_DB_PATH → pasta do add-in / CALCCP_DIR e
    suas pastas-mãe (join direto e depois um nível de subpastas).
    O resultado é memoizado por processo (inclusive a ausência); `forcar=True`
    refaz a busca (usado pelo =cpLimparCache)."""
    global _cacheTradesDb
    if forcar:
        _cacheTradesDb = None
    if _cacheTradesDb is not None:
        return _cacheTradesDb or None

    override = os.environ.get(ENV_TRADES_DB, "").strip()
    if override:
        override = os.path.abspath(os.path.expandvars(override))
        if os.path.isfile(override):
            _cacheTradesDb = override
            return override
        # Setada mas inválida: não silencia — sem fallback, o erro fica explícito.
        _cacheTradesDb = ""
        return None

    for raiz in _RaizesCandidatas():
        achado = _TentarRaiz(raiz)
        if achado:
            _cacheTradesDb = achado
            return achado

    _cacheTradesDb = ""
    return None
