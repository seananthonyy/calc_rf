import os
import sys

# Garante que o diretório deste arquivo esteja no path antes de qualquer import
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)

from datetime import datetime, date as date_type, timedelta

import xlwings as xw

# Epoch do serial de datas do Excel (Windows): base efetiva 30/12/1899.
_EXCEL_EPOCH = date_type(1899, 12, 30)

# Versão da lógica (.py). O .xlam é versionado por NOME de arquivo (CalcCP_vN.xlam);
# a lógica aqui é retrocompatível (só adiciona UDF, nunca remove/renomeia) — ver CHANGELOG.md.
VERSION = "4.0.0"

# Todas as UDFs têm o prefixo cp (ex.: =cpPu, =cpTaxa, =cpVna...).
# Títulos com API (debênture, CRI/CRA, NTN-B…): via B3 → FI Analytics → bondbuilder.
# DI (tickers DI1...): sem API → calculado LOCALMENTE em di.py.
# Séries macro (SELIC, CDI, IPCA, dólar…): API pública do Banco Central (SGS).
_IMPORT_ERROR = None
try:
    from apis import (
        Preco, TaxaOp, LimparCache, FatorDi,
        CampoBond, CampoFi, FluxoCadastrado, BcbValor, IpcaIndice,
    )
    import di
except Exception as _e:
    _IMPORT_ERROR = str(_e)

# Fórmulas ANBIMA: leem o trades.db (negociação secundária) em SOMENTE LEITURA.
# Import separado de propósito — se a base/módulo faltar, só as fórmulas ANBIMA
# param; o resto do add-in (que é API) continua funcionando.
_DB_ERROR = None
try:
    import basedados
except Exception as _e:
    _DB_ERROR = str(_e)


# ─── helpers ─────────────────────────────────────────────────────────────────

def _parse_data(val):
    """Converte o argumento de data para date. Aceita datetime/date, serial do
    Excel (ex.: TODAY() aninhado) ou string 'dd/mm/aaaa'."""
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date_type):
        return val
    if isinstance(val, bool):
        raise ValueError(f"data inválida: {val!r}")
    if isinstance(val, (int, float)):
        return _EXCEL_EPOCH + timedelta(days=int(val))
    return datetime.strptime(str(val).strip(), "%d/%m/%Y").date()


def _pct(taxa_excel) -> float:
    """% do Excel → número (6,4618% → 6.4618)."""
    return float(taxa_excel) * 100


def _data_iso(data) -> str:
    return _parse_data(data).strftime("%Y-%m-%d")


def _data_saida(iso):
    """'YYYY-MM-DD' → date (para o Excel exibir como data). Passa o resto adiante."""
    if isinstance(iso, str) and len(iso) == 10 and iso[4] == "-":
        try:
            return date_type.fromisoformat(iso)
        except ValueError:
            return iso
    return iso


def _vazio(v) -> bool:
    return v is None or v == "" or (isinstance(v, float) and v != v)  # None/""/NaN


def _resolve_taxa(ticker, taxa) -> float:
    """Taxa em % a.a. para chamar a API. Se omitida, usa a taxa de EMISSÃO do papel
    (pupar/VNA/fluxo independem da taxa negociada, então isso é seguro)."""
    if _vazio(taxa):
        y = CampoBond(ticker, "yield")
        if y is None:
            raise ValueError("informe a taxa (a taxa de emissão não foi encontrada na B3)")
        return float(y)
    return _pct(taxa)


def _erro(pref, e):
    return f"ERRO: {e}"


def main():
    """Handler do botão 'Run' do ribbon (opcional)."""
    xw.apps.active.selection.cells(1, 1).value = (
        "CalcCP — use =cpPu, =cpTaxa, =cpVna, =cpFluxo, =cpSelic… (prefixo cp)"
    )


# =============================================================================
# DIAGNÓSTICO
# =============================================================================

def _diag_banco():
    """Uma linha sobre o trades.db (fórmulas ANBIMA) para o diagnóstico."""
    if _DB_ERROR:
        return f"trades.db: ERRO import ({_DB_ERROR})"
    try:
        return basedados.Diagnostico()
    except Exception as e:
        return f"trades.db: ERRO ({e})"


@xw.func
def cpTeste():
    """Diagnóstico: versão, pasta, origem do path e estado do trades.db."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    env = os.environ.get("CALCCP_DIR")
    origem = f"CALCCP_DIR={env}" if env else "CALCCP_DIR não setada (fallback)"
    return f"OK v{VERSION} — path: {_THIS_DIR} — {origem} — {_diag_banco()}"


@xw.func
def cpLimparCache():
    """Esvazia o cache das APIs (B3/FI/BCB) e o das consultas ao trades.db."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    LimparCache()
    if not _DB_ERROR:
        basedados.LimparCache()
    return "cache limpo"


def Sobre():
    """Caixa 'Sobre' do botão do ribbon (chamada por RunPython, não é UDF).

    Mostra a versão da LÓGICA (deste .py) — que é o que muda quando a pasta da share é
    atualizada —, a pasta de onde o add-in está lendo e o estado do import. Como roda o
    Python de verdade, se a ponte estiver quebrada o erro aparece aqui."""
    import ctypes
    udfs = sorted(n for n, v in globals().items() if n.startswith("cp") and callable(v))
    estado = f"ERRO no import: {_IMPORT_ERROR}" if _IMPORT_ERROR else "APIs carregadas (B3 / FI / BCB / IBGE)"
    env = os.environ.get("CALCCP_DIR")
    origem = f"CALCCP_DIR = {env}" if env else "CALCCP_DIR não setada (usando fallback do .xlam)"
    texto = (
        f"CalcCP — Renda Fixa no Excel\n\n"
        f"Versão da lógica:  {VERSION}\n"
        f"Fórmulas:          {len(udfs)} (todas com prefixo cp)\n"
        f"Estado:            {estado}\n"
        f"Origem do path:    {origem}\n\n"
        f"{_diag_banco()}\n\n"
        f"Pasta:\n{_THIS_DIR}\n\n"
        f"Python:\n{sys.executable}\n\n"
        f"Guia de uso: abra o COMO_USAR.html que está na pasta acima."
    )
    ctypes.windll.user32.MessageBoxW(0, texto, "CalcCP", 0x40)  # 0x40 = ícone de informação


# =============================================================================
# PRECIFICAÇÃO — PU, Taxa, Duration (B3 → FI → bondbuilder; DI local)
# =============================================================================

@xw.func
@xw.arg('taxa', numbers=float)
def cpPu(ticker, data, taxa):
    """PU de Operação. taxa como % do Excel (ex.: 6,4618%). Fonte: B3 → FI."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        if di.EhTickerDi(ticker):
            return float(di.PuDi(_pct(taxa), _parse_data(data), ticker))
        r = Preco(ticker, _data_iso(data), _pct(taxa))
        if r and r.get("pu") is not None:
            return float(r["pu"])
        return "ERRO: APIs sem resposta (B3/FI)"
    except Exception as e:
        return _erro("pu", e)


@xw.func
@xw.arg('pu', numbers=float)
def cpTaxa(ticker, data, pu):
    """Taxa de negociação dado o PU (retorna decimal — formate a célula como %)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        if di.EhTickerDi(ticker):
            return float(di.TaxaDi(float(pu), _parse_data(data), ticker)) / 100
        t = TaxaOp(ticker, _data_iso(data), float(pu))
        if t is not None:
            return float(t) / 100
        return "ERRO: APIs sem resposta (B3/FI)"
    except Exception as e:
        return _erro("taxa", e)


@xw.func
@xw.arg('taxa', numbers=float)
def cpDur(ticker, data, taxa):
    """Duration de Macaulay em ANOS. Fonte: B3 → FI."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        if di.EhTickerDi(ticker):
            return float(di.DurationDi(_parse_data(data), ticker))
        r = Preco(ticker, _data_iso(data), _pct(taxa))
        if r and r.get("duration") is not None:
            return float(r["duration"])
        return "ERRO: APIs sem resposta (B3/FI)"
    except Exception as e:
        return _erro("dur", e)


# =============================================================================
# DADOS DO PAPEL — PU Par, VNA, fluxo restante, datas, VNE (taxa opcional)
# =============================================================================

@xw.func
@xw.arg('taxa', numbers=float)
def cpPupar(ticker, data, taxa=None):
    """PU Par (valor nominal atualizado + juros acumulados) na data. taxa opcional
    (pupar independe da taxa; se omitida usa a de emissão)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        r = Preco(ticker, _data_iso(data), _resolve_taxa(ticker, taxa))
        if r and r.get("pupar") is not None:
            return float(r["pupar"])
        return "ERRO: PU Par indisponível (B3/FI)"
    except Exception as e:
        return _erro("pupar", e)


@xw.func
@xw.arg('taxa', numbers=float)
def cpVna(ticker, data, taxa=None):
    """VNA — Valor Nominal Atualizado na data. taxa opcional (VNA independe da taxa)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        r = Preco(ticker, _data_iso(data), _resolve_taxa(ticker, taxa))
        if r and r.get("vna") is not None:
            return float(r["vna"])
        return "ERRO: VNA indisponível (B3/FI)"
    except Exception as e:
        return _erro("vna", e)


@xw.func
@xw.arg('taxa', numbers=float)
def cpFluxo(ticker, data, taxa=None):
    """Agenda RESTANTE do papel a partir da data (spill): Data | Tipo | %Amort | %Incorp.
    Uma linha por data de evento (a data de amortização, que a B3 manda em 2 eventos, vira uma
    linha só). Tipo: J = só juros, J+A = juros e amortização, J+I = juros e incorporação
    (incorporação só em papel IPCA-I). Datas ajustadas p/ dia útil (feriados ANBIMA).
    `taxa` é aceita por compatibilidade e ignorada (a agenda não depende dela).
    (=cpFluxoCompleto traz a agenda INTEIRA, desde a emissão.)"""
    if _IMPORT_ERROR:
        return [[f"ERRO import: {_IMPORT_ERROR}"]]
    try:
        ag = FluxoCadastrado(str(ticker).upper().strip())
        if not ag:
            return [["ERRO: agenda indisponível (B3 getBondDetails)"]]
        corte = _data_iso(data)
        linhas = [["Data", "Tipo", "%Amort", "%Incorp"]]
        for e in ag:
            iso = _ajusta_du_iso(e["data"])
            if iso < corte:
                continue
            linhas.append([_data_saida(iso), e["tipo"], e["amort"], e["incorp"]])
        if len(linhas) == 1:
            return [["ERRO: nenhum evento restante (papel vencido?)"]]
        return linhas
    except Exception as e:
        return [[f"ERRO: {e}"]]


def _ajusta_du_iso(iso):
    """Ajusta uma data ISO 'AAAA-MM-DD' para o próximo dia útil (ANBIMA)."""
    try:
        d = date_type.fromisoformat(str(iso)[:10])
        return di.ProximoDu(d).isoformat()
    except Exception:
        return iso


@xw.func
def cpFluxoCompleto(ticker):
    """Agenda CADASTRADA do papel (B3 getBondDetails, spill): Data | %Amort | %Incorp.
    UMA linha por data de evento, da emissão ao vencimento — datas de cupom puro entram com
    0 em %Amort e 0 em %Incorp. Datas ajustadas p/ dia útil (feriados ANBIMA). Incorporação só
    aparece em papéis IPCA-I. (=cpFluxo é o fluxo de caixa RESTANTE calculado.)"""
    if _IMPORT_ERROR:
        return [[f"ERRO import: {_IMPORT_ERROR}"]]
    try:
        ag = FluxoCadastrado(str(ticker).upper().strip())
        if not ag:
            return [["ERRO: agenda indisponível (B3 getBondDetails)"]]
        linhas = [["Data", "%Amort", "%Incorp"]]
        for e in ag:
            linhas.append([
                _data_saida(_ajusta_du_iso(e["data"])), e["amort"], e["incorp"],
            ])
        return linhas
    except Exception as e:
        return [[f"ERRO: {e}"]]


@xw.func
def cpVencimento(ticker):
    """Data de vencimento do papel (B3 getBondDetails)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        v = CampoBond(str(ticker).upper().strip(), "expiredate")
        return _data_saida(v) if v else "ERRO: vencimento indisponível"
    except Exception as e:
        return _erro("venc", e)


@xw.func
def cpEmissao(ticker):
    """Data de emissão do papel (B3 getBondDetails)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        v = CampoBond(str(ticker).upper().strip(), "issuedate")
        return _data_saida(v) if v else "ERRO: emissão indisponível"
    except Exception as e:
        return _erro("emissao", e)


@xw.func
def cpInicioRentabilidade(ticker):
    """Data de início de rentabilidade (B3 getBondDetails)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        v = CampoBond(str(ticker).upper().strip(), "startingdate")
        return _data_saida(v) if v else "ERRO: início indisponível"
    except Exception as e:
        return _erro("inicio", e)


@xw.func
def cpTaxaEmissao(ticker):
    """Taxa de emissão do papel em DECIMAL (ex.: 6,4618% → 0,064618; %DI 113,5% → 1,135)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        v = CampoBond(str(ticker).upper().strip(), "yield")
        return float(v) / 100 if v is not None else "ERRO: taxa de emissão indisponível"
    except Exception as e:
        return _erro("taxaEmi", e)


@xw.func
def cpVne(ticker):
    """Valor Nominal de Emissão (B3 getBondDetails)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        v = CampoBond(str(ticker).upper().strip(), "vne")
        return float(v) if v is not None else "ERRO: VNE indisponível"
    except Exception as e:
        return _erro("vne", e)


@xw.func
def cpAniversario(ticker):
    """Dia de aniversário do papel (B3 getBondDetails)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        v = CampoBond(str(ticker).upper().strip(), "anniversaryday")
        return float(v) if v is not None else "ERRO: aniversário indisponível"
    except Exception as e:
        return _erro("aniv", e)


# =============================================================================
# ANBIMA — referência e taxas indicativas (trades.db, SOMENTE LEITURA)
# -----------------------------------------------------------------------------
# Únicas fórmulas que NÃO vêm de API: esses dados só existem na base do projeto
# de negociação secundária. O caminho do arquivo é descoberto em tempo de
# execução (config.ResolverTradesDb) — ou fixado na variável TRADES_DB_PATH.
#
# Convenção de retorno:
#   "#N/A"   → o papel/dado não existe na base (resposta legítima);
#   "ERRO: " → problema de infraestrutura (base não encontrada, ilegível).
# As taxas saem em DECIMAL (7,5983% → 0,075983), como nas demais fórmulas de
# taxa do add-in (=cpTaxa, =cpTaxaEmissao, =cpSelic) — formate a célula como %.
# =============================================================================

_NA = "#N/A"


def _taxa_decimal(taxa):
    """% a.a. publicado → decimal, sem ruído de float (7,6411 → 0,076411)."""
    return float(f"{float(taxa) / 100:.10g}")


def _db_indisponivel():
    """Mensagem de erro se a camada de base não puder ser usada; senão None."""
    if _DB_ERROR:
        return f"ERRO import: {_DB_ERROR}"
    return None


@xw.func
def cpAnbimaRef(ticker):
    """Vértice/título público de REFERÊNCIA cadastrado para o papel
    (ex.: "NTN-B 35", "DI1F29"). Fonte: trades.db (InfoAtivos.cdReferencia).
    Devolve #N/A se o papel não estiver na base ou não tiver referência."""
    erro = _db_indisponivel()
    if erro:
        return erro
    try:
        ref = basedados.ReferenciaAnbima(ticker)
        return ref if ref else _NA
    except basedados.BancoIndisponivel as e:
        return f"ERRO: {e}"
    except Exception as e:
        return _erro("anbimaRef", e)


@xw.func
def cpAnbimaIndicativo(ticker):
    """Taxa indicativa ANBIMA mais RECENTE disponível na base para o papel, em
    DECIMAL (formate a célula como %). Fonte: trades.db (AnbimaIndicativos).
    Devolve #N/A se o papel não tiver indicativo publicado.
    (O histórico com as datas está em =cpAnbimaIndicativoHistorico.)"""
    erro = _db_indisponivel()
    if erro:
        return erro
    try:
        ind = basedados.IndicativoAnbima(ticker)
        return _taxa_decimal(ind["taxa"]) if ind else _NA
    except basedados.BancoIndisponivel as e:
        return f"ERRO: {e}"
    except Exception as e:
        return _erro("anbimaInd", e)


@xw.func
def cpAnbimaSpread(ticker):
    """Spread ANBIMA mais RECENTE disponível na base para o papel, sobre a
    referência de =cpAnbimaRef, em DECIMAL (formate a célula como %; -0,5193%
    = -51,93 bps). Fonte: trades.db (AnbimaIndicativos.vrSpreadAnbima).
    Quando a referência é FUNDING (CDI+/%CDI) o spread É a própria taxa
    indicativa — não há benchmark externo para descontar.
    Devolve #N/A se o papel não tiver spread calculado (é mais esparso que a
    taxa: exige referência cadastrada e curva de MtM na data).
    (O histórico com as datas está em =cpAnbimaIndicativoHistorico.)"""
    erro = _db_indisponivel()
    if erro:
        return erro
    try:
        spr = basedados.SpreadAnbima(ticker)
        return _taxa_decimal(spr["spread"]) if spr else _NA
    except basedados.BancoIndisponivel as e:
        return f"ERRO: {e}"
    except Exception as e:
        return _erro("anbimaSpread", e)


@xw.func
def cpAnbimaIndicativoHistorico(ticker):
    """Histórico COMPLETO de indicativos ANBIMA do papel (spill), do mais recente
    para o mais antigo: Data | Ticker | Ref | Taxa Indicativa | Spread.
    Taxa e spread em DECIMAL (formate as colunas como %). Fonte: trades.db.
    A coluna Spread sai #N/A nas datas sem spread calculado."""
    erro = _db_indisponivel()
    if erro:
        return erro
    try:
        linhas = basedados.HistoricoIndicativoAnbima(ticker)
        if not linhas:
            return _NA
        saida = [["Data", "Ticker", "Ref", "Taxa Indicativa", "Spread"]]
        for l in linhas:
            spread = l["spread"]
            saida.append([
                _data_saida(l["data"]), l["ticker"], l["ref"], _taxa_decimal(l["taxa"]),
                _NA if spread is None else _taxa_decimal(spread),
            ])
        return saida
    except basedados.BancoIndisponivel as e:
        return f"ERRO: {e}"
    except Exception as e:
        return f"ERRO: {e}"


# =============================================================================
# MÉTRICAS / GROSS UP (FI Analytics)
# =============================================================================

@xw.func
@xw.arg('taxa', numbers=float)
def cpGrossUp(ticker, data, taxa=None):
    """Gross up / taxa equivalente após imposto (FI 'taxedM2MRate'). PRECISA da taxa
    (o gross up depende da taxa negociada). Veja o tipo em =cpGrossUpTipo."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        if _vazio(taxa):
            return "ERRO: informe a taxa"
        ticker = str(ticker).upper().strip()
        v = CampoFi(ticker, _data_iso(data), _pct(taxa), "taxedM2MRate")
        return float(v) if v is not None else "ERRO: gross up indisponível (FI)"
    except Exception as e:
        return _erro("grossup", e)


@xw.func
@xw.arg('taxa', numbers=float)
def cpGrossUpTipo(ticker, data, taxa=None):
    """Tipo do ajuste fiscal: 'GROSS_UP' (papel isento) ou 'AFTER_TAX' (tributado)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        v = CampoFi(ticker, _data_iso(data), _resolve_taxa(ticker, taxa), "taxedType")
        return v if v is not None else "ERRO: indisponível (FI)"
    except Exception as e:
        return _erro("grossupTipo", e)


@xw.func
@xw.arg('taxa', numbers=float)
def cpDv01(ticker, data, taxa=None):
    """DV01 — variação do PU por 1 bp na taxa (FI Analytics). PRECISA da taxa."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        if _vazio(taxa):
            return "ERRO: informe a taxa"
        ticker = str(ticker).upper().strip()
        v = CampoFi(ticker, _data_iso(data), _pct(taxa), "dv01")
        return float(v) if v is not None else "ERRO: indisponível (FI)"
    except Exception as e:
        return _erro("dv01", e)


# =============================================================================
# CDI (endpoint público B3 /di/calculo)
# =============================================================================

@xw.func
def cpCdi(dataInicio, dataFim, percentual=100):
    """Fator de CDI acumulado entre duas datas (ex.: =cpCdi("02/01/2024";"30/12/2024";100)).
    percentual = número puro (100 = 100% do CDI). Multiplique pelo valor base para capitalizar."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        if _vazio(percentual):
            percentual = 100.0
        f = FatorDi(_data_iso(dataInicio), _data_iso(dataFim), float(percentual))
        return float(f) if f is not None else "ERRO: API DI sem resposta"
    except Exception as e:
        return _erro("cdi", e)


# =============================================================================
# BANCO CENTRAL — séries macro (SGS, público). Valor mais recente por padrão.
# =============================================================================

def _bcb_valor(serie, data=None):
    if _vazio(data):
        return "ERRO: informe a data"
    v = BcbValor(serie, _data_br(data))
    return float(v) if v is not None else "ERRO: BCB sem dado"


def _bcb_taxa(serie, data=None):
    """Como _bcb_valor, mas devolve a TAXA em decimal (6% -> 0.06). Erros passam adiante."""
    v = _bcb_valor(serie, data)
    return v / 100 if isinstance(v, (int, float)) else v


def _data_br(data):
    return _parse_data(data).strftime("%d/%m/%Y")


# Todas aceitam [data] opcional (dd/mm/aaaa). Semântica as-of: sem publicação no
# dia exato (fim de semana/feriado ou série mensal), retorna o último valor ATÉ a data.
@xw.func
def cpSelic(data=None):
    """Meta SELIC (DECIMAL, ex.: 0,02 = 2% a.a.) — BCB série 432. data OBRIGATÓRIA (as-of: último valor até a data)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_taxa(432, data)
    except Exception as e:
        return _erro("selic", e)


@xw.func
def cpCdiAno(data=None):
    """CDI anualizado base 252 (DECIMAL) — BCB série 4389 (data obrigatória, as-of)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_taxa(4389, data)
    except Exception as e:
        return _erro("cdiano", e)


@xw.func
def cpIpca(data=None):
    """IPCA do mês (DECIMAL, ex.: 0,0093 = 0,93%) — BCB série 433. [data] opcional (mês de referência)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_taxa(433, data)
    except Exception as e:
        return _erro("ipca", e)


@xw.func
def cpIpcaIndice(data=None):
    """IPCA número-índice (base dez/1993 = 100) do mês de referência de [data];
    sem data = último mês. Fonte: IBGE SIDRA. A razão de dois números-índice
    (=cpIpcaIndice(fim)/cpIpcaIndice(ini)) é o fator de correção do IPCA entre as datas."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        if _vazio(data):
            return "ERRO: informe a data"
        v = IpcaIndice(_data_br(data))
        return float(v) if v is not None else "ERRO: SIDRA sem dado"
    except Exception as e:
        return _erro("ipcaindice", e)


@xw.func
def cpIgpm(data=None):
    """IGP-M do mês (DECIMAL) — BCB série 189 (data obrigatória, as-of)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_taxa(189, data)
    except Exception as e:
        return _erro("igpm", e)


@xw.func
def cpIpca15(data=None):
    """IPCA-15 do mês (DECIMAL) — BCB série 7478 (data obrigatória, as-of)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_taxa(7478, data)
    except Exception as e:
        return _erro("ipca15", e)


@xw.func
def cpDolar(data=None):
    """Dólar PTAX venda (R$) — BCB série 1. data OBRIGATÓRIA (as-of: último valor até a data)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(1, data)
    except Exception as e:
        return _erro("dolar", e)


@xw.func
def cpEuro(data=None):
    """Euro venda (R$) — BCB série 21619. data OBRIGATÓRIA (as-of: último valor até a data)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(21619, data)
    except Exception as e:
        return _erro("euro", e)


# =============================================================================
# DIAS ÚTEIS — feriados ANBIMA (via di.py)
# =============================================================================

@xw.func
def cpEhDiaUtil(data):
    """VERDADEIRO se `data` é dia útil (feriados ANBIMA)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return bool(di.EhDu(_parse_data(data)))
    except Exception as e:
        return _erro("ehdu", e)


@xw.func
def cpDiasUteis(dataInicio, dataFim):
    """Número de dias úteis entre dataInicio (inclusive) e dataFim (exclusive), feriados ANBIMA."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return int(di.ContarDu(_parse_data(dataInicio), _parse_data(dataFim)))
    except Exception as e:
        return _erro("dias", e)


@xw.func
def cpDiaUtilPosterior(data):
    """Próximo dia útil ESTRITAMENTE após `data` (feriados ANBIMA)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        d = _parse_data(data) + timedelta(days=1)
        while not di.EhDu(d):
            d += timedelta(days=1)
        return d
    except Exception as e:
        return _erro("dupos", e)


@xw.func
def cpDiaUtilAnterior(data):
    """Dia útil ESTRITAMENTE anterior a `data` (feriados ANBIMA)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        d = _parse_data(data) - timedelta(days=1)
        while not di.EhDu(d):
            d -= timedelta(days=1)
        return d
    except Exception as e:
        return _erro("duant", e)


@xw.func
def cpDiaUtilMaisN(data, n):
    """`data` deslocada em `n` dias úteis (n negativo anda para trás). Feriados ANBIMA."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        d = _parse_data(data)
        n = int(n)
        passo = 1 if n >= 0 else -1
        while n != 0:
            d += timedelta(days=passo)
            if di.EhDu(d):
                n -= passo
        return d
    except Exception as e:
        return _erro("dumaisn", e)
