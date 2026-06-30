import os
import sys

# Garante que o diretório deste arquivo esteja no path antes de qualquer import
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)

from datetime import datetime, date as date_type

import xlwings as xw

# Captura qualquer erro de import para exibir na célula.
# O add-in usa SOMENTE as APIs (B3 / FI Analytics) — sem calculadora local e
# sem leitura de qualquer base nossa (trades.db).
_IMPORT_ERROR = None
try:
    from apis import PrecoB3, TaxaB3, PrecoFi, TaxaFi, LimparCache
except Exception as _e:
    _IMPORT_ERROR = str(_e)


def _parse_data(val):
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date_type):
        return val
    return datetime.strptime(str(val).strip(), "%d/%m/%Y").date()


def _pct(taxa_excel: float) -> float:
    return float(taxa_excel) * 100


def _data_iso(data) -> str:
    """Converte o argumento de data do Excel para 'YYYY-MM-DD' (formato das APIs)."""
    return _parse_data(data).strftime("%Y-%m-%d")


def main():
    """Handler do botão 'Run' do ribbon (opcional). As UDFs =PU/=DUR/=TAXA
    funcionam de forma independente — este botão só dá um retorno visual."""
    import xlwings as xw
    xw.apps.active.selection.cells(1, 1).value = "CalcRF carregado — use =PU, =DUR, =TAXA"


# =============================================================================
# UDFs — todas puxam exclusivamente das APIs (B3 → FI Analytics)
# =============================================================================

@xw.func
def TESTE():
    """Diagnóstico: retorna OK ou o erro de import."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    return f"OK — path: {_THIS_DIR}"


@xw.func(async_mode='threading')
@xw.arg('taxa', numbers=float)
def PU(ticker, data, taxa):
    """PU de Operação. taxa como % no Excel (ex: 6,4618%).

    Fonte: B3 → FI Analytics (nesta ordem). Sem fallback local.
    """
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker  = str(ticker).upper().strip()
        dataIso = _data_iso(data)
        taxaPct = _pct(taxa)

        resp = PrecoB3(ticker, dataIso, taxaPct) or PrecoFi(ticker, dataIso, taxaPct)
        if resp and resp.get("pu") is not None:
            return float(resp["pu"])
        return "ERRO: APIs sem resposta (B3/FI)"
    except Exception as e:
        return f"ERRO: {e}"


@xw.func(async_mode='threading')
@xw.arg('taxa', numbers=float)
def DUR(ticker, data, taxa):
    """Duration Macaulay em anos. taxa como % no Excel.

    Fonte: B3 → FI Analytics (nesta ordem). Sem fallback local.
    """
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker  = str(ticker).upper().strip()
        dataIso = _data_iso(data)
        taxaPct = _pct(taxa)

        resp = PrecoB3(ticker, dataIso, taxaPct) or PrecoFi(ticker, dataIso, taxaPct)
        if resp and resp.get("duration") is not None:
            return float(resp["duration"])
        return "ERRO: APIs sem resposta (B3/FI)"
    except Exception as e:
        return f"ERRO: {e}"


@xw.func(async_mode='threading')
@xw.arg('pu', numbers=float)
def TAXA(ticker, data, pu):
    """Taxa de negociação dado PU Op (retorna decimal para formatar como %).

    Fonte: B3 → FI Analytics (nesta ordem). Sem fallback local.
    """
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker  = str(ticker).upper().strip()
        dataIso = _data_iso(data)
        puFloat = float(pu)

        taxa = TaxaB3(ticker, dataIso, puFloat)
        if taxa is None:
            taxa = TaxaFi(ticker, dataIso, puFloat)
        if taxa is not None:
            return float(taxa) / 100
        return "ERRO: APIs sem resposta (B3/FI)"
    except Exception as e:
        return f"ERRO: {e}"


@xw.func
def LIMPARCACHE():
    """Esvazia o cache de respostas das APIs (B3/FI) e força novas chamadas."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    LimparCache()
    return "cache limpo"
