import os
import sys

# Garante que o diretório deste arquivo esteja no path antes de qualquer import
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)

from datetime import datetime, date as date_type, timedelta

import xlwings as xw

# Epoch do serial de datas do Excel (Windows): dia 1 = 01/01/1900, mas com o bug
# do ano-bissexto de 1900 a base efetiva é 30/12/1899.
_EXCEL_EPOCH = date_type(1899, 12, 30)

# Captura qualquer erro de import para exibir na célula.
# Títulos com API (debênture, CRI/CRA, NTN-B…): via B3 → FI Analytics.
# DI (tickers DI1...): NÃO há API → calculado LOCALMENTE em di.py.
_IMPORT_ERROR = None
try:
    from apis import Preco, TaxaOp, LimparCache, FatorDi
    import di
except Exception as _e:
    _IMPORT_ERROR = str(_e)


def _parse_data(val):
    """Converte o argumento de data para date. Aceita:
      - datetime/date (xlwings converte célula formatada como data);
      - número serial do Excel (ex.: TODAY() aninhado direto: =PU(...; TODAY(); ...));
      - string 'dd/mm/aaaa'.
    """
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date_type):
        return val
    if isinstance(val, bool):  # bool é subclasse de int — barra antes do numérico
        raise ValueError(f"data inválida: {val!r}")
    if isinstance(val, (int, float)):
        return _EXCEL_EPOCH + timedelta(days=int(val))
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
    xw.apps.active.selection.cells(1, 1).value = "AntonioOliveiraCalc carregado — use =PU, =DUR, =TAXA, =CDI"


# =============================================================================
# UDFs — todas puxam exclusivamente das APIs (B3 → FI Analytics)
# =============================================================================

@xw.func
def TESTE():
    """Diagnóstico: retorna OK ou o erro de import."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    return f"OK — path: {_THIS_DIR}"


@xw.func
@xw.arg('taxa', numbers=float)
def PU(ticker, data, taxa):
    """PU de Operação. taxa como % no Excel (ex: 6,4618%).

    Fonte: B3 → FI Analytics (nesta ordem). Sem fallback local.
    Síncrona (não-async) de propósito: async fazia write-back que disparava recálculo
    em loop quando o argumento vinha de fórmula viva (XLOOKUP) em planilha do SharePoint.
    O cache de respostas em apis.py evita travar (só bate na rede na 1ª vez por input).
    """
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker  = str(ticker).upper().strip()
        dataIso = _data_iso(data)
        taxaPct = _pct(taxa)

        if di.EhTickerDi(ticker):
            return float(di.PuDi(taxaPct, _parse_data(data), ticker))

        resp = Preco(ticker, dataIso, taxaPct)
        if resp and resp.get("pu") is not None:
            return float(resp["pu"])
        return "ERRO: APIs sem resposta (B3/FI)"
    except Exception as e:
        return f"ERRO: {e}"


@xw.func
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

        if di.EhTickerDi(ticker):
            return float(di.DurationDi(_parse_data(data), ticker))

        resp = Preco(ticker, dataIso, taxaPct)
        if resp and resp.get("duration") is not None:
            return float(resp["duration"])
        return "ERRO: APIs sem resposta (B3/FI)"
    except Exception as e:
        return f"ERRO: {e}"


@xw.func
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

        if di.EhTickerDi(ticker):
            return float(di.TaxaDi(puFloat, _parse_data(data), ticker)) / 100

        taxa = TaxaOp(ticker, dataIso, puFloat)
        if taxa is not None:
            return float(taxa) / 100
        return "ERRO: APIs sem resposta (B3/FI)"
    except Exception as e:
        return f"ERRO: {e}"


@xw.func
def CDI(dataInicio, dataFim, percentual=100):
    """Fator de CDI acumulado entre duas datas (endpoint público B3 /di/calculo).

    Ex.: =CDI("02/01/2024"; "30/12/2024"; 100) → 1,10775126 (100% do CDI no período).
    - dataInicio / dataFim: 'dd/mm/aaaa', data do Excel ou TODAY();
    - percentual: % do CDI como NÚMERO puro (100 = 100%, 110 = 110%); opcional, default 100.
      Atenção: NÃO é % do Excel — passe 100, não 100%.
    Multiplique pelo valor base para capitalizar (ex.: valor * =CDI(...))."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ini = _data_iso(dataInicio)
        fim = _data_iso(dataFim)
        if percentual is None or percentual == "":
            percentual = 100.0
        fator = FatorDi(ini, fim, float(percentual))
        if fator is not None:
            return float(fator)
        return "ERRO: API DI sem resposta"
    except Exception as e:
        return f"ERRO: {e}"


@xw.func
def LIMPARCACHE():
    """Esvazia o cache de respostas das APIs (B3/FI) e força novas chamadas."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    LimparCache()
    return "cache limpo"
