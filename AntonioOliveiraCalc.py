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

# Todas as UDFs têm o prefixo cp (ex.: =cpPu, =cpTaxa, =cpVna...).
# Títulos com API (debênture, CRI/CRA, NTN-B…): via B3 → FI Analytics → bondbuilder.
# DI (tickers DI1...): sem API → calculado LOCALMENTE em di.py.
# Séries macro (SELIC, CDI, IPCA, dólar…): API pública do Banco Central (SGS).
_IMPORT_ERROR = None
try:
    from apis import (
        Preco, TaxaOp, LimparCache, FatorDi,
        CampoFi, CampoBond, FluxoRestante, BcbSerie, BcbValor,
    )
    import di
except Exception as _e:
    _IMPORT_ERROR = str(_e)


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
        "AntonioOliveiraCalc — use =cpPu, =cpTaxa, =cpVna, =cpFluxo, =cpSelic… (prefixo cp)"
    )


# =============================================================================
# DIAGNÓSTICO
# =============================================================================

@xw.func
def cpTeste():
    """Diagnóstico: OK ou o erro de import."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    return f"OK — path: {_THIS_DIR}"


@xw.func
def cpLimparCache():
    """Esvazia o cache das APIs (B3/FI/BCB) e força novas chamadas."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    LimparCache()
    return "cache limpo"


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
    """Fluxo de caixa RESTANTE na data (spill): colunas Data | Tipo | Prazo(DU) | VF | VP.
    taxa opcional (VF/VP dependem da taxa só no VP; se omitida usa a de emissão)."""
    if _IMPORT_ERROR:
        return [[f"ERRO import: {_IMPORT_ERROR}"]]
    try:
        ticker = str(ticker).upper().strip()
        fl = FluxoRestante(ticker, _data_iso(data), _resolve_taxa(ticker, taxa))
        if not fl:
            return [["ERRO: fluxo indisponível (B3/FI)"]]
        linhas = [["Data", "Tipo", "Prazo(DU)", "VF", "VP"]]
        for e in fl:
            linhas.append([
                _data_saida(e.get("data")), e.get("tipo"), e.get("prazo"),
                e.get("vf"), e.get("vp"),
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
def cpInicio(ticker):
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
    """Taxa de emissão do papel em unidade nativa (ex.: 113.5 para %DI, 6.4618 para IPCA+)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        v = CampoBond(str(ticker).upper().strip(), "yield")
        return float(v) if v is not None else "ERRO: taxa de emissão indisponível"
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
# MÉTRICAS / GROSS UP (FI Analytics)
# =============================================================================

@xw.func
@xw.arg('taxa', numbers=float)
def cpGrossUp(ticker, data, taxa=None):
    """Gross up / taxa equivalente após imposto (FI 'taxedM2MRate'). Veja o tipo em =cpGrossUpTipo."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        v = CampoFi(ticker, _data_iso(data), _resolve_taxa(ticker, taxa), "taxedM2MRate")
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
def cpConvexidade(ticker, data, taxa=None):
    """Convexidade (FI Analytics)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        v = CampoFi(ticker, _data_iso(data), _resolve_taxa(ticker, taxa), "convexity")
        return float(v) if v is not None else "ERRO: indisponível (FI)"
    except Exception as e:
        return _erro("conv", e)


@xw.func
@xw.arg('taxa', numbers=float)
def cpDv01(ticker, data, taxa=None):
    """DV01 — variação do PU por 1 bp na taxa (FI Analytics)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        v = CampoFi(ticker, _data_iso(data), _resolve_taxa(ticker, taxa), "dv01")
        return float(v) if v is not None else "ERRO: indisponível (FI)"
    except Exception as e:
        return _erro("dv01", e)


@xw.func
@xw.arg('taxa', numbers=float)
def cpDurMod(ticker, data, taxa=None):
    """Duration modificada em anos (FI Analytics)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        v = CampoFi(ticker, _data_iso(data), _resolve_taxa(ticker, taxa), "modifiedDuration")
        return float(v) if v is not None else "ERRO: indisponível (FI)"
    except Exception as e:
        return _erro("durmod", e)


@xw.func
@xw.arg('taxa', numbers=float)
def cpDiPerc(ticker, data, taxa=None):
    """Percentual do DI equivalente (FI 'diPercentage'; ex.: 1.135 = 113,5% do DI)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        v = CampoFi(ticker, _data_iso(data), _resolve_taxa(ticker, taxa), "diPercentage")
        return float(v) if v is not None else "ERRO: indisponível (FI)"
    except Exception as e:
        return _erro("diperc", e)


@xw.func
@xw.arg('taxa', numbers=float)
def cpSpreadDi(ticker, data, taxa=None):
    """Spread sobre o DI (FI 'spreadOverDI', em taxa aditiva)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        v = CampoFi(ticker, _data_iso(data), _resolve_taxa(ticker, taxa), "spreadOverDI")
        return float(v) if v is not None else "ERRO: indisponível (FI)"
    except Exception as e:
        return _erro("spreaddi", e)


# =============================================================================
# GENÉRICOS — qualquer campo das APIs (power user)
# =============================================================================

@xw.func
@xw.arg('taxa', numbers=float)
def cpFi(ticker, data, taxa, campo):
    """Qualquer campo da resposta da FI Analytics. Ex.: =cpFi("X";hoje;6,5%;"modifiedDuration").
    Campos úteis: taxedM2MRate, convexity, dv01, spreadOverDI, diPercentage, accruedInterest,
    modifiedDuration, spreadOverNTNB, spreadOverPRE, nominalRate, adjustedFaceValue, isin..."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        ticker = str(ticker).upper().strip()
        v = CampoFi(ticker, _data_iso(data), _resolve_taxa(ticker, taxa), str(campo).strip())
        if v is None:
            return "#N/D"
        return _data_saida(v) if isinstance(v, str) else v
    except Exception as e:
        return _erro("fi", e)


@xw.func
def cpBond(ticker, campo):
    """Qualquer campo do getBondDetails da B3 (estático). Ex.: =cpBond("X";"expiredate").
    Campos: expiredate, issuedate, startingdate, vne, yield, anniversaryday, issuer, method, status."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        v = CampoBond(str(ticker).upper().strip(), str(campo).strip())
        if v is None:
            return "#N/D"
        return _data_saida(v) if isinstance(v, str) else v
    except Exception as e:
        return _erro("bond", e)


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
    v = BcbValor(serie, _data_br(data) if not _vazio(data) else None)
    return float(v) if v is not None else "ERRO: BCB sem dado"


def _data_br(data):
    return _parse_data(data).strftime("%d/%m/%Y")


@xw.func
def cpSelic():
    """Meta SELIC atual (% a.a.) — BCB série 432."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(432)
    except Exception as e:
        return _erro("selic", e)


@xw.func
def cpSelicOver():
    """SELIC over anualizada (% a.a.) — BCB série 1178."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(1178)
    except Exception as e:
        return _erro("selicover", e)


@xw.func
def cpCdiAno():
    """CDI anualizado, base 252 (% a.a.) — BCB série 4389."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(4389)
    except Exception as e:
        return _erro("cdiano", e)


@xw.func
def cpCdiDia():
    """CDI do dia (% ao dia) — BCB série 12."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(12)
    except Exception as e:
        return _erro("cdidia", e)


@xw.func
def cpIpca():
    """IPCA do último mês divulgado (% no mês) — BCB série 433."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(433)
    except Exception as e:
        return _erro("ipca", e)


@xw.func
def cpIpcaAno():
    """IPCA acumulado em 12 meses (%) — BCB série 13522."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(13522)
    except Exception as e:
        return _erro("ipcaano", e)


@xw.func
def cpIgpm():
    """IGP-M do último mês (% no mês) — BCB série 189."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(189)
    except Exception as e:
        return _erro("igpm", e)


@xw.func
def cpInpc():
    """INPC do último mês (% no mês) — BCB série 188."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(188)
    except Exception as e:
        return _erro("inpc", e)


@xw.func
def cpTr():
    """TR — Taxa Referencial do último dia (%) — BCB série 226."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(226)
    except Exception as e:
        return _erro("tr", e)


@xw.func
def cpDolar(data=None):
    """Dólar PTAX venda (R$) — BCB série 1. data opcional (último por padrão)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(1, data)
    except Exception as e:
        return _erro("dolar", e)


@xw.func
def cpEuro(data=None):
    """Euro venda (R$) — BCB série 21619. data opcional (último por padrão)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        return _bcb_valor(21619, data)
    except Exception as e:
        return _erro("euro", e)


@xw.func
def cpBcb(serie, dataInicio=None, dataFim=None):
    """Série do Banco Central (SGS) por número.
    - Sem datas → valor mais recente (número).
    - Com dataInicio (e opcional dataFim) → série no período (spill: Data | Valor)."""
    if _IMPORT_ERROR:
        return f"ERRO import: {_IMPORT_ERROR}"
    try:
        serie = int(float(serie))
        if _vazio(dataInicio):
            v = BcbValor(serie)
            return float(v) if v is not None else "ERRO: BCB sem dado"
        ini = _data_br(dataInicio)
        fim = _data_br(dataFim) if not _vazio(dataFim) else None
        dados = BcbSerie(serie, ini=ini, fim=fim)
        if not dados:
            return [["ERRO: BCB sem dado no período"]]
        linhas = [["Data", "Valor"]]
        for d in dados:
            try:
                dt = datetime.strptime(d["data"], "%d/%m/%Y").date()
            except Exception:
                dt = d.get("data")
            try:
                val = float(str(d["valor"]).replace(",", "."))
            except Exception:
                val = d.get("valor")
            linhas.append([dt, val])
        return linhas
    except Exception as e:
        return f"ERRO: {e}"


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
