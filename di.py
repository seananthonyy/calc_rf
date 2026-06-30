"""
DI — Contratos Futuros de DI (DI1)
==================================
Precificação LOCAL (sem API) de contratos futuros de DI da B3, base 252 dias úteis.

Self-contained: carrega o calendário de feriados de `feriados_anbima.csv` (na mesma pasta
deste arquivo ou em `./files/`) e tem as próprias funções de dia útil — não depende do
`calculadora_rf`, para poder ser distribuído junto do add-in.

Cobre QUALQUER vencimento mensal (F,G,H,J,K,M,N,Q,U,V,X,Z), não só DI1F.

Vencimento
----------
1º dia útil do mês/ano do contrato (próximo DU se cair em FDS/feriado).
    mês = 4º caractere do ticker; ano = 2 últimos dígitos.  Ex.: DI1F25 → 1º DU de jan/2025.

Fórmulas (du = dias úteis entre a data base e o vencimento)
----------------------------------------------------------
    PU       = 100000 / (1 + taxa/100) ^ (du/252)
    Taxa     = ((100000 / PU) ^ (252/du) - 1) × 100
    Duration = du / 252
"""

import csv
import re
from datetime import date, timedelta
from pathlib import Path

# código do mês de vencimento (4º caractere do ticker)
MESES_DI = {
    'F': 1, 'G': 2, 'H': 3, 'J': 4, 'K': 5, 'M': 6,
    'N': 7, 'Q': 8, 'U': 9, 'V': 10, 'X': 11, 'Z': 12,
}

_PADRAO_DI = re.compile(r'^DI1?([FGHJKMNQUVXZ])(\d{2})$')


# ─── CALENDÁRIO DE DIAS ÚTEIS (self-contained) ───────────────────────────────

def _CarregarFeriados() -> frozenset:
    base = Path(__file__).parent
    for caminho in (base / 'feriados_anbima.csv', base / 'files' / 'feriados_anbima.csv'):
        if caminho.exists():
            with open(caminho, encoding='utf-8') as f:
                return frozenset(date.fromisoformat(l['data'].strip()) for l in csv.DictReader(f))
    raise FileNotFoundError('feriados_anbima.csv não encontrado (mesma pasta ou ./files/)')


FERIADOS = _CarregarFeriados()


def EhDu(d: date, feriados=FERIADOS) -> bool:
    return d.weekday() < 5 and d not in feriados


def ProximoDu(d: date, feriados=FERIADOS) -> date:
    while not EhDu(d, feriados):
        d += timedelta(days=1)
    return d


def ContarDu(ini: date, fim: date, feriados=FERIADOS) -> int:
    if ini >= fim:
        return 0
    return sum(1 for i in range((fim - ini).days) if EhDu(ini + timedelta(days=i), feriados))


# ─── DI ──────────────────────────────────────────────────────────────────────

def EhTickerDi(ticker) -> bool:
    """True se o ticker é um contrato futuro de DI (ex.: DI1F27, DI1N29, DIF28)."""
    return bool(_PADRAO_DI.match(str(ticker).upper().strip()))


def _Parsear(ticker) -> date:
    m = _PADRAO_DI.match(str(ticker).upper().strip())
    if not m:
        raise ValueError(f"Ticker DI inválido: {ticker!r} (esperado tipo 'DI1F25')")
    return date(2000 + int(m.group(2)), MESES_DI[m.group(1)], 1)


def VencimentoDi(ticker) -> date:
    """Vencimento: 1º dia útil do mês/ano (salta FDS/feriado)."""
    return ProximoDu(_Parsear(ticker))


def _Du(dataBase: date, ticker) -> int:
    du = ContarDu(dataBase, VencimentoDi(ticker))
    if du <= 0:
        raise ValueError(f"Vencimento de {ticker} não é futuro em relação a {dataBase}")
    return du


def PuDi(taxa: float, dataBase: date, ticker) -> float:
    """PU dado a taxa (% a.a.):  PU = 100000 / (1 + taxa/100) ^ (du/252)."""
    return 100000 / ((1 + taxa / 100) ** (_Du(dataBase, ticker) / 252))


def TaxaDi(pu: float, dataBase: date, ticker) -> float:
    """Taxa (% a.a.) dado o PU:  Taxa = ((100000/PU) ^ (252/du) - 1) × 100."""
    return ((100000 / pu) ** (252 / _Du(dataBase, ticker)) - 1) * 100


def DurationDi(dataBase: date, ticker) -> float:
    """Duration (anos) = du / 252."""
    return _Du(dataBase, ticker) / 252
