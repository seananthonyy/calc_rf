# CHANGELOG — AntonioOliveiraCalc

Versionamento: a **lógica** (`.py`) é retrocompatível; o **`.xlam` é versionado por nome de arquivo**
(`AntonioOliveiraCalc_vN.xlam`) — versões antigas ficam na share e não quebram quem já usa.
`VERSION` no `AntonioOliveiraCalc.py` acompanha a lógica.

## v1.0.0 — 2026-07-12
Add-in `cp*` API-only (B3 → FI → bondbuilder; DI local; BCB/SGS; IBGE SIDRA). **32 UDFs.**

- **Precificação:** `cpPu`, `cpTaxa`, `cpDur`, `cpCdi` (fator CDI acumulado, API B3).
- **Papel:** `cpPupar`, `cpVna`, `cpFluxo` (fluxo de caixa **restante calculado**, FI→B3),
  `cpFluxoCompleto` (**agenda cadastrada** `Data·%Amort·%Incorp`, B3 `getBondDetails`, agregada por
  data — vencimento não duplica; linhas puras de cupom omitidas), `cpVencimento`, `cpEmissao`,
  `cpInicioRentabilidade`, `cpTaxaEmissao` (em **decimal**), `cpVne`, `cpAniversario`.
- **Métricas FI:** `cpGrossUp`, `cpGrossUpTipo`, `cpDv01`. `cpGrossUp` e `cpDv01` **exigem a taxa**
  (dependem da taxa negociada).
- **Indicadores BCB/IBGE (exigem data):** `cpSelic`, `cpCdiAno`, `cpIpca`, `cpIpca15`, `cpIpcaIndice`,
  `cpIgpm`, `cpDolar`, `cpEuro`.
- **Dias úteis (feriados ANBIMA):** `cpEhDiaUtil`, `cpDiasUteis`, `cpDiaUtilPosterior`,
  `cpDiaUtilAnterior`, `cpDiaUtilMaisN`.
- **Diagnóstico:** `cpTeste` (mostra a versão), `cpLimparCache`.
- `apis.py` **sem circuit-breaker** (removido) — cada chamada tenta a rede direto (com cache + timeout).
- `PYTHONPATH` baked no `.xlam`: `Z:\AntonioOliveira\AntonioOliveiraCalc`.
