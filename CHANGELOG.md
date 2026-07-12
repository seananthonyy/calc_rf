# CHANGELOG — AntonioOliveiraCalc

Versionamento: a **lógica** (`.py`) é retrocompatível (só adiciona UDF, nunca remove/renomeia);
o **`.xlam` é versionado por nome de arquivo** (`AntonioOliveiraCalc_v{N}.xlam`). Versões antigas
ficam na share e nunca são apagadas. `VERSION` no `AntonioOliveiraCalc.py` acompanha a lógica.

## v2.0.0 — 2026-07-12 (remove UDFs → exige re-bake; novo `_v2.xlam`, o `_v1` permanece)
Removidas 11 UDFs a pedido:
- Métricas: `cpConvexidade`, `cpDurMod`, `cpDiPerc`, `cpSpreadDi`.
- BCB/IBGE: `cpSelicOver`, `cpCdiDia`, `cpIpcaAno`, `cpIgpDi`, `cpInpc`, `cpPoupanca`, `cpTr`.
BCB/IBGE mantidos: `cpSelic`, `cpCdiAno`, `cpIpca`, `cpIpca15`, `cpIpcaIndice`, `cpIgpm`, `cpDolar`, `cpEuro`.
**32 UDFs.** Major bump = mudança da superfície de UDFs → novo `AntonioOliveiraCalc_v2.xlam` (o `_v1`
fica na share e continua funcionando p/ quem já usa). Precisa re-bake (Excel fechado).

## v1.0.1 — 2026-07-12 (logic-only — sem re-bake; .xlam segue _v1)
- **Fix:** cpGrossUp/cpGrossUpTipo/cpConvexidade/cpDv01/cpDurMod/cpDiPerc/cpSpreadDi quebravam com
  `name 'CampoFi' is not defined` — `CampoFi` agora importado do `apis.py`.
- `cpTaxaEmissao` passa a retornar em **decimal** (6,4618% → 0,064618; %DI 113,5% → 1,135).
- `cpFluxoCompleto`: saída enxugada para **Data · %Amort · %Incorp** (removida a coluna Tipo).
- Funções BCB/IBGE agora **exigem data** (as-of); sem data → "ERRO: informe a data".

## v1.0.0 — 2026-07-12
- **Nova UDF `=cpFluxoCompleto(ticker)`**: agenda cadastrada inteira do papel (B3 `getBondDetails`),
  spill `Data · Tipo · %Amort · %Incorp`, datas ajustadas por dia útil (ANBIMA). Distingue
  Amortização, Incorporação (só em papéis `method=IPCA-I`) e Cupom (o `yield` do `J` em não-IPCA-I
  é a taxa do cupom pago, não incorporação — não é exibido como incorp).
- `=cpTeste()` agora mostra a versão (`OK v1.0.0 — path: …`).
- Bundle (`AntonioOliveiraCalc_bundle.py`) regenerado via `gerar_bundle.py`; embute o `.xlam`
  versionado `_v1`.
- Consolidação de documentação em `README.md` (uso/instalação) + `CLAUDE.md` (arquitetura/dev).
- Correção: o `PYTHONPATH` baked real é `Z:\AntonioOliveira\AntonioOliveiraCalc` (docs antigos diziam
  `...\CalcRF`, o que causaria falha de install).

## histórico anterior (pré-versionamento)
Add-in `cp*` API-only (B3 → FI → BCB/IBGE; DI local). 43 UDFs. Ver `git log` para detalhes.
