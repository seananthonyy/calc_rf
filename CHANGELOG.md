# CHANGELOG — AntonioOliveiraCalc

Versionamento: a **lógica** (`.py`) é retrocompatível (só adiciona UDF, nunca remove/renomeia);
o **`.xlam` é versionado por nome de arquivo** (`AntonioOliveiraCalc_v{N}.xlam`). Versões antigas
ficam na share e nunca são apagadas. `VERSION` no `AntonioOliveiraCalc.py` acompanha a lógica.

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
