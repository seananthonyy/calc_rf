# CHANGELOG — AntonioOliveiraCalc

Versionamento: a **lógica** (`.py`) é retrocompatível; o **`.xlam` é versionado por nome de arquivo**
(`AntonioOliveiraCalc_vN.xlam`) — versões antigas ficam na share e não quebram quem já usa.
`VERSION` no `AntonioOliveiraCalc.py` acompanha a lógica.

## v1.0.1 — 2026-07-13
Fluxos: as duas UDFs de agenda passam a mostrar **toda data de evento** (as de cupom puro saem com
`%Amort`=0 e `%Incorp`=0). Antes elas eram filtradas, e num papel bullet com cupom (ISAEC2:
`method=IPCA`, 30 eventos `J` + 1 `A`) sobrava só o vencimento.

- `cpFluxo` muda de **formato**: era `Data·Tipo·Prazo(DU)·VF·VP` (fluxo em R$ calculado via FI/B3),
  agora é `Data·Tipo·%Amort·%Incorp` — a **agenda restante** a partir da data, da mesma fonte do
  `cpFluxoCompleto` (B3 `getBondDetails`). `Tipo`: `J` (só juros), `J+A`, `J+I` (incorporação, só
  em `IPCA-I`). O argumento `taxa` segue aceito e é **ignorado** (assinatura preservada → sem re-bake).
- `cpFluxoCompleto`: mesma agenda, escopo **inteiro** (desde a emissão). Colunas inalteradas.
- ⚠️ A soma dos `%Amort` cadastrados pode fechar **abaixo de 100** (FGEN13: 91,975) — a B3 não
  cadastra o principal residual quitado no vencimento. A linha do vencimento sai com o que a B3
  informa; nada é inferido.

## v1.0.0 — 2026-07-12
Add-in `cp*` API-only (B3 → FI → bondbuilder; DI local; BCB/SGS; IBGE SIDRA). **32 UDFs.**

- **Precificação:** `cpPu`, `cpTaxa`, `cpDur`, `cpCdi` (fator CDI acumulado, API B3).
- **Papel:** `cpPupar`, `cpVna`, `cpFluxo` (fluxo de caixa **restante calculado**, FI→B3),
  `cpFluxoCompleto` (**agenda cadastrada** `Data·%Amort·%Incorp`, B3 `getBondDetails`, agregada por
  data — vencimento não duplica; linhas puras de cupom omitidas), `cpVencimento`, `cpEmissao`,
  `cpInicioRentabilidade`, `cpTaxaEmissao` (em **decimal**), `cpVne`, `cpAniversario`.
- **Métricas FI:** `cpGrossUp`, `cpGrossUpTipo`, `cpDv01`. `cpGrossUp` e `cpDv01` **exigem a taxa**
  (dependem da taxa negociada).
- **Indicadores BCB/IBGE (exigem data):** taxas em **decimal** (6% → 0,06) — `cpSelic`, `cpCdiAno`,
  `cpIpca`, `cpIpca15`, `cpIgpm`. Número-índice/câmbio ficam em valor natural — `cpIpcaIndice`,
  `cpDolar`, `cpEuro`.
- **Dias úteis (feriados ANBIMA):** `cpEhDiaUtil`, `cpDiasUteis`, `cpDiaUtilPosterior`,
  `cpDiaUtilAnterior`, `cpDiaUtilMaisN`.
- **Diagnóstico:** `cpTeste` (mostra a versão), `cpLimparCache`.
- `apis.py` **sem circuit-breaker** (removido) — cada chamada tenta a rede direto (com cache + timeout).
- `PYTHONPATH` baked no `.xlam`: `Z:\AntonioOliveira\AntonioOliveiraCalc`.
- **Guia dos traders `COMO_USAR.html`** (vai no bundle): paleta Itaú, fonte de dados por fórmula
  (B3 / FI Analytics / Banco Central / IBGE / cálculo próprio), câmbio rotulado PTAX de venda.
- **Empacotamento:** um único `.xlam` por versão (`AntonioOliveiraCalc_v1.xlam`); a versão é a
  constante `VERSION` no `AntonioOliveiraCalc.py` (sem arquivo `VERSION` avulso). O
  `AntonioOliveiraCalc_bundle.py` é **gerado** por `gerar_bundle.py` (gitignored — não versionar).
