# CHANGELOG — AntonioOliveiraCalc

Versionamento: a **lógica** (`.py`) é retrocompatível; o **`.xlam` é versionado por nome de arquivo**
(`AntonioOliveiraCalc_vN.xlam`) — versões antigas ficam na share e não quebram quem já usa.
`VERSION` no `AntonioOliveiraCalc.py` acompanha a lógica.

## v1.0.2 — 2026-07-13
**`cpGrossUp`, `cpGrossUpTipo` e `cpDv01` passam a funcionar em papel do bondbuilder** (LCD, LF,
CDB… — os que não existem no `/deb` nem no `/cr`). As três usam `CampoFi` → `_FiFull`, que só
tentava `/deb` e `/cr`; agora cai no **bondbuilder** como 3ª fonte (novo `apis._BbFull`), a mesma
cadeia que o `Preco`/`TaxaOp` já fazia — por isso `cpPu`/`cpTaxa`/`cpDur` funcionavam nesses papéis
e o gross up não. O `/bb` devolve os mesmos campos (`taxedM2MRate`, `taxedType`, `dv01`).
Validado no `LCD BNDES 10 ANOS`: gross up 0,150595 · dv01 −0,420325 · tipo `GROSS_UP`. Sem regressão
em debênture (FGEN13 inalterado).

> ⚠️ Papel do bondbuilder **não tem cadastro na B3**, então as UDFs que dependem do `getBondDetails`
> (`cpVencimento`, `cpTaxaEmissao`, `cpFluxo`, `cpFluxoCompleto`…) seguem indisponíveis nele — e a
> taxa vira **obrigatória** no `cpGrossUpTipo` (não há taxa de emissão pra herdar).

## v1.0.1 — 2026-07-13

**Bundle (auto-extrator):** volta a ser **versionado no repo** (estava no `.gitignore` como
"derivado" — mas é o único jeito de instalar no banco, que não tem `git pull`). E agora ele **pula
arquivo já idêntico** em vez de reescrever tudo: como o `.xlam` só muda quando há UDF/assinatura
nova, atualizar a lógica (`.py`) **não exige mais fechar o Excel** — o `.xlam` travado é pulado por
ser igual. Se um arquivo mudou E está travado, o extrator diz qual e sai com erro (antes: estourava
`PermissionError` no `.xlam`, que é o 1º da lista, e não escrevia nenhum `.py`).

Fluxos: as duas UDFs de agenda passam a mostrar **toda data de evento** (as de cupom puro saem com
`%Amort`=0 e `%Incorp`=0). Antes elas eram filtradas, e num papel bullet com cupom (ISAEC2:
`method=IPCA`, 30 eventos `J` + 1 `A`) sobrava só o vencimento.

- `cpFluxo` muda de **formato**: era `Data·Tipo·Prazo(DU)·VF·VP` (fluxo em R$ calculado via FI/B3),
  agora é `Data·Tipo·%Amort·%Incorp` — a **agenda restante** a partir da data, da mesma fonte do
  `cpFluxoCompleto` (B3 `getBondDetails`). `Tipo`: `J` (só juros), `J+A`, `J+I` (incorporação, só
  em `IPCA-I`). O argumento `taxa` segue aceito e é **ignorado** (assinatura preservada → sem re-bake).
- `cpFluxoCompleto`: mesma agenda, escopo **inteiro** (desde a emissão). Colunas inalteradas.
- **Principal residual do vencimento** (`apis._CompletarResiduo`): a B3 não cadastra o principal
  quitado no vencimento dos papéis **IPCA-I** (FGEN13: os `A` somam 91,975; MESA13: 86,823). A linha
  do vencimento recebe `100 − Σ%amort`, e a agenda passa a fechar 100%. Valor **derivado**, não
  cadastrado — mas conferido contra o `trades.db` (fonte independente): FGEN13 8,025 e MESA13
  13,176856 no vencimento, iguais aos derivados (delta ≤ 1e-15). Papéis cujos `A` já somam 100
  (IPCA simples: ISAEC2, PALF38, SAVI13, ATHD11, VIALA5) não são tocados.

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
