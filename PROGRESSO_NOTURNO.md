# Trabalho noturno no add-in — PROGRESSO (10→11/07/2026)

> Doc de checkpoint. Atualizo a cada etapa pra não perder contexto se bater no limite de tokens.

## Objetivo (pedido do usuário, dormindo até ~7h)
1. Renomear TODAS as UDFs com prefixo `cp` (`=cpPu`, `=cpTaxa`, `=cpDur`, `=cpCdi`…).
2. Novas funções de dados (das APIs B3/FI): PU Par, fluxo restante (spill, na data calculada),
   VNA, vencimento, emissão, taxa de emissão, gross up (FI).
3. Funções do Banco Central (SGS) — várias (SELIC, CDI, IPCA, IGP-M, dólar PTAX, …).
4. Funções de dias úteis (contar, dia anterior, dia posterior) com feriados ANBIMA.
5. Performance: mais rápido, travar menos o Excel.

## Decisões do usuário (autorizações)
- **Excel**: usuário fecha antes de dormir → EU re-bako o `.xlam` quando estiver fechado.
- **Fonte de dados**: APIs (B3/FI) primeiro; add-in segue API-only. Onde a API não expõe → aviso.
- **Codebase**: editar PROD `D:\ItauBBA\AntonioOliveiraCalc`; sincronizar `apis.py` no dev
  `D:\ItauBBA\calculadora-renda-fixa`.
- **Git**: só commits LOCAIS em checkpoints (sem push).
- Autonomia total; não esperar o usuário.

## Estado atual do add-in (antes de começar)
- `AntonioOliveiraCalc.py`: UDFs `TESTE, PU, DUR, TAXA, CDI, LIMPARCACHE` (@xw.func).
- `apis.py`: B3(`PrecoB3/TaxaB3/FatorDi`)→FI(`PrecoFi/TaxaFi`)→bondbuilder(`PrecoBb/TaxaBb`);
  roteadores `Preco/TaxaOp`; cache disco/mem TTL; circuit-breaker; memo de fonte. Só extrai
  `pu/duration/pupar` das respostas.
- `di.py`: precifica DI1 local + tem `EhDu/ProximoDu/ContarDu` (feriados ANBIMA) — REUSAR p/ dias úteis.
- `.xlam` PROD aponta PYTHONPATH `Z:\AntonioOliveira\AntonioOliveiraCalc` (banco). DEV `.xlam` p/ testar local.
- Re-bake: via COM `xlwings.udfs.import_udfs("AntonioOliveiraCalc", wb.api)` (precisa Excel fechado).

## Plano de execução (ordem)
- [ ] E0. Descobrir campos das respostas B3(calcPU/getBondDetails) e FI (grossUp/VNA/fluxo/datas).
- [ ] E1. apis.py: funções que expõem os campos novos (Detalhes/BondDetails/Fluxo/GrossUp) + SGS BCB.
- [ ] E2. UDF module: renomear tudo p/ `cp*` + novas UDFs (dados, BCB, dias úteis).
- [ ] E3. Performance.
- [ ] E4. Re-bake `.xlam` (quando Excel fechado) + scrub PII + bundle.
- [ ] E5. Docs + commit local.

## Mapa de campos das APIs (descoberto E0)
**B3 calcPU** (`/calcPU/{tk}/{data}/{taxa}`): `PU, PUPar, VNA, cashFlowList[{date,eventType,...}],
duration, interest, issuer, method, issuerYield, yield`.
**B3 getBondDetails** (`/getBondDetails/{tk}`, SEM data): `expiredate`(venc), `issuedate`(emissão),
`startingdate`(início rent), `vne`, `yield`(taxa emissão, unidade nativa: 113.5 / 6.4618),
`anniversaryday`, `events`, `issuer`, `method`, `firstDay`, `status`.
**FI** (`/deb`·`/cr`, rate): TUDO num call — `m2m`(pu), `currentNotionalPlusAccruedInterest`(pupar),
`adjustedFaceValue`(VNA), `maculayDuration/modifiedDuration/convexity/dv01/dv1d/cv01`,
`accruedInterest`, `issueDate`, `maturityDate`, `issueRate`(decimal), `cashFlowEvents[{date,
eventType,term,rate,futureValue,presentValue}]`, `diPercentage`, `spreadOverDI/NTNB/PRE/DAP`,
`nominalRate`, `inflationAdjustment`, `currentAmortization/Coupon/Premium`,
**gross up** = `taxedM2MRate` + `taxedType`("GROSS_UP" isento / "AFTER_TAX" tributado) + todos `taxed*`.
→ FI é a fonte mais rica (1 call); B3 getBondDetails p/ datas/taxa em unidade nativa.

## BCB SGS (público, sem auth): séries úteis
432 Selic meta%aa | 1178 Selic over%aa | 12 CDI%dia | 4389 CDI%aa | 433 IPCA%mês |
13522 IPCA acum12m | 189 IGP-M%mês | 190 IGP-DI | 188 INPC | 226 TR | 195/196 poupança |
1 Dólar PTAX venda | 21619 Euro | 7832 IPCA-15. Endpoint `/dados/ultimos/{N}` p/ último valor.

## UDFs FINAIS (todas prefixo cp) — implementadas e testadas ✓
**Preço:** cpPu, cpTaxa, cpDur, cpCdi. **Dados:** cpPupar, cpVna, cpFluxo(spill),
cpVencimento, cpEmissao, cpInicio, cpTaxaEmissao, cpVne, cpAniversario. **Métricas FI:**
cpGrossUp, cpGrossUpTipo, cpConvexidade, cpDv01, cpDurMod, cpDiPerc, cpSpreadDi. **Genéricos:**
cpFi(campo), cpBond(campo). **BCB:** cpSelic, cpSelicOver, cpCdiAno, cpCdiDia, cpIpca, cpIpcaAno,
cpIgpm, cpInpc, cpTr, cpDolar([data]), cpEuro([data]), cpBcb(serie,[ini],[fim]). **Dias úteis:**
cpEhDiaUtil, cpDiasUteis, cpDiaUtilPosterior, cpDiaUtilAnterior, cpDiaUtilMaisN. **Diag:** cpTeste, cpLimparCache.
- pupar/vna/fluxo/grossup: **taxa OPCIONAL** (usa taxa de emissão da B3 se omitida).
- Datas retornam como `date` (Excel formata). cpFluxo/cpBcb(c/datas) = spill 2D.

## Log de progresso
- E0 ✓ descoberta de campos.
- E1 ✓ apis.py: +vna no Preco (cache compartilhada c/ cpPu — PERF), +BondDetailsB3, _CalcPuB3Full,
  _FiFull, Detalhes, CampoFi/CampoB3/CampoBond, FluxoRestante, BcbSerie/BcbValor. Testado OK.
- E2 ✓ AntonioOliveiraCalc.py reescrito: todas UDFs cp* + novas. Testado OK (dados, BCB, dias úteis).
- sync apis.py→dev ✓ (cp PROD→dev, direção segura).
- E4 ✓ RE-BAKE feito (Excel fechado): rebake.py via COM (import_udfs) no DEV; transplantei o
  vbaProject.bin novo (scrubbed anton→user1 ascii+utf16) pros .xlam LIMPOS dos backups
  (.bak_noturno), preservando config (Z:\ PROD / D:\ DEV) + docProps limpo. Verificado:
  cpPu=9/cpSelic=8 no bin dos dois, 0 PII, zip íntegro, configs intactas. Bundle regenerado
  (round-trip byte-idêntico, 0 segredos). Backups: *.xlam.bak_noturno.
- Falta: E5 docs (CLAUDE.md/LEIA-ME) + commit final; E3 nota de performance; (opcional) +funções.

## PERFORMANCE (E3) — decisões
- UDFs SÍNCRONAS de propósito (async foi revertido: causava loop de write-back no SharePoint).
- Ganho principal: `vna` no dict do `Preco` → cpPu/cpPupar/cpVna/cpDur compartilham 1 chamada
  (mesma cache). FI-only (grossup/convexidade/dv01) compartilham a cache `_FiFull`. Estáticos
  (venc/emissão/taxaEmi/vne) compartilham `BondDetailsB3` (1 call/ticker). BCB e /di/calculo cacheados.
  → p/ um papel com N fórmulas cp*, só ~3 chamadas de rede distintas (calcPU, FI, getBondDetails),
  e só na 1ª vez (cache disco TTL 600s sobrevive a reabrir Excel). Timeout 6s + circuit-breaker.
