# CHANGELOG — CalcCP

Versionamento: a **lógica** (`.py`) é retrocompatível; o **`.xlam` é versionado por nome de arquivo**
(`CalcCP_vN.xlam`) — versões antigas ficam na share e não quebram quem já usa.
`VERSION` no `CalcCP.py` acompanha a lógica.

## v4.1.0 — 2026-07-27 — `=cpFonteCalculo` (37 UDFs) — `CalcCP_v4.xlam` REGERADO

> ⚠️ **O `CalcCP_v4.xlam` publicado horas antes (v4.0.0, 36 UDFs) foi substituído por este**, em vez
> de virar um v5. Foi decisão consciente: o v4.0.0 ainda **não tinha sido instalado em nenhum PC**
> (saiu do repositório para o banco no mesmo dia), então regerar o próprio v4 evita uma segunda
> re-registração do add-in. Se você chegou a baixar o v4 antes desta linha existir, baixe de novo —
> um v4 de 36 UDFs dá `#NAME?` no `=cpFonteCalculo`. `v1`, `v2` e `v3` seguem intocados.

**1 fórmula nova (37 no total):**

| Fórmula | Retorno |
|---|---|
| `=cpFonteCalculo(ticker; [data]; [taxa%])` | de onde vem o cálculo do papel: `B3`, `FI Analytics`, `FI Analytics (bondbuilder)` ou `DI (local)` |

- **É auditoria da cascata, não uma fonte a mais.** `=cpPu`/`=cpTaxa`/`=cpDur` tentam B3 → FI →
  bondbuilder e usam a primeira que responder; esta fórmula diz **qual foi**. A lógica ficou em
  `apis.FonteCalculo`, ao lado da própria cascata, então as duas não podem divergir.
- **Determinística, não depende da ordem de cálculo do Excel.** Se o `_fonteTicker` (memo por ticker)
  já sabe, responde sem rede; se não sabe, **sonda com a mesma cascata** em vez de devolver "não sei".
  Uma fórmula cuja resposta mudasse conforme a célula fosse calculada antes ou depois do `=cpPu`
  seria inútil para auditoria.
- **Caminho rápido de verdade:** o memo é consultado **antes** de resolver data/taxa — resolver a
  taxa custaria uma ida à B3 (`getBondDetails`) só para descobrir algo que já sabíamos. Medido:
  0,03 ms com o memo quente, contra ~120 ms se a resolução viesse primeiro.
- **`data` e `taxa` são opcionais** e servem só para a sondagem — quem responde não depende delas.
  Sem `data` usa hoje (informe a data se o papel **já venceu**); sem `taxa` usa a de emissão e, se o
  papel não estiver no cadastro da B3, uma taxa neutra, para conseguir sondar a FI mesmo assim.
- Validado ao vivo contra as duas fontes: `FGEN13`/`NTNB35` → `B3`; `CRA0210012Y`, `24I1419236`,
  `CRA0190066O`, `26C5564546` (tickers que a base marca como precificados pela FI) → `FI Analytics`;
  `DI1F27` → `DI (local)`; ticker inexistente → `#N/A`. O rótulo do **bondbuilder** não foi
  exercitado ao vivo (cobre LCD/LF/CDB/CPF, que não aparecem na base de negociação secundária).

## v4.0.0 — 2026-07-27 — `CalcCP_v4.xlam`: `=cpAnbimaSpread` + coluna Spread no histórico

**1 fórmula nova (36 no total)**, na mesma família das ANBIMA — vem do `trades.db`, não de API:

| Fórmula | Retorno |
|---|---|
| `=cpAnbimaSpread(ticker)` | spread ANBIMA mais recente do papel sobre a referência de `=cpAnbimaRef`, em **decimal** (`-0,005193` = −51,93 bps) |

E o `=cpAnbimaIndicativoHistorico` ganhou uma **5ª coluna, `Spread`** (mesma unidade), ao lado da
Taxa Indicativa. Quem já usa a fórmula passa a ver uma coluna a mais no spill — se houver algo
escrito à direita do intervalo, o Excel devolve `#SPILL!` até a célula ser liberada.

- **Fonte:** `AnbimaIndicativos.vrSpreadAnbima`, calculado pelo pipeline de negociação secundária.
  Uma consulta indexada nova em `basedados.SpreadAnbima` (mesmo padrão da taxa: linha mais recente
  com valor não-nulo, cache de 5 min). O histórico passou a trazer a coluna na **mesma** consulta —
  nenhuma consulta extra.
- **Unidade:** decimal, igual à taxa. Com referência **`FUNDING`** (papéis CDI+ e %CDI) o spread **é
  a própria taxa indicativa** — não há benchmark externo a descontar (CDI+ `0,006967` = CDI +
  0,6967%; %CDI `1,029075` = 102,9% do CDI). Nos demais é a taxa do papel contra a do vértice.
- **É mais esparso que a taxa** (~42,9k de 136,1k linhas da base): exige referência cadastrada **e**
  curva de mercado na data. Logo `#N/A` aparece em mais papéis/datas que em `=cpAnbimaIndicativo`, a
  data do último spread pode ser **anterior** à da última taxa, e no histórico as datas sem spread
  saem com `#N/A` só nessa coluna.
- **Arquivo NOVO `CalcCP_v4.xlam`** (v1, v2 e v3 ficam intocados na pasta) — fórmula nova exige
  re-bake, e a política é **versão por arquivo**: quem quiser o spread registra o v4; quem não migrar
  segue no v3 sem quebrar. Migração = desmarcar o v3 nos Suplementos e procurar o v4.
  `vbaProject.bin` regenerado com as 36 UDFs (201.728 → 203.776 bytes) e transplantado para uma cópia
  do v3, que já tinha a config `%CALCCP_DIR%;Z:\CP` e os `docProps` limpos — verificado no arquivo
  final: 36 UDFs no `CodeModule`, `CallUDF("CalcCP")`, zip íntegro, 0 PII, `absPath` ausente.
- **Quem ficar no v3 pega metade:** os `.py` novos valem para qualquer `.xlam` (o VBA só chama o
  Python), então a **coluna Spread do histórico aparece no v3 também** — ela não depende do re-bake.
  Só o `=cpAnbimaSpread`, que é UDF nova, dá `#NAME?` até registrar o v4.

## v3.0.1 — 2026-07-23 — otimização de performance (só lógica; **sem re-bake**, `.xlam` v3 inalterado)

Somente `.py` — nenhuma UDF nova, nenhuma assinatura mudou → o `CalcCP_v3.xlam` continua o mesmo
(o bundle o pula por ser idêntico). Atualizar = baixar o `CalcCP_bundle.py` e reabrir o Excel.
**Nenhum número ou formato de retorno mudou** (validado contra golden dos resultados da v3.0.0).

- **SQLite (fórmulas ANBIMA) usa índice em vez de varrer a tabela.** As consultas comparavam
  `upper(cdTicker) = ?`, o que **anulava o índice** e forçava um SCAN — no histórico (136k linhas)
  isso era ~160× mais lento que o necessário. Como a base guarda os tickers sempre em maiúsculas e
  já os maiusculamos no Python, passamos a comparar a coluna crua (`cdTicker = ?`): vira SEARCH por
  índice, e o `ORDER BY dtReferencia DESC` sai de graça (2ª coluna da PK, sem sort). Medido: 200
  consultas de referência de 163 ms → 1 ms.
- **Histórico sem JOIN.** `HistoricoIndicativoAnbima` fazia um JOIN linha-a-linha com InfoAtivos só
  para repetir a referência (constante por ticker) em cada linha. Agora busca a referência uma vez
  (reaproveitando o cache de `ReferenciaAnbima`) e faz uma consulta de tabela única, indexada.
  Resultado idêntico.
- **Cache em disco das APIs com gravação *debounced*.** Cada `_CacheSet` válido reescrevia o arquivo
  JSON inteiro (lendo-o antes) — num F9 de N células novas, N leituras + N reescritas. Agora o
  registro vive em memória e é gravado no máximo a cada 3 s, com `atexit` para o flush final ao
  fechar o Excel; o flush mescla o que outras instâncias gravaram. Um burst de 300 respostas passou
  de ~300 gravações para **1**.
- **DI local memoizado.** `VencimentoDi` e `ContarDu` (o laço dia-a-dia de dias úteis) ganharam
  `lru_cache`. Numa planilha com muitos contratos DI, `cpPu`/`cpDur`/`cpTaxa` do mesmo contrato+data
  e as várias linhas apontando para o mesmo vencimento passam a reusar a contagem: cenário de 400
  linhas × 3 fórmulas caiu de ~1790 ms para ~7 ms.

## v3.0.0 — 2026-07-23 — `CalcCP_v3.xlam`: fórmulas ANBIMA (referência e taxas indicativas)

**3 fórmulas novas (35 no total)**, as únicas do add-in que **não vêm de API** — esses dados só
existem no `trades.db` do projeto de negociação secundária, lido em **somente leitura**:

| Fórmula | Retorno |
|---|---|
| `=cpAnbimaRef(ticker)` | vértice/título público de referência do papel (`InfoAtivos.cdReferencia`) |
| `=cpAnbimaIndicativo(ticker)` | taxa indicativa ANBIMA mais recente da base, em **decimal** |
| `=cpAnbimaIndicativoHistorico(ticker)` | histórico completo (spill): Data · Ticker · Ref · Taxa Indicativa |

- **Arquivo NOVO `CalcCP_v3.xlam`** (v1 e v2 ficam intocados na pasta) — fórmula nova exige re-bake, e
  a política é versão por arquivo: quem quiser as fórmulas ANBIMA registra o v3; quem não migrar
  segue no v2 sem quebrar. `vbaProject.bin` regenerado com as 35 UDFs; config `%CALCCP_DIR%;Z:\CP`
  e `docProps` herdados do v2 (0 PII, `absPath` ausente, zip íntegro).
- **Módulo novo `basedados.py`** (vai no bundle): conexão SQLite por URI **`?mode=ro`** — o banco
  recusa escrita e não cria `-journal`/`-wal` ao lado do arquivo, então o pipeline de negociação
  continua escrevendo sem disputa. Só a stdlib (`sqlite3`). Cache de 5 min por processo, esvaziado
  pelo `=cpLimparCache()` (que agora limpa APIs **e** base).
- **Caminho da base descoberto em runtime** (`config.ResolverTradesDb`): variável **`TRADES_DB_PATH`**
  (override, caminho completo) → senão procura `…\NegociacaoSecundario\code\data\trades.db` a partir
  da pasta do add-in / `CALCCP_DIR` e das pastas-mãe, incluindo um nível de subpastas. Nada de
  caminho pessoal escrito no código (repo público).
- **Taxas em decimal** (7,5983% → `0,075983`), como `=cpTaxa`, `=cpTaxaEmissao` e os indicadores do
  BCB — formate a célula como %. Registros com taxa nula são ignorados.
- **`#N/A` × `ERRO:`**: `#N/A` = papel/indicativo não está na base (resposta legítima); `ERRO: …` =
  base não encontrada ou ilegível (ambiente). O import da camada de base é separado do das APIs: se
  a base sumir, **só** as fórmulas ANBIMA param.
- **Diagnóstico:** `=cpTeste()` e o botão **Sobre** passam a mostrar o estado do `trades.db`
  (caminho em uso e data do último indicativo).

## v2.0.0 — 2026-07-22 — `CalcCP_v2.xlam`: pasta configurável por variável de ambiente (`CALCCP_DIR`)

**Arquivo NOVO `CalcCP_v2.xlam` (o `CalcCP_v1.xlam` fica intocado na pasta).** Como há várias
pessoas usando o v1, não dá pra sobrescrevê-lo (fica travado com o Excel aberto) — então o v2 entra
**do lado**, no modelo de versão por arquivo: quem quiser migra (desmarca o v1, procura o v2); quem
não migrar segue no v1 sem quebrar.

**O que muda no v2:** o `PYTHONPATH` embutido passa de `Z:\CP\CalcCP` (baked no v1, apontava para uma
subpasta inexistente — os arquivos ficam direto em `Z:\CP`) para **`%CALCCP_DIR%;Z:\CP`**: o add-in lê
a pasta da variável de ambiente **`CALCCP_DIR`** (de usuário, sem admin) e cada PC aponta para onde a
pasta estiver de fato — share com letra de drive diferente, cópia local etc.

- **Fórmulas idênticas ao v1:** o `vbaProject.bin` do v2 é byte-a-byte igual ao do v1 (mesmas 32 UDFs
  `cp*`); a única diferença é a config de path. A migração v1→v2 é só re-registrar o `.xlam`.
- **Fallback certo:** quem não setar `CALCCP_DIR` cai em `Z:\CP`, onde os arquivos realmente estão
  (conserta o `Z:\CP\CalcCP` do v1, que não existia).
- **Config manual no banco:** `setx CALCCP_DIR "<pasta>"` (uma linha, por usuário) → reiniciar o Excel.
- **Diagnóstico:** `=cpTeste()` e o botão **Sobre** mostram se a `CALCCP_DIR` foi lida ou se está no fallback.
- **Repo público:** o `.xlam` só carrega `%CALCCP_DIR%` + a letra de drive `Z:\CP` (nenhum nome de servidor).
- **Como foi feito:** `CalcCP_v2.xlam` gerado a partir do v1 original editando o `sharedStrings.xml`
  direto no zip (sem abrir no Excel) → `vbaProject.bin` intacto, sem PII. `VERSION` do `.py` → `2.0.0`
  (o `gerar_bundle.py` nomeia o `.xlam` pelo major → `CalcCP_v2.xlam`). Bundle passa a cuspir o v2.

## v1.1.0 — 2026-07-13 — renomeado para **CalcCP** + nova pasta na share
> ✅ **Em produção no banco desde 14/07/2026** (instalado em `Z:\CP\CalcCP` via bundle, add-in
> re-registrado, confirmado funcionando pelo usuário).
> 📄 Doc (14/07, sem mudança de lógica → sem bump de `VERSION`): README ganhou as 5 fórmulas de
> indicadores que faltavam na tabela (`cpSelic`, `cpCdiAno`, `cpIpca`, `cpIpca15`, `cpIgpm` — existiam
> no código desde a v1.0.0, mas não estavam documentadas) e deixou de mandar dar `git pull` pra
> atualizar (no banco não existe `git pull`; o caminho é o bundle).

**O add-in deixa de se chamar `AntonioOliveiraCalc` e passa a ser `CalcCP`** (nome pessoal num
utilitário usado pelo time todo). **As fórmulas NÃO mudam** — seguem todas com prefixo `cp`
(`=cpPu`, `=cpTaxa`, …), então nenhuma planilha quebra.

| | Antes | Agora |
|---|---|---|
| Pasta na share | `Z:\AntonioOliveira\AntonioOliveiraCalc` | **`Z:\CP\CalcCP`** |
| Add-in (.xlam) | `AntonioOliveiraCalc_v1.xlam` | **`CalcCP_v1.xlam`** |
| Módulo Python | `AntonioOliveiraCalc.py` | **`CalcCP.py`** |
| Bundle | `AntonioOliveiraCalc_bundle.py` | **`CalcCP_bundle.py`** |
| Aba do ribbon | AntonioOliveiraCalc | **CalcCP** |

- **Botão "Sobre" no ribbon** (novo): mostra versão da lógica, nº de fórmulas, pasta, Python e estado
  do import. Substitui o botão "Run" do template do xlwings, que chamava um `main()` inexistente.
  Roda via `RunPython` → também serve de diagnóstico da ponte Excel↔Python.
- **Exige re-registrar o add-in** (Opções → Suplementos → Procurar → `Z:\CP\CalcCP\CalcCP_v1.xlam`).
  Não tem como evitar: o `.xlam` mudou de lugar, e o PYTHONPATH da share é **baked** dentro dele.
- Re-bake feito com o módulo novo (`CallUDF("CalcCP", …)`), 32 UDFs conferidas via COM. `.xlam`
  higienizado: 0 PII, 0 `ItauBBA` (removido o `absPath` que o Excel injeta no save — o repo é público).

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
- `PYTHONPATH` baked no `.xlam`: `Z:\CP\CalcCP`.
- **Guia dos traders `COMO_USAR.html`** (vai no bundle): paleta Itaú, fonte de dados por fórmula
  (B3 / FI Analytics / Banco Central / IBGE / cálculo próprio), câmbio rotulado PTAX de venda.
- **Empacotamento:** um único `.xlam` por versão (`CalcCP_v1.xlam`); a versão é a
  constante `VERSION` no `CalcCP.py` (sem arquivo `VERSION` avulso). O
  `CalcCP_bundle.py` é **gerado** por `gerar_bundle.py` (gitignored — não versionar).
