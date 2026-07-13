# CLAUDE.md — Add-in CalcCP (calculadora de renda fixa via API)

> **Uso e instalação: `README.md`.** Este arquivo é a referência de arquitetura/manutenção (re-bake, xlwings internals).

Documentação técnica para o Claude Code. Leia antes de alterar qualquer coisa nesta pasta.

> ## ⭐ ATUALIZAÇÃO 11/07/2026 — prefixo `cp` + muitas funções novas
> **TODAS as UDFs agora têm o prefixo `cp`** (`=cpPu`, `=cpTaxa`, `=cpDur`, `=cpCdi`). Adicionadas:
> - **Dados do papel:** `cpPupar`, `cpVna`, `cpFluxo` (spill), `cpVencimento`, `cpEmissao`,
>   `cpInicio`, `cpTaxaEmissao`, `cpVne`, `cpAniversario` (extraídos de FI + B3 getBondDetails).
> - **Métricas FI:** `cpGrossUp`, `cpGrossUpTipo`, `cpConvexidade`, `cpDv01`, `cpDurMod`,
>   `cpDiPerc`, `cpSpreadDi`.
> - **Banco Central / IBGE (todas com `[data]` opcional, semântica as-of):** `cpSelic`,
>   `cpSelicOver`, `cpCdiAno`, `cpCdiDia`, `cpIpca`, `cpIpcaAno`, `cpIpca15`, `cpIpcaIndice`
>   (número-índice IBGE SIDRA), `cpIgpm`, `cpIgpDi`, `cpInpc`, `cpPoupanca`, `cpTr`,
>   `cpDolar`, `cpEuro`. **Removidas:** `cpFi`, `cpBond`, `cpBcb`.
> - **Dias úteis (feriados ANBIMA, via di.py):** `cpEhDiaUtil`, `cpDiasUteis`, `cpDiaUtilPosterior`,
>   `cpDiaUtilAnterior`, `cpDiaUtilMaisN`.
> - `apis.py`: `Preco` agora também devolve `vna` (cache compartilhada = performance); novos
>   `BondDetailsB3`, `CampoFi`, `CampoBond`, `FluxoRestante`, `Detalhes`, cliente BCB `BcbSerie`/`BcbValor`
>   (com data as-of), e `IpcaIndice` (número-índice IPCA via IBGE SIDRA).
> - Lista completa e uso: `LEIA-ME.md`.
>
> **RE-BAKE (obrigatório ao renomear/adicionar UDF):** os nomes das funções ficam baked nos
> wrappers VBA do `.xlam`. Trocar só o `.py` NÃO basta. Processo (Excel FECHADO): abrir o `.xlam`
> via COM (`xw.App`), `wb.api.IsAddin=False`, `xlwings.udfs.import_udfs("CalcCP",
> wb.api)`, `IsAddin=True`, `save`. Excel injeta PII (`C:\Users\<usuario>`) no `vbaProject.bin` no save
> → scrub `<usuario>`→`user1` (ascii **e** utf-16). Config (PYTHONPATH) fica na worksheet baked, não no
> `.bin` → dá para transplantar só o `.bin` novo entre DEV (D:\) e PROD (Z:\). Scripts noturnos:
> `scratchpad/rebake.py` + `scratchpad/gerar_bundle.py`; passo a passo em `PROGRESSO_NOTURNO.md`.
> Backups do re-bake: `*.xlam.bak_noturno`.

> **Mapa dos docs desta pasta:** instalar/configurar no banco → `INSTALAR_NO_BANCO.md`; fixar o
> Python por PC → `CONFIGURAR_PYTHON.md`; erro de DLL → `DIAGNOSTICO_DLL.md`; visão geral p/ humano
> → `LEIA-ME.md`; arquitetura/como alterar (este) → `CLAUDE.md`.

## O que é

Add-in de Excel que expõe 5 funções de planilha (UDFs) para precificar títulos de renda fixa.
Depois de instalado, a aba do ribbon aparece como **CalcCP**.
- **Títulos com API** (debênture, CRI/CRA, NTN-B/NTN-F): números vêm das APIs (B3 Calculator →
  FI Analytics, nessa ordem). Sem leitura de base de dados local.
- **DI (tickers `DI1...`)**: NÃO existe API → calculado **localmente** em `di.py` (contagem de dias
  úteis + `feriados_anbima.csv`). O `CalcCP.py` roteia: se `di.EhTickerDi(ticker)` → local,
  senão → APIs.

| UDF no Excel | Retorno |
|---|---|
| `=PU(ticker; "dd/mm/yyyy"; taxa%)` | PU de Operação |
| `=DUR(ticker; "dd/mm/yyyy"; taxa%)` | Duration de Macaulay (anos) |
| `=TAXA(ticker; "dd/mm/yyyy"; pu)` | Taxa de negociação (decimal → formatar como %) |
| `=CDI(dataInicio; dataFim; [percentual])` | Fator de CDI acumulado no período (endpoint público B3 `/di/calculo`); `percentual` = número puro (100 = 100% do CDI, default 100) |
| `=TESTE()` | Diagnóstico ("OK — path: ...") |
| `=LIMPARCACHE()` | Esvazia o cache de respostas das APIs |

> **PU/DUR/TAXA são SÍNCRONAS de propósito** (decorador `@xw.func`, sem `async_mode`). Já foram
> assíncronas (`async_mode='threading'`), mas isso causava **recálculo em loop no SharePoint** — ver
> seção "SharePoint: recálculo em loop" abaixo. **Não** reintroduza async sem entender esse trade-off.

## Arquitetura

É um **add-in customizado do xlwings** (gerado por `xlwings quickstart ... --addin --ribbon`):
- `CalcCP.xlam` embute o VBA do xlwings + um ribbon + o módulo `xlwings_udfs` (as
  "casquinhas" VBA que chamam o Python). É **standalone**: não depende do add-in genérico do xlwings.
- Quando o Excel chama uma UDF, o VBA dispara o Python (interpretador configurado), importa o
  módulo `CalcCP` e executa a função, que faz a chamada HTTP via `apis.py`.
- O add-in **adiciona automaticamente a própria pasta ao PYTHONPATH** — por isso `CalcCP.py`,
  `apis.py` e `config.py` ficam todos juntos do `.xlam`.

### Arquivos
| Arquivo | Papel | Mexe? |
|---|---|---|
| `CalcCP.py` | as UDFs (`@xw.func`). Roteia DI→`di.py`, resto→`apis.py`. | sim — lógica das funções |
| `apis.py` | cliente HTTP B3/FI (urllib, stdlib). Segredos + proxy + cache + normaliza NTN-B/F. | sim — APIs, proxy, parsing |
| `di.py` | cálculo LOCAL de DI (DI1...). Self-contained (lê `feriados_anbima.csv`). | sim — fórmula/calendário DI |
| `feriados_anbima.csv` | calendário de feriados (usado pelo `di.py`). | só ao estender datas |
| `config.py` | só `ENV_PATH` (fallback de dev). | raramente |
| `CalcCP.xlam` | o add-in. Binário. | só ao mudar nome/assinatura de UDF (ver abaixo) |

## Segredos e proxy — via VARIÁVEIS DE AMBIENTE

`apis.py` resolve tudo de variáveis de ambiente (no banco), com fallback opcional num `.env` local
(dev). **Nunca** hardcode valores nos arquivos.

| Para | Variável de ambiente | Fallback `.env` (dev) |
|---|---|---|
| Token B3 Calculator | `token_calc_b3` | `B3_CALC_TOKEN` |
| API key FI Analytics | `token_fianalytics` | `FIANALYTICS_API_KEY` |
| Proxy HTTP | `proxy_http` | (sem proxy = direto) |
| Proxy HTTPS | `proxy_https` | (idem) |

- Resolução de segredo: `_Cred(nomeEnv, chaveArquivo)` = `os.getenv(nome) or CREDENCIAIS.get(chave)`.
- Proxy: `_ConstruirOpener()` monta um `urllib` opener com `ProxyHandler` lendo `proxy_http`/
  `proxy_https`. As 3 chamadas HTTP usam `_opener.open(...)`. urllib trata proxy autenticado em
  HTTPS (move `Proxy-Authorization` para o tunnel CONNECT). Sem env vars → conexão direta.

## Como as APIs funcionam (em `apis.py`)

- **B3 Calculator**: login em `POST /login` com `{"token": <token_calc_b3>}` → devolve um
  `Authorization` (sem "Bearer"). Depois `GET /calcPU/{ticker}/{data}/{taxa}` (→ campo `PU`,
  `PUPar`, `duration`) e `GET /calcYield/{ticker}/{data}/{pu}` (→ campo `yield`). Token cacheado,
  renova em 401. **Tickers de título público**: o `apis.py` normaliza nome amigável → código CETIP
  via `_NormalizarTicker` antes de chamar a B3 (`NTNB35`/`NTN-B 35` → `76019920350515`; ano par→
  ago/15, ímpar→mai/15; `NTNF27` → `95019920270101`). Código cetip já pronto passa inalterado.
  Validado: B3 e FI gov dão o MESMO PU (NTNB35 @7% = 4486,066906).
- **FI Analytics**: `POST` com header `x-api-key: <token_fianalytics>`. Dois endpoints com a MESMA
  resposta: `/deb/debenturecalculator` (debêntures) e `/cr/cricracalculator` (CRI/CRA). O `apis.py`
  tenta o de debênture e, se não vier resultado (ex.: ticker é CRA), tenta o de CRI/CRA
  (`_PostFiAuto`). Resposta é **double-encoded** (JSON dentro de string). Modo `rate` →
  `m2m`/`maculayDuration`; modo `pu` → `m2mRate`.
- Datas para a API: formato `YYYY-MM-DD`. No Excel entram como `dd/mm/yyyy`, data nativa OU serial
  (`TODAY()` aninhado) — `_parse_data` (em `CalcCP.py`) cobre os três; `_EXCEL_EPOCH=30/12/1899`.
- **Roteamento B3→FI** fica em `apis.py`: `Preco()` (taxa→PU) e `TaxaOp()` (PU→taxa). As UDFs
  chamam essas, não `PrecoB3`/`PrecoFi` direto.
- **Cache em memória** por `(origem, ticker, data, taxa/pu)`, inclusive `None`, p/ não martelar a
  rede a cada recálculo.
- **Cache em disco (TTL)** — `_CacheSet`/`_Persistir`/`_CarregarCacheDisco`: só resultados VÁLIDOS
  (nunca `None`/erro) vão p/ `%TEMP%\calcrf_cache.json`, com validade `CACHE_TTL_SEG` (600 s).
  Sobrevive a reabrir o Excel; fora do TTL, refaz a chamada. Escrita atômica, falha de I/O ignorada.
- **Memo de fonte por ticker** — `_fonteTicker` (`Preco`/`TaxaOp`): lembra se B3, FI ou bondbuilder
  respondeu o ticker (só em sucesso) e tenta essa primeiro; SEMPRE mantém o fallback (só reordena).
- **Circuit-breaker por base** — `_Abrir`/`_EmCooldown` (`COOLDOWN_SEG=20`): timeout/erro de REDE
  numa base (B3/FI) marca-a indisponível por 20 s → chamadas seguintes fast-fail sem tocar a rede
  (num storm de recálculo só a 1ª célula paga o timeout). HTTP 400/404 NÃO derruba (base no ar).
- **Token B3** obtido 1×, reusado, renova em 401.
- `=LIMPARCACHE()` limpa tudo: memória, disco, breaker, memo e token.
- **Bondbuilder** (3º fallback, `PrecoBb`/`TaxaBb`): p/ papéis fora de `/deb` e `/cr` (ex.: LCD, LF,
  CDB, CPF). 2 passos — `getuserbonds` (cacheado 1×/processo, casa `bond_name`==ticker → `doc_id`) e
  `/bb/bondbuildercalculator`. É **simétrico** como o `/deb`: manda `rate`→volta `m2m` (PU)+`maculayDuration`;
  manda `pu`→volta `m2mRate`. Cobre `=PU`/`=DUR`/`=TAXA`. Exige a env var **`user_fianalytics`** (e-mail;
  fallback `.env` = `FIANALYTICS_USER`). ⚠️ Usa `http.client` (não urllib) via `_PostFiRaw`: o WAF do
  `/bb` exige o header `x-api-key` em **minúsculo** e o urllib title-caseia → 502. `_PostFiRaw` também
  faz o proxy (CONNECT tunnel). Só dispara depois de B3 e FI falharem (custa 1 `getuserbonds`).

## Como ALTERAR (fluxo de trabalho)

### Mudar lógica/parsing/proxy (caso comum) — NÃO precisa mexer no `.xlam`
1. Edite `apis.py` ou `CalcCP.py`.
2. Teste por fora (sem Excel):
   ```
   python -c "import CalcCP as m; print(m.PU('FGEN13','13/06/2025',0.064686))"
   ```
   (≈ 961,70 para esse caso).
3. O usuário pega a mudança ao **reabrir o Excel** (o Python recarrega o módulo no novo processo).

### Adicionar/renomear UDF, mudar argumentos, ou trocar sync↔async — precisa RE-GERAR o `.xlam`
O módulo `xlwings_udfs` (as casquinhas VBA) é **baked no `.xlam`** e difere conforme a assinatura e o
`async_mode` das funções (ver `udfs.py`, `generate_vba_wrapper`, ~linha 515). Então mudar isso exige
regenerar o `xlwings_udfs`. Duas formas:

**(1) Manual, no Excel:** com o add-in carregado, **Alt+F11** → módulo `xlwings` do projeto
`CalcCP.xlam` → rode a sub **`ImportPythonUDFsToAddin`** (F5) → **Ctrl+S**. Requer "Confiar no
acesso ao modelo de objeto do VBA" ligado.

**(2) Via COM, do Python (sem Alt+F11) — foi como o re-bake síncrono foi feito (01/07):**
chama-se a MESMA função que a sub VBA chama (`xlwings.udfs.import_udfs`). Esqueleto:
```python
import sys, xlwings as xw
from xlwings import udfs
sys.path.insert(0, r"<pasta com CalcCP.py>")   # p/ import_udfs achar o modulo
app = xw.App(visible=True, add_book=False); app.display_alerts = False
wb = app.books.open(r"<...>\CalcCP.xlam")
wb.api.IsAddin = False        # ⚠️ senao MacroOptions falha: "Cannot edit a macro on a hidden workbook"
udfs.import_udfs("CalcCP", wb.api)   # = ImportPythonUDFsToAddin (addin:=True -> ThisWorkbook)
wb.api.IsAddin = True; wb.save(); wb.close(); app.quit()
```
Só o `xl/vbaProject.bin` muda.

> ⚠️ **PII ao salvar no Excel:** o save injeta `C:\Users\<voce>\...` no `vbaProject.bin` (ascii **e**
> utf-16-le) e o seu nome no `docProps`. Como o repo é público (ver regras abaixo), **scrub antes de
> commitar**: substitua `<seu-usuario>`→`user1` no `vbaProject.bin` (mesmo tamanho, preserva offsets).
> Estratégia usada p/ o PROD: re-bakear num `.xlam` de dev (config local), depois **transplantar só o
> `vbaProject.bin` síncrono+scrubbado** para o `CalcCP.xlam` de produção (que mantém o
> `docProps=CalcCP` e a config `Z:\` limpos). Assim o PROD nunca reabre no Excel → não pega PII nova.

3. Distribua o `.xlam` novo (commit + o banco dá `git pull`). No banco, reabrir o Excel basta.

## Planilha do SharePoint/OneDrive CONGELA o Excel ao usar uma UDF

> ✅ **JÁ APLICADO neste `.xlam`** (30/06): o sheet `myaddin.conf` embutido tem
> `ADD_WORKBOOK_TO_PYTHONPATH=false` e `PYTHONPATH=Z:\CP\CalcCP`. Funciona com o
> add-in carregado **desse caminho fixo**. Se a pasta mudar, edite o valor de `PYTHONPATH` no
> sheet `myaddin.conf` (ou via o script `scratchpad/bake_config.py`). Abaixo, a explicação/teoria.

**Sintoma:** numa planilha aberta do SharePoint/OneDrive (cujo `Workbook.FullName` é uma URL
`https://...`), qualquer fórmula do add-in trava o Excel. Em planilha local, funciona normal.

**Causa:** com `ADD_WORKBOOK_TO_PYTHONPATH` ligado (padrão), o xlwings (`prepare_sys_path` →
`fullname_url_to_local_path`) tenta **mapear a URL do arquivo para um caminho local**, e essa busca
no sistema de arquivos trava.

**Fix (2 chaves de config do xlwings):**
```
ADD_WORKBOOK_TO_PYTHONPATH = false   # não tenta resolver a URL → não trava
PYTHONPATH = <pasta do add-in>       # re-adiciona a pasta dos .py (o item acima a removeria)
```

**Onde colocar — duas formas:**

1. **Por usuário** — arquivo `%USERPROFILE%\.myaddin\myaddin.conf` (PROJECT_NAME do addin =
   `myaddin`), formato `"Chave","Valor"`:
   ```
   "ADD_WORKBOOK_TO_PYTHONPATH","false"
   "PYTHONPATH","C:\caminho\para\CalcCP"
   ```
   Implantar via login script / GPO, ou criar à mão. Simples, mas é um arquivo por usuário.

2. **Para todos (rollout)** — embutir essas 2 linhas num **sheet `myaddin.conf` dentro do `.xlam`**
   (precisa abrir o `.xlam` no Excel uma vez e adicionar o sheet). Para o `PYTHONPATH` valer em
   qualquer máquina, use uma **variável de ambiente** (o xlwings expande com `os.path.expandvars`):
   `"PYTHONPATH","%CALCRF_DIR%"`, e cada máquina seta `CALCRF_DIR` = pasta do add-in (junto com os
   demais env vars). Assim o `.xlam` é uniforme e zero config por usuário.

> Numa **share de rede** (UNC igual p/ todos), pode-se fixar o caminho direto no `PYTHONPATH`
> (sem env var), já que é o mesmo para todos.

Reiniciar o Excel após qualquer um dos dois.

> ⚠️ **Precedência da config (importa muito):** o VBA (`GetConfig` em `xlwings_custom_addin.bas`)
> lê nesta ordem: sheet do workbook ativo → **sheet embutido do `.xlam`** → arquivo do diretório →
> arquivo do usuário `%USERPROFILE%\.myaddin\myaddin.conf`. Uma chave que **existe** no sheet
> embutido (mesmo com valor vazio) **curto-circuita** e ignora o arquivo do usuário. Consequência
> prática: o sheet embutido traz `Interpreter_Win` **vazio**, então para fixar o Python por PC via
> arquivo do usuário use a chave legada **`Interpreter`** (que não está no sheet) — não
> `Interpreter_Win`. Ver `CONFIGURAR_PYTHON.md`.

## SharePoint: recálculo em LOOP (célula piscando / Excel travado)

> ✅ **JÁ CORRIGIDO (01/07):** as UDFs viraram síncronas. Esta seção explica o porquê — não é o mesmo
> problema do "congela ao abrir" acima (aquele era resolução de URL do PYTHONPATH).

**Sintoma:** numa planilha do SharePoint, `=PU/=DUR/=TAXA` **funcionam** quando os argumentos são
digitados; mas quando o ticker (ou a data) vem de uma **fórmula viva** (ex.: `XLOOKUP`), a célula fica
**recalculando sem parar** (valor piscando) e o Excel trava.

**Causa:** as funções eram **assíncronas** (`async_mode='threading'`). Async roda em thread e
**escreve o resultado de volta** na célula; essa escrita conta como mudança e **dispara um
recálculo**. Com argumento constante a cadeia estabiliza; mas com um `XLOOKUP` alimentando o argumento,
a escrita re-dispara o XLOOKUP → re-dispara a UDF → escreve de novo → **loop infinito**. O cache do
`apis.py` não ajuda porque o que fica em loop é o *recálculo*, não a rede.

**Correção:** tornar PU/DUR/TAXA **síncronas** (remover `async_mode` → `@xw.func`). Sem write-back, não
há loop; e o cache mantém rápido (rede só na 1ª vez por input). Exigiu re-gerar o `.xlam` (ver "Como
ALTERAR" (2)). Paliativo enquanto não atualiza: cálculo em **Manual (F9)**.

## ⚠️ REGRAS DESTA PASTA (deploy para o banco)

1. **NENHUM arquivo `.bat`** nesta pasta. Nunca crie scripts `.bat` aqui.
2. **NENHUM dado sensível** em arquivo: nada de tokens, senhas, e-mails, caminhos pessoais,
   identificadores de máquina/usuário. Segredos vêm SÓ de variáveis de ambiente. Antes de
   entregar/alterar, varra a pasta atrás de valores reais de credenciais e remova.
3. Não reintroduza dependência de base local (trades.db) nem a calculadora ANBIMA. **Exceção:** o
   DI é calculado localmente em `di.py` (não existe API de DI) — isso é intencional e permitido.
4. `apis.py` usa só a stdlib (urllib). Evite adicionar dependências (mantém o deploy leve).

## Teste rápido (sanity)
```
python -c "import CalcCP as m; print(m.TESTE()); print(m.PU('FGEN13','13/06/2025',0.064686))"
```
Esperado: `OK — path: <esta pasta>` e `961.699686` (com as env vars/`.env` resolvendo os segredos).
