# CLAUDE.md — Add-in CalcRF (calculadora de renda fixa via API)

Documentação técnica para o Claude Code. Leia antes de alterar qualquer coisa nesta pasta.

## O que é

Add-in de Excel que expõe 3 funções de planilha (UDFs) para precificar títulos de renda fixa,
**buscando os números exclusivamente nas APIs** (B3 Calculator → FI Analytics, nessa ordem de
prioridade). **Não há cálculo local e não há leitura de nenhuma base de dados** — só chamadas HTTP.

| UDF no Excel | Retorno |
|---|---|
| `=PU(ticker; "dd/mm/yyyy"; taxa%)` | PU de Operação |
| `=DUR(ticker; "dd/mm/yyyy"; taxa%)` | Duration de Macaulay (anos) |
| `=TAXA(ticker; "dd/mm/yyyy"; pu)` | Taxa de negociação (decimal → formatar como %) |
| `=TESTE()` | Diagnóstico ("OK — path: ...") |
| `=LIMPARCACHE()` | Esvazia o cache de respostas das APIs |

## Arquitetura

É um **add-in customizado do xlwings** (gerado por `xlwings quickstart ... --addin --ribbon`):
- `calcrf_addin.xlam` embute o VBA do xlwings + um ribbon + o módulo `xlwings_udfs` (as
  "casquinhas" VBA que chamam o Python). É **standalone**: não depende do add-in genérico do xlwings.
- Quando o Excel chama uma UDF, o VBA dispara o Python (interpretador configurado), importa o
  módulo `calcrf_addin` e executa a função, que faz a chamada HTTP via `apis.py`.
- O add-in **adiciona automaticamente a própria pasta ao PYTHONPATH** — por isso `calcrf_addin.py`,
  `apis.py` e `config.py` ficam todos juntos do `.xlam`.

### Arquivos
| Arquivo | Papel | Mexe? |
|---|---|---|
| `calcrf_addin.py` | as UDFs (decoradas com `@xw.func`). Importa de `apis.py`. | sim — lógica das funções |
| `apis.py` | cliente HTTP B3/FI (urllib, stdlib). Segredos + proxy + cache. | sim — APIs, proxy, parsing |
| `config.py` | só `ENV_PATH` (fallback de dev). | raramente |
| `calcrf_addin.xlam` | o add-in. Binário. | só ao mudar nome/assinatura de UDF (ver abaixo) |

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
  renova em 401.
- **FI Analytics**: `POST` com header `x-api-key: <token_fianalytics>`. Dois endpoints com a MESMA
  resposta: `/deb/debenturecalculator` (debêntures) e `/cr/cricracalculator` (CRI/CRA). O `apis.py`
  tenta o de debênture e, se não vier resultado (ex.: ticker é CRA), tenta o de CRI/CRA
  (`_PostFiAuto`). Resposta é **double-encoded** (JSON dentro de string). Modo `rate` →
  `m2m`/`maculayDuration`; modo `pu` → `m2mRate`.
- Datas para a API: formato `YYYY-MM-DD`. No Excel entram como `dd/mm/yyyy` ou data nativa.
- Cache em memória por `(origem, ticker, data, taxa/pu)`, inclusive resultados `None`, para não
  martelar a rede a cada recálculo. `=LIMPARCACHE()` limpa.

## Como ALTERAR (fluxo de trabalho)

### Mudar lógica/parsing/proxy (caso comum) — NÃO precisa mexer no `.xlam`
1. Edite `apis.py` ou `calcrf_addin.py`.
2. Teste por fora (sem Excel):
   ```
   python -c "import calcrf_addin as m; print(m.PU('FGEN13','13/06/2025',0.064686))"
   ```
   (≈ 961,70 para esse caso).
3. O usuário pega a mudança ao **reabrir o Excel** (o Python recarrega o módulo no novo processo).

### Adicionar/renomear UDF ou mudar argumentos — precisa reimportar no `.xlam`
1. Edite as funções em `calcrf_addin.py` (mantenha o decorador `@xw.func`).
2. No Excel, com o add-in carregado: **Alt+F11** → módulo `xlwings` do projeto `calcrf_addin.xlam`
   → rode a sub **`ImportPythonUDFsToAddin`** (F5). Isso regenera o módulo `xlwings_udfs`.
3. **Salve o `.xlam`** (Ctrl+S no editor VBA). Distribua o `.xlam` novo.
   Requer "Confiar no acesso ao modelo de objeto do VBA" ligado.

## Planilha do SharePoint/OneDrive CONGELA o Excel ao usar uma UDF

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
   "PYTHONPATH","C:\caminho\para\calcrf_addin"
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

## ⚠️ REGRAS DESTA PASTA (deploy para o banco)

1. **NENHUM arquivo `.bat`** nesta pasta. Nunca crie scripts `.bat` aqui.
2. **NENHUM dado sensível** em arquivo: nada de tokens, senhas, e-mails, caminhos pessoais,
   identificadores de máquina/usuário. Segredos vêm SÓ de variáveis de ambiente. Antes de
   entregar/alterar, varra a pasta atrás de valores reais de credenciais e remova.
3. Não reintroduza dependência de base local (trades.db) nem cálculo local — o add-in é só API.
4. `apis.py` usa só a stdlib (urllib). Evite adicionar dependências (mantém o deploy leve).

## Teste rápido (sanity)
```
python -c "import calcrf_addin as m; print(m.TESTE()); print(m.PU('FGEN13','13/06/2025',0.064686))"
```
Esperado: `OK — path: <esta pasta>` e `961.699686` (com as env vars/`.env` resolvendo os segredos).
