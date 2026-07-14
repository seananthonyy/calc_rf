# CalcCP — Add-in de Renda Fixa para Excel

Add-in de planilha (UDFs) para precificar renda fixa e puxar indicadores. Todas as funções têm o
prefixo **`cp`** (ex.: `=cpPu`). As fontes são as APIs **B3 Calculator**, **FI Analytics** e
**Banco Central (SGS)/IBGE**; contratos de **DI futuro** (`DI1F27`…) são calculados **localmente**.
Datas aceitam `"dd/mm/aaaa"`, célula de data ou `HOJE()`.

> **Documentação:** este `README.md` cobre uso, instalação e operação. Detalhes de arquitetura,
> re-bake do `.xlam` e internals do xlwings estão no `CLAUDE.md` (referência de desenvolvimento).
> Para os **traders** (guia visual, sem jargão, com a fonte de cada fórmula), abrir o
> **`COMO_USAR.html`** no navegador.

---

## Índice
1. [Pré-requisitos do PC](#pré-requisitos-do-pc)
2. [Instalar no banco](#instalar-no-banco)
3. [Referência de funções](#referência-de-funções)
4. [Versionamento e atualização](#versionamento-e-atualização)
5. [Solução de problemas](#solução-de-problemas)

---

## Pré-requisitos do PC

1. **Excel 64-bit** (Arquivo → Conta → Sobre o Excel).
2. **Python 64-bit** com **xlwings exatamente 0.36.6** e a DLL `xlwings64-0.36.6.dll` na raiz do
   Python (ao lado do `pythonw.exe`). Instalação padrão do python.org já põe a DLL na raiz. Verificar:
   ```
   python -c "import xlwings,os,sys,glob; r=os.path.dirname(sys.executable); print(xlwings.__version__); print(glob.glob(os.path.join(r,'xlwings*.dll')))"
   ```
   Esperado: `0.36.6` e a lista com `xlwings64-0.36.6.dll`. Se não, ver [erro de DLL](#erro-file-not-found-xlwings64-0366dll).
3. **A pasta do add-in na share, em `Z:\CP\CalcCP`** — é o `PYTHONPATH`
   embutido no `.xlam` (tem que ser ESSE caminho). Conteúdo mínimo:
   `CalcCP_v1.xlam`, `CalcCP.py`, `apis.py`, `di.py`, `config.py`,
   `feriados_anbima.csv`. Puxar do repositório do projeto (ou extrair o `CalcCP_bundle.py`).

---

## Instalar no banco

### 1. Colocar a pasta na share
Garanta que `Z:\CP\CalcCP` existe com os arquivos acima. A share precisa
estar acessível quando o Excel abre. (Alternativa rápida: jogue só o `CalcCP_bundle.py`
na pasta e rode `python CalcCP_bundle.py` — ele extrai tudo.)

> **Atualizar depois** = baixar o bundle novo na share e rodar de novo. Ele **pula todo arquivo que
> já está idêntico**, então o `.xlam` — que só muda quando há UDF/assinatura nova — não é reescrito e
> **não precisa que ninguém feche o Excel**: dá pra atualizar a lógica (`.py`) com o time usando a
> planilha. Se algum arquivo tiver MESMO mudado e estiver travado pelo Excel, ele avisa qual e sai
> com erro (aí sim, fechar o Excel e repetir).

### 2. Variáveis de ambiente do usuário (segredos + proxy)
No banco não há `.env`; os segredos entram por variável de ambiente **do usuário**:

| Variável | Conteúdo |
|---|---|
| `token_calc_b3` | token do B3 Calculator |
| `token_fianalytics` | API key do FI Analytics |
| `user_fianalytics` | e-mail do usuário FI (só p/ o fallback bondbuilder; opcional) |
| `proxy_http` | `http://USUARIO:SENHA@HOST:PORTA` |
| `proxy_https` | `http://USUARIO:SENHA@HOST:PORTA` |

Sem elas, `=cpTeste()` dá "OK" mas `=cpPu` retorna `ERRO: APIs sem resposta (B3/FI)`.

### 3. Liberar acesso ao VBA (uma vez)
Excel → Opções → **Central de Confiabilidade** → Configurações de Macro → marcar
**"Confiar no acesso ao modelo de objeto de projeto do VBA"**.

### 4. Habilitar o add-in
Excel → Opções → **Suplementos** → "Suplementos do Excel" → **Ir...** → **Procurar** →
selecionar `Z:\CP\CalcCP\CalcCP_v1.xlam` → OK.
As UDFs já vêm **registradas** no `.xlam` (não precisa "Import Functions"). Aparece a aba
**CalcCP** no ribbon, com o botão **Sobre** — que mostra a versão da lógica, a pasta de onde o
add-in está lendo, o Python em uso e o estado do import. Como ele roda o Python de verdade
(`RunPython`), serve também de diagnóstico: se a ponte estiver quebrada, o erro aparece ali.

### 5. (Só se o PC tiver vários Python) Fixar o Python certo
Criar `%USERPROFILE%\.myaddin\myaddin.conf` com **uma linha**:
```
"Interpreter","C:\caminho\do\pythonw.exe"
```
- Use a chave **`Interpreter`** (NÃO `Interpreter_Win` — essa vem vazia na planilha embutida do
  `.xlam`, tem precedência e seria ignorada no arquivo do usuário).
- Aponte para o **`pythonw.exe`** (com `w`) do Python que tem xlwings 0.36.6.
- **Uma linha só, sem linha em branco final** (um `\n` extra dá erro 62 "Input past end of file" e
  quebra todas as UDFs). Encoding ANSI/ASCII, sem BOM. Snippet PowerShell seguro:
  ```powershell
  $py = "C:\caminho\do\pythonw.exe"
  $dir = "$env:USERPROFILE\.myaddin"; New-Item -ItemType Directory -Force $dir | Out-Null
  [System.IO.File]::WriteAllText("$dir\myaddin.conf", "`"Interpreter`",`"$py`"`r`n", [System.Text.Encoding]::ASCII)
  ```
Esse arquivo é **por PC e por usuário** (não vai pro repositório).

### 6. Testar
Feche e reabra o Excel:
```
=cpTeste()                              -> OK v1.1.0 — path: Z:\CP\CalcCP
=cpPu("FGEN13"; "13/06/2025"; 6,4686%)  -> ~961,70
=cpPu("DI1F27"; "01/07/2026"; 10%)      -> ~95310,20   (DI, cálculo local)
```

### Checklist (por PC)
- [ ] Pasta em `Z:\CP\CalcCP` atualizada
- [ ] Python 64-bit com xlwings **0.36.6** + DLL na raiz
- [ ] `token_calc_b3`, `token_fianalytics`, `proxy_http`, `proxy_https` setados
- [ ] "Confiar no acesso ao modelo de objeto de projeto do VBA" ligado
- [ ] `.xlam` habilitado (aba **CalcCP** aparece)
- [ ] (se multi-Python) `myaddin.conf` com a chave `Interpreter`
- [ ] `=cpPu("FGEN13";"13/06/2025";6,4686%)` ≈ 961,70

---

## Referência de funções

### Precificação
| Função | Retorno |
|---|---|
| `=cpPu(ticker; data; taxa%)` | PU de Operação (taxa como % do Excel, ex. 6,4618%) |
| `=cpTaxa(ticker; data; pu)` | Taxa de negociação (decimal → formate como %) |
| `=cpDur(ticker; data; taxa%)` | Duration de Macaulay em anos |
| `=cpCdi(dataIni; dataFim; percentual)` | Fator de CDI acumulado (percentual = número puro, 100 = 100%) |

### Dados do papel (taxa OPCIONAL em PU Par/VNA/Fluxo — usa a de emissão se omitida)
| Função | Retorno |
|---|---|
| `=cpPupar(ticker; data; [taxa%])` | PU Par (valor nominal atualizado + juros) |
| `=cpVna(ticker; data; [taxa%])` | VNA — Valor Nominal Atualizado |
| `=cpFluxo(ticker; data; [taxa%])` | Agenda **RESTANTE** a partir da data (spill): Data · Tipo · %Amort · %Incorp |
| `=cpFluxoCompleto(ticker)` | Agenda **INTEIRA**, desde a emissão (spill): Data · %Amort · %Incorp |
| `=cpVencimento(ticker)` | Data de vencimento |
| `=cpEmissao(ticker)` | Data de emissão |
| `=cpInicioRentabilidade(ticker)` | Data de início de rentabilidade |
| `=cpTaxaEmissao(ticker)` | Taxa de emissão em **decimal** (6,4618% → 0,064618; %DI → 1,135) |
| `=cpVne(ticker)` | Valor Nominal de Emissão |
| `=cpAniversario(ticker)` | Dia de aniversário |

> `cpFluxo` × `cpFluxoCompleto`: as duas vêm da **agenda cadastrada** na B3 (`getBondDetails`) e trazem
> uma linha por data de evento — a data de amortização, que a B3 manda em 2 eventos (o `A` e o `J`),
> é agregada numa linha só. A diferença é o **escopo**: `cpFluxo` mostra só o que **ainda vai
> acontecer** a partir da data informada (e traz a coluna `Tipo`); `cpFluxoCompleto` mostra a agenda
> **inteira**, desde a emissão. Datas de cupom puro saem com `%Amort`=0 e `%Incorp`=0.
>
> `Tipo` (só no `cpFluxo`): **J** = só juros · **J+A** = juros e amortização · **J+I** = juros e
> incorporação (incorporação só existe em papel `IPCA-I`).
>
> O argumento `taxa` do `cpFluxo` é aceito por compatibilidade e **ignorado** (a agenda não depende
> da taxa). As colunas antigas `Prazo(DU)`·`VF`·`VP` (fluxo calculado em R$) **não existem mais**.
>
> **Principal do vencimento:** nos papéis `IPCA-I` a B3 não cadastra o principal residual quitado no
> vencimento (FGEN13: os `A` somam 91,975%). A linha do vencimento é completada com `100 − Σ%amort`,
> então a agenda **sempre fecha 100%**. É valor derivado — validado contra o `trades.db` (FGEN13
> 8,025 e MESA13 13,176856, idênticos).

### Métricas / gross up (FI Analytics)
| Função | Retorno |
|---|---|
| `=cpGrossUp(ticker; data; [taxa%])` | Taxa equivalente após imposto |
| `=cpGrossUpTipo(ticker; data; [taxa%])` | `GROSS_UP` (isento) ou `AFTER_TAX` (tributado) |
| `=cpDv01(ticker; data; [taxa%])` | DV01 (variação do PU por 1 bp) |

### Indicadores — Banco Central / IBGE (data OBRIGATÓRIA = as-of)
**Data obrigatória** — retorna o último valor publicado **até** a data (cobre
fim de semana/feriado e séries mensais). Ex.: `=cpSelic("31/12/2020")` → 2,0.

**Taxas saem em DECIMAL** (6% → 0,06) — formate a célula como % se quiser ver "6%".

| Função | Retorno | Fonte |
|---|---|---|
| `=cpSelic(data)` | Meta SELIC a.a., **decimal** | BCB série 432 |
| `=cpCdiAno(data)` | CDI anualizado base 252, **decimal** | BCB série 4389 |
| `=cpIpca(data)` | IPCA **do mês**, decimal (0,93% → 0,0093) | BCB série 433 |
| `=cpIpca15(data)` | IPCA-15 do mês, decimal | BCB série 7478 |
| `=cpIgpm(data)` | IGP-M do mês, decimal | BCB série 189 |
| `=cpIpcaIndice([data])` | IPCA **número-índice** (base dez/1993 = 100) — valor natural, não é taxa | IBGE SIDRA |
| `=cpDolar([data])` / `=cpEuro([data])` | Câmbio PTAX venda (R$) — valor natural | BCB |

> **Correção IPCA entre datas** = `=cpIpcaIndice(dataFim)/cpIpcaIndice(dataIni)` — mesmo fator que
> atualiza o VNA de uma NTN-B/papel IPCA+. Use o **número-índice** pra isso, não o `cpIpca` (que é a
> variação de um mês só).

### Dias úteis (feriados ANBIMA)
| Função | Retorno |
|---|---|
| `=cpEhDiaUtil(data)` | VERDADEIRO/FALSO |
| `=cpDiasUteis(dataIni; dataFim)` | Nº de dias úteis (início inclusive, fim exclusive) |
| `=cpDiaUtilPosterior(data)` / `=cpDiaUtilAnterior(data)` | Próximo / anterior dia útil |
| `=cpDiaUtilMaisN(data; n)` | Data ± n dias úteis |

### Diagnóstico
| Função | Retorno |
|---|---|
| `=cpTeste()` | "OK v{versão} — path: …" ou o erro de import |
| `=cpLimparCache()` | Esvazia o cache das APIs e força novas chamadas |

**Observações:** erros aparecem como texto `ERRO: …` na célula. O 1º cálculo de um papel bate na
rede (~1–5 s); os demais vêm do cache (disco, 10 min) — `=cpLimparCache()` força atualização.

---

## Versionamento e atualização

**Regra de ouro:** a **lógica** (`.py`) é retrocompatível — **só adiciona** UDF, **nunca remove nem
renomeia** (senão quebra planilhas e `.xlam` antigos). O **`.xlam` é versionado por NOME de arquivo**.

| Mudou | O que fazer | Efeito nos usuários |
|---|---|---|
| **Lógica** (`apis.py`, `CalcCP.py`, `di.py`) | **no banco não há `git pull`** → baixar o `CalcCP_bundle.py` do GitHub na share e rodar (`python CalcCP_bundle.py`) | reabrir o Excel — **nada quebra**, nem precisa fechar (o `.xlam` idêntico é pulado) |
| **Nova UDF / assinatura** (precisa re-bake) | gerar `CalcCP_v{N+1}.xlam` (arquivo **novo**), atualizar `VERSION` no `.py`, rodar `python gerar_bundle.py`, registrar `CHANGELOG.md` | quem quiser o novo registra o `_vN+1`; **o `_vN` antigo fica na share e continua funcionando** |

> Por isso os `.xlam` **nunca são apagados/sobrescritos**: cada versão é um arquivo novo
> (`_v1`, `_v2`, …). Assim ninguém no banco fica com o add-in quebrado, e você nunca precisa deletar
> um `.xlam` que esteja em uso (travado pelo Excel). 99% das atualizações são só **colar um `.py` e
> reabrir o Excel** (sem tocar no `.xlam`). Detalhe do re-bake em `CLAUDE.md`.

---

## Solução de problemas

| Sintoma na célula | Causa provável | Solução |
|---|---|---|
| `file not found: xlwings64-0.36.6.dll` | Python errado / xlwings ≠ 0.36.6 / DLL ausente | ver abaixo |
| fórmulas não calculam, PC com vários Python | add-in pegou o Python errado | passo 5 (fixar `Interpreter`) |
| `ERRO: APIs sem resposta (B3/FI)` | faltam env vars / proxy | passo 2 |
| `Input past end of file` | `myaddin.conf` com linha em branco sobrando | recriar com o snippet do passo 5 |
| aba não aparece / `#NAME?` | add-in não habilitado ou VBA não liberado | passos 3 e 4 |
| Excel congela/pisca em planilha do SharePoint | resolução de URL do PYTHONPATH / write-back de UDF assíncrona | já mitigado no `.xlam` (config + UDFs síncronas); paliativo: cálculo Manual/F9. Ver `CLAUDE.md` |

### Erro `file not found: xlwings64-0.36.6.dll`
O VBA carrega `xlwings64-0.36.6.dll` **da pasta do `pythonw.exe`**. Diagnóstico:
```
where pythonw
python -c "import xlwings,os,sys,glob; r=os.path.dirname(sys.executable); print('versao:',xlwings.__version__); print('raiz:',r); print('DLLs:',glob.glob(os.path.join(r,'xlwings*.dll'))); print('pkg:',os.path.dirname(xlwings.__file__))"
python -c "import struct; print(struct.calcsize('P')*8,'bit')"
```
- **Versão ≠ 0.36.6** → `pip install --force-reinstall --proxy http://USER:SENHA@HOST:PORTA "xlwings==0.36.6"`.
- **`pythonw` da Microsoft Store** (`...\WindowsApps\`) ou Python sem a DLL → fixar o Python real via
  `myaddin.conf` (passo 5), usando o Python onde `xlwings==0.36.6` está com a DLL na raiz.
- **DLL só em `site-packages\xlwings`** → copiar `xlwings64-0.36.6.dll` para a raiz do Python.
- **Bitness**: Excel e Python têm que ser a mesma arquitetura (aqui, 64-bit).

Depois de qualquer correção, **feche e reabra o Excel** e teste `=cpTeste()` + `=cpPu(...)`.
