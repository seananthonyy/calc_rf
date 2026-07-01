# Diagnóstico — erro "file not found: xlwings64-0.36.6.dll"

> Doc autocontido para o Claude Code do banco. Trata **um erro específico** ao usar o add-in
> CalcRF no Excel. Para a arquitetura geral do add-in, ver `CLAUDE.md` nesta mesma pasta.

## Sintoma
Ao usar qualquer UDF do add-in (ex.: `=TESTE()`, `=PU(...)`), o Excel retorna:
```
file not found: xlwings64-0.36.6.dll
```
(ou "Could not load xlwings64-0.36.6.dll").

> ⚠️ Não confundir com `ERRO: APIs sem resposta (B3/FI)` — esse é **outro** problema (variáveis de
> ambiente / proxy), não a DLL.

## Causa raiz (confirmada lendo o VBA do add-in)
O add-in é um **xlwings custom add-in versão 0.36.6**. O VBA dele carrega a DLL com `LoadLibrary`
**da pasta do interpretador** — concretamente (xlwings_custom_addin.bas, ~linha 496):
```vba
If LoadLibrary(ParentFolder(PYTHON_WIN) + "\" + "xlwings64-0.36.6.dll") = 0 Then  ' Standard
    If LoadLibrary(ParentFolder(ParentFolder(PYTHON_WIN)) + "\" + "xlwings64-0.36.6.dll") = 0 ...
```
Ou seja: a DLL **precisa estar na pasta do `pythonw.exe`** (a raiz do Python), com o nome
**exato** `xlwings64-0.36.6.dll`. Essa DLL é instalada pelo pacote pip `xlwings` e, numa instalação
padrão (python.org), cai justamente na raiz do Python (ao lado do `pythonw.exe`), **não** em
`site-packages`.

O erro acontece quando:
- **(A)** a versão instalada do xlwings **não é exatamente 0.36.6** → a DLL na raiz tem outro nome
  (ex.: `xlwings64-0.37.0.dll`), e o VBA procura `...0.36.6`; **ou**
- **(B)** o `pythonw` que o add-in usa aponta para uma pasta **sem** a DLL — clássico: o `pythonw`
  da **Microsoft Store** (fica em `...\WindowsApps\`, que não tem a DLL); **ou**
- **(C)** distribuição atípica de Python em que a DLL não foi para a raiz.

### Layout de referência (máquina que FUNCIONA)
- `xlwings.__version__ == 0.36.6`
- Na pasta do `pythonw.exe` (= `os.path.dirname(sys.executable)`) existem:
  `xlwings64-0.36.6.dll` **e** `xlwings32-0.36.6.dll`.
- Em `site-packages\xlwings\` **NÃO** há a DLL.
- Excel 64-bit ↔ Python 64-bit (a DLL é `xlwings64`).

## Comandos de diagnóstico (rodar no terminal do banco)
```bash
# 1) Qual pythonw está no PATH (o add-in usa este por padrão)
where pythonw

# 2) Versão do xlwings + raiz do Python + DLLs presentes na raiz
python -c "import xlwings,os,sys,glob; r=os.path.dirname(sys.executable); print('versao:',xlwings.__version__); print('raiz:',r); print('DLLs na raiz:',glob.glob(os.path.join(r,'xlwings*.dll'))); print('pkg:',os.path.dirname(xlwings.__file__))"

# 3) Bitness do Python (precisa casar com o Excel; aqui esperamos 64-bit)
python -c "import struct; print(struct.calcsize('P')*8, 'bit')"
```
Interpretação:
- Se a versão **≠ 0.36.6** ou não houver `xlwings64-0.36.6.dll` na raiz → **Caso A**.
- Se `where pythonw` apontar para `...\WindowsApps\` (Store) ou uma pasta diferente da raiz do
  Python real → **Caso B**.
- Se a DLL existir só em `site-packages\xlwings` e não na raiz → **Caso C**.

## Correções

### Caso A — fixar a versão exata 0.36.6
```bash
pip install --force-reinstall --proxy http://USUARIO:SENHA@HOST:PORTA "xlwings==0.36.6"
```
(o pip precisa do proxy do banco para baixar). Confirme com o comando de diagnóstico (2) que
`xlwings64-0.36.6.dll` apareceu na raiz e a versão é 0.36.6.

### Caso B — apontar o add-in para o Python real
Crie/edite o arquivo `%USERPROFILE%\.myaddin\myaddin.conf` com:
```
"Interpreter","C:\caminho\real\PythonXX\pythonw.exe"
```
Use o **mesmo** Python onde o `xlwings==0.36.6` está instalado (o da raiz que tem a DLL).
(O nome `myaddin` é o PROJECT_NAME interno do add-in; é proposital.)

> ⚠️ **Use a chave `Interpreter`, NÃO `Interpreter_Win`.** O `.xlam` traz `Interpreter_Win` **vazio**
> na planilha de config embutida, e essa planilha tem **precedência** sobre o arquivo do usuário —
> então `Interpreter_Win` no arquivo seria ignorado. A chave legada `Interpreter` não está na
> planilha embutida, então é a que "passa" para o arquivo do usuário. (No VBA: `INTERPRETER_WIN`
> vazio → cai em `GetConfig("INTERPRETER","pythonw")`, que lê o arquivo; chave é case-insensitive.)
> Detalhes e snippet seguro (encoding/sem newline final) em `CONFIGURAR_PYTHON.md`.

### Caso C — copiar a DLL para a raiz do Python
```bash
copy "<site-packages>\xlwings\xlwings64-0.36.6.dll" "<pasta do pythonw.exe>\"
```
(descubra `<site-packages>\xlwings` no campo `pkg:` do diagnóstico; e a pasta do pythonw no `raiz:`).

### Bitness (se aplicável)
Se o Excel for 32-bit, o erro seria `xlwings32-...`. Garanta Excel e Python com a **mesma**
arquitetura. Aqui o erro é `xlwings64` → Excel 64-bit → use Python 64-bit.

## Depois de qualquer correção
1. **Feche e reabra o Excel** (ele só relê o ambiente/DLL num processo novo).
2. Teste:
   ```
   =TESTE()                              -> "OK — path: ...\calcrf_addin"
   =PU("FGEN13";"13/06/2025";6,4686%)    -> ~961,70
   ```
3. Se o `=TESTE()` voltar "OK" mas o `=PU` der `ERRO: APIs sem resposta (B3/FI)`, a DLL está
   resolvida — o que falta são as **variáveis de ambiente** `token_calc_b3`, `token_fianalytics`,
   `proxy_http`, `proxy_https` (ver `LEIA-ME.md` / `CLAUDE.md`).

## Resumo
O VBA carrega `xlwings64-0.36.6.dll` **da pasta do `pythonw.exe`**. Garanta que: (1) o xlwings é
**exatamente 0.36.6**, (2) o add-in usa o `pythonw` **certo** (não o da Store), e (3) a DLL
`xlwings64-0.36.6.dll` está na raiz desse Python. Reinicie o Excel e teste.
