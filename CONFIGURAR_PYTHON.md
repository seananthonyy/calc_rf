# Fixar o Python do add-in CalcRF (PCs com vários Python/xlwings)

Guia para quando o add-in **CalcRF** precisa usar um Python específico da máquina —
tipicamente em PCs que têm **mais de um Python instalado** e o add-in acaba pegando
o errado.

> Resumo em uma linha: crie o arquivo `%USERPROFILE%\.myaddin\myaddin.conf` com a
> linha `"Interpreter","<caminho do pythonw.exe>"`. Só isso.

---

## Quando usar este guia

Use se, ao chamar `=PU(...)`/`=DUR(...)`/`=TAXA(...)`, aparecer:

- `Could not load xlwings64-0.36.6.dll` (o Python escolhido não tem o **xlwings 0.36.6**), ou
- as fórmulas não calculam / erro de import, **e** a máquina tem vários Python.

Se o add-in já funciona, **não precisa** deste guia.

## Por que acontece

O add-in é **travado no xlwings 0.36.6** (a versão está embutida no VBA, e a partir
dela ele carrega o DLL `xlwings64-0.36.6.dll` de dentro do Python usado). Se nenhum
Python for fixado, o add-in usa o **`pythonw` do PATH** — que num PC com vários Python
pode ser um que **não tem** o xlwings 0.36.6. Fixando o Python certo, o problema some
(e o DLL certo vem junto, porque o DLL mora dentro do Python escolhido).

---

## Passo a passo

### 1. Descobrir o Python certo (o que tem xlwings 0.36.6)

Abra o **Prompt de Comando** (cmd) e liste os Python instalados:

```
py -0p
```

Teste os candidatos até achar o que responde **0.36.6** (troque o caminho):

```
"C:\Caminho\Do\Python\pythonw.exe" -c "import xlwings; print(xlwings.__version__)"
```

- Se imprimir `0.36.6` → é esse. Anote o caminho do **`pythonw.exe`** (com `w`).
- Se der erro ou outra versão → nesse Python instale a versão certa:
  `"C:\Caminho\Do\Python\python.exe" -m pip install xlwings==0.36.6`

### 2. Criar o arquivo de config

O arquivo fica em `C:\Users\<seu-usuário>\.myaddin\myaddin.conf` e tem **uma linha**.

**Jeito seguro (PowerShell)** — cria a pasta e o arquivo no formato/encoding corretos,
sem a linha em branco final que já quebrou UDFs aqui:

```powershell
$py = "C:\Caminho\Do\Python\pythonw.exe"   # <-- ajuste
$dir = "$env:USERPROFILE\.myaddin"
New-Item -ItemType Directory -Force $dir | Out-Null
[System.IO.File]::WriteAllText("$dir\myaddin.conf", "`"Interpreter`",`"$py`"`r`n", [System.Text.Encoding]::ASCII)
Get-Content "$dir\myaddin.conf"   # confere: deve mostrar UMA linha
```

O conteúdo final tem que ser **exatamente** (aspas incluídas, uma linha só):

```
"Interpreter","C:\Caminho\Do\Python\pythonw.exe"
```

### 3. Testar

Feche e reabra o Excel. Numa célula:

```
=PU("FGEN13"; "13/06/2025"; 6,4686%)
```

Deve retornar **~961,70**. Se vier número, funcionou.

---

## ⚠️ Detalhes que importam (não pule)

- **Use a chave `Interpreter`, NÃO `Interpreter_Win`.** O `.xlam` já traz `Interpreter_Win`
  em branco na planilha de config embutida, e essa planilha tem **precedência** sobre o
  arquivo do usuário — ou seja, `Interpreter_Win` no arquivo seria **ignorado**. A chave
  legada `Interpreter` não está na planilha embutida, então é a que "passa" para o arquivo.
  (No VBA do xlwings: `INTERPRETER_WIN` vazio → cai em `GetConfig("INTERPRETER","pythonw")`,
  que lê o arquivo do usuário; casamento de chave é case-insensitive.)
- **Aponte para o `pythonw.exe`** (com `w`), não o `python.exe` — evita abrir janela de console.
- **Uma linha só, sem linha em branco no final.** Um `\n` extra faz o `Input #` do VBA
  estourar (erro 62 "Input past end of file") e **quebra todas as UDFs**. Use o snippet
  PowerShell acima que já cuida disso. Encoding **ANSI/ASCII, sem BOM**.
- Esse arquivo é **por PC e por usuário** — não vai para o repositório, é local de cada máquina.

## Se não resolver

- Confirme o xlwings do Python fixado: `"<pythonw>" -c "import xlwings; print(xlwings.__version__)"` → `0.36.6`.
- Confirme o arquivo: `type %USERPROFILE%\.myaddin\myaddin.conf` → uma linha `"Interpreter","..."`.
- Central de Confiabilidade → Macros → **"Confiar no acesso ao modelo de objeto do VBA"** ligado.
- Se aparecer `Input past end of file` → o arquivo tem linha em branco sobrando; recrie com o snippet.
