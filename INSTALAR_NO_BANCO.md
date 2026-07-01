# Instalar e configurar o add-in CalcRF no banco

Guia completo, passo a passo, para deixar o add-in **CalcRF** funcionando num PC do banco.
Escrito para ser lido por uma pessoa **ou** pelo Claude Code do banco.

> **Peça ao Claude do banco:** "leia o `INSTALAR_NO_BANCO.md` e me guie". Ele tem tudo aqui.

---

## Mapa dos documentos (o que ler para cada coisa)

| Documento | Para quê |
|---|---|
| **INSTALAR_NO_BANCO.md** (este) | instalar do zero e configurar por PC |
| `CONFIGURAR_PYTHON.md` | fixar qual Python o add-in usa (PC com vários Python/xlwings) |
| `DIAGNOSTICO_DLL.md` | erro `xlwings64-0.36.6.dll` (DLL não encontrada) |
| `LEIA-ME.md` | visão geral rápida (humano) |
| `CLAUDE.md` | arquitetura técnica (para alterar o código) |

---

## O que o add-in faz

Expõe 5 funções de planilha (UDFs) para precificar renda fixa:

| UDF | Retorno | Fonte |
|---|---|---|
| `=PU(ticker; "dd/mm/aaaa"; taxa%)` | PU de Operação | B3 → FI (DI é local) |
| `=DUR(ticker; "dd/mm/aaaa"; taxa%)` | Duration de Macaulay (anos) | B3 → FI (DI é local) |
| `=TAXA(ticker; "dd/mm/aaaa"; pu)` | Taxa de negociação (formatar célula como %) | B3 → FI (DI é local) |
| `=TESTE()` | Diagnóstico ("OK — path: ...") | local |
| `=LIMPARCACHE()` | Esvazia o cache das APIs | local |

- **Debênture, CRI/CRA, NTN-B/NTN-F**: vêm das APIs (B3 Calculator → FI Analytics).
- **DI (tickers `DI1...`)**: não há API → calculado **localmente** em `di.py`.
- Depois de instalado, a aba do ribbon aparece como **CalcRF**.

---

## Pré-requisitos do PC

1. **Excel 64-bit** (confirmar: Arquivo → Conta → Sobre o Excel).
2. **Python 64-bit** com **xlwings exatamente 0.36.6** e a DLL `xlwings64-0.36.6.dll` na raiz do
   Python (ao lado do `pythonw.exe`). Instalação padrão do python.org já põe a DLL na raiz.
   Verificar:
   ```
   python -c "import xlwings,os,sys,glob; r=os.path.dirname(sys.executable); print(xlwings.__version__); print(glob.glob(os.path.join(r,'xlwings*.dll')))"
   ```
   Esperado: `0.36.6` e a lista mostrando `xlwings64-0.36.6.dll`. Se não, ver `DIAGNOSTICO_DLL.md`.
3. **A pasta do add-in na share**, no caminho **`Z:\AntonioOliveira\CalcRF`** (é o `PYTHONPATH`
   embutido no `.xlam` — tem que ser esse caminho). Conteúdo mínimo:
   `calcrf_addin.xlam`, `calcrf_addin.py`, `apis.py`, `di.py`, `config.py`, `feriados_anbima.csv`.
   Atualizar puxando do repositório público do projeto.

---

## Passo a passo

### 1. Colocar a pasta na share
Garanta que `Z:\AntonioOliveira\CalcRF` existe e tem os arquivos acima, atualizados do repositório.
A share precisa estar acessível quando o Excel abre.

### 2. Confirmar o Python (pré-requisito 2)
Rode o comando de verificação acima. Se `pythonw` do PATH já for esse Python com 0.36.6, ótimo —
o add-in o encontra sozinho. Se o PC tem **vários Python**, veja o passo 6.

### 3. Setar as 4 variáveis de ambiente do usuário (segredos + proxy)
No banco não existe arquivo `.env`; os segredos entram por variável de ambiente. Setar **no usuário**:

| Variável | Conteúdo |
|---|---|
| `token_calc_b3` | token do B3 Calculator |
| `token_fianalytics` | API key do FI Analytics |
| `proxy_http` | `http://USUARIO:SENHA@HOST:PORTA` |
| `proxy_https` | `http://USUARIO:SENHA@HOST:PORTA` |

(Painel de Controle → "Editar as variáveis de ambiente do usuário", ou via GPO/login script.)
Sem elas, o `=TESTE()` dá "OK" mas o `=PU` retorna `ERRO: APIs sem resposta (B3/FI)`.

### 4. Liberar o acesso ao VBA (uma vez)
Excel → Arquivo → Opções → **Central de Confiabilidade** → Configurações → **Configurações de
Macro** → marcar **"Confiar no acesso ao modelo de objeto de projeto do VBA"**.

### 5. Habilitar o add-in
Excel → Arquivo → Opções → **Suplementos** → "Gerenciar: Suplementos do Excel" → **Ir...** →
**Procurar** → selecionar `Z:\AntonioOliveira\CalcRF\calcrf_addin.xlam` → OK.
> As UDFs já vêm **registradas** no `.xlam`. **Não** precisa "Import Functions" no banco.
> Deve aparecer a aba **CalcRF** no ribbon.

### 6. (Só se o PC tiver vários Python) Fixar o Python certo
Se der erro de DLL ou as fórmulas não calcularem num PC com múltiplos Python, siga o
**`CONFIGURAR_PYTHON.md`**: criar `%USERPROFILE%\.myaddin\myaddin.conf` com a linha
`"Interpreter","<caminho do pythonw.exe com xlwings 0.36.6>"`.
> ⚠️ Use a chave **`Interpreter`** (não `Interpreter_Win` — essa é ignorada; explicação no
> `CONFIGURAR_PYTHON.md`).

### 7. Testar
Feche e reabra o Excel. Numa célula:
```
=TESTE()                              -> OK — path: Z:\AntonioOliveira\CalcRF
=PU("FGEN13"; "13/06/2025"; 6,4686%)  -> ~961,70
=PU("DI1F27"; "01/07/2026"; 10%)      -> ~95310,20   (DI, cálculo local)
```

---

## Atualizar no futuro

| Mudou | O que fazer | Usuário final |
|---|---|---|
| **Lógica** (`apis.py`, `calcrf_addin.py`, `di.py`) | colar o `.py` novo na share | reabrir o Excel |
| **Nome/assinatura de UDF** ou **o `.xlam`** | trocar o `.xlam` (ver `CLAUDE.md`) | reabrir o Excel |

99% das atualizações são só **colar um `.py` na share** e reabrir o Excel.
O `.xlam` fica **travado enquanto algum Excel o tiver aberto** — troque fora de uso.

---

## Solução de problemas (rápida)

| Sintoma na célula | Causa provável | Onde resolver |
|---|---|---|
| `file not found: xlwings64-0.36.6.dll` | Python errado / xlwings ≠ 0.36.6 / DLL ausente | `DIAGNOSTICO_DLL.md` |
| fórmulas não calculam, PC com vários Python | add-in pegou o Python errado | `CONFIGURAR_PYTHON.md` (passo 6) |
| `ERRO: APIs sem resposta (B3/FI)` | faltam env vars / proxy | passo 3 |
| `Input past end of file` | `myaddin.conf` com linha em branco sobrando | `CONFIGURAR_PYTHON.md` |
| aba CalcRF não aparece / `#NAME?` | add-in não habilitado ou VBA não liberado | passos 4 e 5 |
| Excel congela em planilha do SharePoint | (já mitigado no `.xlam`) | ver `CLAUDE.md` |

---

## Checklist final (por PC)

- [ ] Pasta em `Z:\AntonioOliveira\CalcRF` atualizada
- [ ] Python 64-bit com xlwings **0.36.6** + DLL na raiz
- [ ] `token_calc_b3`, `token_fianalytics`, `proxy_http`, `proxy_https` setados
- [ ] "Confiar no acesso ao modelo de objeto de projeto do VBA" ligado
- [ ] `.xlam` habilitado (aba **CalcRF** aparece)
- [ ] (se multi-Python) `myaddin.conf` com a chave `Interpreter`
- [ ] `=PU("FGEN13";"13/06/2025";6,4686%)` ≈ 961,70
