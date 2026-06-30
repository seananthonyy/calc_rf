# Add-in CalcRF — instalação e atualização (guia humano)

Add-in Excel **API-only** (B3 → FI Analytics). Empacotado como add-in customizado do xlwings:
um único `.xlam` que disponibiliza as UDFs `=PU`, `=DUR`, `=TAXA` em **qualquer planilha**, sem
"Import Functions" por workbook. (Detalhe técnico/arquitetura para alterações: ver `CLAUDE.md`.)

## Conteúdo da pasta

| Arquivo | O que é |
|---|---|
| `calcrf_addin.xlam` | o add-in (VBA do xlwings + ribbon + UDFs já registradas) |
| `calcrf_addin.py` | as UDFs (módulo do add-in) — só API, sem cálculo local, sem banco |
| `apis.py` | cliente B3/FI; lê segredos e proxy de variáveis de ambiente |
| `config.py` | mínimo (só `ENV_PATH`, fallback de dev) |
| `CLAUDE.md` | documentação técnica para o Claude Code |

> O add-in adiciona **a própria pasta ao PYTHONPATH automaticamente** — por isso os `.py` ficam
> junto do `.xlam`. Mantenha os arquivos sempre na mesma pasta.

## Pré-requisitos no PC

1. **Python** acessível (instalado na máquina ou portátil na share) com o pacote **xlwings**
   (`pip install xlwings`). Se o `pythonw.exe` estiver no PATH, o add-in já o encontra.
2. **Variáveis de ambiente** setadas no usuário (segredos e proxy):
   `token_calc_b3`, `token_fianalytics`, `proxy_http`, `proxy_https`.
3. Excel: **Central de Confiabilidade → Macros → "Confiar no acesso ao modelo de objeto do VBA"**.

## Instalação (uma vez por usuário)

1. Excel → **Arquivo → Opções → Suplementos** → "Gerenciar: Suplementos do Excel" → **Ir...**
2. **Procurar** → selecione `calcrf_addin.xlam` (na pasta da share) → OK.
3. Teste numa célula: `=PU("FGEN13";"13/06/2025";6,4686%)` → ~961,70.

> As UDFs já vêm registradas no `.xlam` (módulo `xlwings_udfs`). O passo de "Import Functions"
> **não** é necessário para o usuário final — só para quem altera as funções (ver `CLAUDE.md`).

## Atualizar (no futuro)

| Mudou | O que fazer | Usuário final |
|---|---|---|
| **Lógica** (`apis.py`, `calcrf_addin.py`, proxy, cálculo) | colar o `.py` novo na pasta | reabrir o Excel |
| **Nome/assinatura de UDF** | reimportar no `.xlam` (ver `CLAUDE.md`) + trocar o `.xlam` | reabrir o Excel |

99% das atualizações são **colar um `.py` na pasta**.

## Cuidados
- O `.xlam` fica **travado enquanto algum Excel o tiver aberto**. Para trocar o **próprio `.xlam`**,
  faça fora do horário ou versione o nome. Trocar os `.py` não tem essa trava.
- A pasta (share) precisa estar acessível quando o Excel abre.
