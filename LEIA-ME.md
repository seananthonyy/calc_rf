# AntonioOliveiraCalc — Add-in de Renda Fixa para Excel

Todas as funções têm o prefixo **`cp`** (ex.: `=cpPu`). Puxam de APIs: **B3 Calculator**,
**FI Analytics** e **Banco Central (SGS)**. Contratos de DI futuro (`DI1F27`…) são calculados
localmente. Datas aceitam `"dd/mm/aaaa"`, célula de data ou `HOJE()`.

## Precificação
| Função | Retorno |
|---|---|
| `=cpPu(ticker; data; taxa%)` | PU de Operação (taxa como % do Excel, ex. 6,4618%) |
| `=cpTaxa(ticker; data; pu)` | Taxa de negociação (decimal → formate como %) |
| `=cpDur(ticker; data; taxa%)` | Duration de Macaulay em anos |
| `=cpCdi(dataIni; dataFim; percentual)` | Fator de CDI acumulado (percentual = número puro, 100 = 100%) |

## Dados do papel (taxa é OPCIONAL em PU Par/VNA/Fluxo — usa a de emissão se omitida)
| Função | Retorno |
|---|---|
| `=cpPupar(ticker; data; [taxa%])` | PU Par (valor nominal atualizado + juros) |
| `=cpVna(ticker; data; [taxa%])` | VNA — Valor Nominal Atualizado |
| `=cpFluxo(ticker; data; [taxa%])` | Fluxo restante (spill): Data · Tipo · Prazo(DU) · VF · VP |
| `=cpVencimento(ticker)` | Data de vencimento |
| `=cpEmissao(ticker)` | Data de emissão |
| `=cpInicio(ticker)` | Data de início de rentabilidade |
| `=cpTaxaEmissao(ticker)` | Taxa de emissão (unidade nativa: 113.5 = %DI, 6.4618 = IPCA+) |
| `=cpVne(ticker)` | Valor Nominal de Emissão |
| `=cpAniversario(ticker)` | Dia de aniversário |

## Métricas / gross up (FI Analytics)
| Função | Retorno |
|---|---|
| `=cpGrossUp(ticker; data; [taxa%])` | Taxa equivalente após imposto |
| `=cpGrossUpTipo(ticker; data; [taxa%])` | `GROSS_UP` (isento) ou `AFTER_TAX` (tributado) |
| `=cpConvexidade(ticker; data; [taxa%])` | Convexidade |
| `=cpDv01(ticker; data; [taxa%])` | DV01 (variação do PU por 1 bp) |
| `=cpDurMod(ticker; data; [taxa%])` | Duration modificada (anos) |
| `=cpDiPerc(ticker; data; [taxa%])` | % do DI equivalente (1.135 = 113,5%) |
| `=cpSpreadDi(ticker; data; [taxa%])` | Spread sobre o DI |

## Campos genéricos (qualquer campo das APIs)
| Função | Retorno |
|---|---|
| `=cpFi(ticker; data; taxa%; "campo")` | Qualquer campo da FI (ex.: `"modifiedDuration"`, `"spreadOverNTNB"`) |
| `=cpBond(ticker; "campo")` | Qualquer campo do getBondDetails (ex.: `"issuer"`, `"method"`, `"status"`) |

## Indicadores — Banco Central (valor mais recente por padrão)
| Função | Série |
|---|---|
| `=cpSelic()` | Meta SELIC (% a.a.) |
| `=cpSelicOver()` | SELIC over anualizada (% a.a.) |
| `=cpCdiAno()` | CDI anualizado base 252 (% a.a.) |
| `=cpCdiDia()` | CDI do dia (% ao dia) |
| `=cpIpca()` | IPCA do último mês (% mês) |
| `=cpIpcaAno()` | IPCA acumulado 12 meses (%) |
| `=cpIgpm()` / `=cpInpc()` | IGP-M / INPC do último mês (% mês) |
| `=cpTr()` | TR — Taxa Referencial (%) |
| `=cpDolar([data])` / `=cpEuro([data])` | Câmbio PTAX venda (R$) |
| `=cpBcb(serie; [dataIni]; [dataFim])` | Qualquer série do SGS por número; sem datas = último valor; com datas = spill (Data · Valor) |

## Dias úteis (feriados ANBIMA)
| Função | Retorno |
|---|---|
| `=cpEhDiaUtil(data)` | VERDADEIRO/FALSO |
| `=cpDiasUteis(dataIni; dataFim)` | Nº de dias úteis (início inclusive, fim exclusive) |
| `=cpDiaUtilPosterior(data)` | Próximo dia útil após a data |
| `=cpDiaUtilAnterior(data)` | Dia útil anterior à data |
| `=cpDiaUtilMaisN(data; n)` | Data ± n dias úteis (n negativo anda para trás) |

## Diagnóstico
| Função | Retorno |
|---|---|
| `=cpTeste()` | "OK — path: …" ou o erro de import |
| `=cpLimparCache()` | Esvazia o cache das APIs e força novas chamadas |

---
**Observações**
- Erros aparecem como texto `ERRO: …` na célula.
- Performance: as funções compartilham cache — o 1º cálculo de um papel bate na rede (~1–5 s),
  os demais vêm do cache (disco, 10 min). Rode `=cpLimparCache()` para forçar atualização.
- Instalação/atualização/config no banco: ver `INSTALAR_NO_BANCO.md` e `CONFIGURAR_PYTHON.md`.
  Atualizar a lógica = trocar os `.py` na share e reabrir o Excel. Trocar nomes/adicionar funções
  exige RE-BAKE do `.xlam` (Excel fechado) — ver `CLAUDE.md`.
