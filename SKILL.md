---
name: planilhas-design
description: Guia de estilo visual padrão (identidade Saldus) para planilhas (Excel .xlsx e Google Sheets) — paleta de cores, fontes, cabeçalhos, bordas, larguras, formatos de número e cores de gráficos. Use SEMPRE que for criar, formatar, estilizar ou "deixar bonita" qualquer planilha, tabela de Excel, Google Sheets, relatório em planilha, controle, orçamento, cronograma ou dashboard em planilha — mesmo que o usuário não peça um estilo específico — para que todas as planilhas saiam com a mesma identidade visual.
---

# Design de Planilhas — Estilo Padrão

Este guia define **um único estilo visual** para todas as planilhas. O objetivo é que qualquer planilha
(feita hoje, mês que vem, por você ou por outra IA) pareça parte da mesma família: limpa, legível e
profissional. Por isso, siga os valores exatos abaixo em vez de improvisar cores ou fontes — a
consistência é o que faz o estilo funcionar.

Se o usuário pedir explicitamente algo diferente (outra cor, outra fonte), atenda o pedido dele;
fora isso, use este padrão.

---

## 1. Paleta de cores

**Cores da marca:** `#12424D` (petróleo) · `#FFB745` (âmbar) · `#FFFFFF` (branco).
Todas as outras cores abaixo são tons derivados delas ou cores de significado (bom/ruim/atenção).

| Token | Hex | Uso |
|---|---|---|
| `primaria` | `#12424D` | **Marca.** Faixa do topo, fundo do cabeçalho da tabela, borda da linha de total |
| `destaque` | `#FFB745` | **Marca.** Linha de 3px sob o cabeçalho, destaque de célula-chave (ex.: o número mais importante), 2ª cor dos gráficos. Só como *fundo/linha* — nunca como cor de texto sobre branco (contraste baixo) |
| `branco` | `#FFFFFF` | **Marca.** Fundo geral da planilha e texto do cabeçalho |
| `primaria-clara` | `#E7EFF0` | Fundo da linha de total e de subcabeçalhos/agrupamentos |
| `zebra` | `#F5F8F8` | Fundo das linhas pares (listras alternadas) |
| `borda` | `#D5E0E2` | Bordas finas entre linhas |
| `texto` | `#1E2B2E` | Texto padrão do corpo |
| `texto-suave` | `#5F7378` | Notas, fonte dos dados, unidades, status neutro (Cancelado) |
| `texto-faixa` | `#B7CDD2` | Subtítulo sobre a faixa petróleo do topo |
| `positivo` | `#2E7D32` | Valores bons / variação positiva / status "OK" |
| `negativo` | `#C62828` | Valores negativos / variação negativa / status "Atrasado" |
| `alerta` | `#B26A00` | Status "Atenção" / pendente |
| `entrada` | `#FFF1D6` | Fundo de células que o usuário deve preencher (inputs) — tom claro do âmbar |
| `entrada-texto` | `#12424D` | Cor da fonte nas células de entrada (negrito leve não é necessário) |

**Por quê:** o petróleo dá identidade e peso; o âmbar aparece pouco, como "tempero" — se ele estiver em
todo lugar, deixa de chamar atenção. Branco e tons claros de petróleo fazem o resto. Verde/vermelho ficam
reservados para *significado* (bom/ruim) — usados como decoração, perdem a função de alerta.

### Gráficos — estilo
- Limpo: sem moldura, sem linhas de grade, sem fundo; título em Sora 14 petróleo.
- **Rótulos mostram só o valor** (ex.: "R$ 364.200"), sem casas decimais e sem nome da série/categoria
  ou quadradinho de legenda. Com rótulos, oculte o eixo de valores (é informação repetida).
- Barras horizontais para comparar categorias com nomes longos; colunas para meses/tempo.
- Legenda só com 2+ séries, embaixo.
- **Posição — a aba deve parecer uma tela só (um painel):** o gráfico fica abaixo da tabela, 2 linhas
  após a nota, com a **mesma largura da tabela** — começa na coluna A e termina na borda direita da
  última coluna, alinhado com a faixa do topo e com a tabela. Nada de gráfico pequeno encostado à
  esquerda deixando um vazio à direita.
  - Tabela muito larga (> ~1400 px): gráfico com largura máxima, centralizado sob a tabela.
  - Dois gráficos: lado a lado, dividindo a largura da tabela, com ~24 px entre eles.
  - Altura ~340 px (≈ 17 linhas); com muitas categorias, aumente para as barras não ficarem espremidas.
  - A **área interna** do gráfico (barras + rótulos) também usa ~100% da largura do quadro — fixe o
    layout da área de plotagem; o Excel por padrão deixa uma faixa vazia à direita.

### Cores para gráficos (nesta ordem)
1. `#12424D` (petróleo) · 2. `#FFB745` (âmbar) · 3. `#3E8A96` (petróleo médio) ·
4. `#8CC2C9` (petróleo claro) · 5. `#D9822B` (âmbar escuro) · 6. `#A7B4B7` (cinza)

Uma série só → use apenas a cor 1. Destaque de um item → item em cor 1, demais em cor 6.

---

## 2. Tipografia

Duas fontes, cada uma com um papel:

| Fonte | Papel | Onde |
|---|---|---|
| **Sora** | Títulos | Título da planilha (A1), cabeçalho da tabela, títulos de gráficos e de seções/blocos |
| **Poppins** | Texto simples | Subtítulo, corpo da tabela, linha de total, notas, rótulos e legendas de gráficos |

- **Google Sheets:** as duas são Google Fonts — menu de fontes → "Mais fontes" → busque Sora e Poppins.
- **Excel:** as fontes precisam estar instaladas no computador (grátis em fonts.google.com — Sora e
  Poppins). Sem elas, o Excel mostra a fonte padrão — avise o usuário disso ao entregar um .xlsx.
- **Fallback:** se não der para usar Sora/Poppins no ambiente, use `Arial` para tudo.
- Poppins e Sora são mais largas que Arial: dê ~15% a mais de largura às colunas.

| Elemento | Fonte | Tamanho | Estilo | Cor |
|---|---|---|---|---|
| Título da planilha | Sora | 16 pt | negrito | branco, sobre a faixa petróleo |
| Subtítulo (período, atualização) | Poppins | 10 pt | itálico | `texto-faixa`, sobre a faixa petróleo |
| Cabeçalho da tabela | Sora | 10 pt | negrito | branco sobre `primaria`, borda inferior média âmbar |
| Corpo | Poppins | 10 pt | normal | `texto` |
| Linha de total | Poppins | 10 pt | negrito | `primaria` |
| Notas/rodapé | Poppins | 9 pt | itálico | `texto-suave` |
| Título de gráfico | Sora | 14 pt | negrito | `primaria` |

Não use outras fontes além dessas duas, nem sublinhado decorativo ou texto todo em MAIÚSCULAS no corpo.

---

## 3. Layout da aba

**Princípio: pense na aba inteira como um único bloco.** Antes de entregar, olhe a aba de cima a baixo:
faixa do topo, cartões, tabela, nota e gráfico devem começar na mesma borda esquerda (coluna A),
terminar na mesma borda direita (última coluna da tabela) e ter espaçamentos iguais entre si. Blocos
lado a lado têm o mesmo tamanho. Nada sobra, vaza ou fica "solto" — se algo parecer estranho em
relação ao resto, ajuste larguras, mescle células ou quebre linhas até o conjunto ficar uniforme.

```
     A                B            C          D
1 │█ Título da planilha (Sora 16, branco)     [logo] █│ ← FAIXA PETRÓLEO nas linhas 1–2,
2 │█ Subtítulo · período · atualizado em …          █│   na largura exata da tabela (altura 34 + 24)
3 │                                                   ← linha em branco (respiro)
4 │ CABEÇALHO │ CABEÇALHO │ CABEÇALHO │ ...           ← fundo petróleo, texto branco, linha âmbar embaixo
5 │ dado      │ dado      │ dado                      ← branco
6 │ dado      │ dado      │ dado                      ← zebra #F5F8F8
… │ …
N │ Total     │           │ =SOMA()                   ← fundo #E7EFF0, negrito, borda superior petróleo
N+2│ Fonte: … / Notas                                 ← 9pt cinza itálico
```

Regras:
- **Faixa do topo:** as linhas 1–2 inteiras, da coluna A até a última coluna da tabela, têm fundo
  `primaria` (petróleo). Título em A1 (Sora 16, branco, recuo 1, alinhado embaixo), subtítulo em A2
  (Poppins 10 itálico, `texto-faixa`, recuo 1, alinhado em cima) — assim os dois ficam juntos no meio da
  faixa. A logo (versão clara) fica à direita, dentro da faixa. É a "capa" de toda aba.
- **Tabela começa na linha 4, coluna A.** Linha 3 fica vazia e branca, separando a faixa do cabeçalho.
- **Aba de resumo = resultado primeiro.** Quem abre a planilha precisa ver a resposta sem ler a tabela.
  Na aba principal, coloque **cartões de resultado** nas linhas 4–6 (rótulo em MAIÚSCULAS 9pt, valor em
  Sora 20 negrito, detalhe em 9pt), lado a lado na largura da tabela, e a tabela começa na linha 8.
  - 2 a 4 cartões, respondendo às perguntas do usuário (ex.: "quanto pagamos hoje", "por ano",
    "quanto economizamos", "quanto era antes").
  - **Blocos de largura igual**: os cartões dividem a largura da tabela em partes iguais (alargue um
    pouco as colunas se preciso). Vale para qualquer conjunto de blocos lado a lado (cartões, gráficos,
    painéis): mesmo tamanho, mesmo espaçamento, bordas alinhadas com a tabela.
  - Só o número principal tem fundo âmbar (`destaque`); os outros, `primaria-clara`.
  - Os valores são fórmulas ligadas à linha de total — nunca números digitados.
  - Evite colunas que confundem (ex.: "% do total" sem contexto): cada coluna precisa responder algo claro.
  - No script: `aplicar_estilo(..., linha_cabecalho=8, cartoes=[{rotulo, valor, formato, detalhe, destaque}, ...])`
    — divide em blocos iguais sozinho.
- **Congelar painéis** abaixo do cabeçalho só em **listas longas** (8+ linhas) — e a primeira coluna se
  ela for o identificador e houver muitas colunas. Em abas de resumo/painel não congele: elas cabem na
  tela e o congelamento corta a primeira linha da tabela.
- **Ocultar linhas de grade** (gridlines). As listras e bordas finas já guiam o olho.
- **Filtro** na linha de cabeçalho em **listas de dados** (8+ linhas). Em tabelas de resumo curtas, com
  subtotais ou só fórmulas, deixe sem filtro — filtrar ali não faz sentido.
- **Bordas:** só horizontais, finas, cor `borda` — exceto sob o cabeçalho (média, âmbar `#FFB745`). Sem bordas verticais, sem grades pretas. Linha de total
  tem borda superior *média* em `primaria`.
- **Altura de linha:** faixa 34 + 24; cabeçalho 24; corpo 20; total 22.
- **Alinhamento** (a regra mais importante para a tabela parecer "arrumada"):

  | Conteúdo | Horizontal | Por quê |
  |---|---|---|
  | Texto (nomes, descrições, categorias) | esquerda | lemos da esquerda para a direita |
  | Números, moeda, %, variação | direita | casas decimais e milhares ficam alinhados e fáceis de comparar |
  | Datas, meses, status, códigos curtos, Sim/Não, dia do mês, plano/tipo | centro | têm largura parecida; centro fica equilibrado |
  | Título (A1), subtítulo (A2), notas | esquerda | alinhados com a borda esquerda da tabela |
  | Rótulo "Total" | esquerda | igual aos demais textos da 1ª coluna |

  - **Cabeçalho segue o alinhamento da sua coluna**: cabeçalho de coluna de valor fica à direita,
    de data fica no centro. Cabeçalho centralizado sobre números desalinhados é o erro mais comum.
  - **Vertical: sempre ao centro**, em todas as células (cabeçalho, corpo e total).
  - **Recuo de 1** (Excel: "Recuo") nas células alinhadas à esquerda e à direita, para o conteúdo não
    encostar na borda da célula. (Google Sheets não tem recuo; compense com a largura extra.)
  - Com filtro ligado, o botão ▾ ocupa o canto direito do cabeçalho: em cabeçalho alinhado à direita
    use recuo 3 e deixe a coluna larga o bastante para o texto + o botão (nunca deixe o botão cortar o nome).
  - Prefira cabeçalhos curtos (1–3 palavras): "Desp. administrativas" em vez de "Despesas administrativas
    do período". Cabeçalho comprido deixa a coluna larga demais.
  - Mesma coluna = mesmo alinhamento do começo ao fim (inclusive na linha de total).
- **Larguras de coluna:** ajuste ao conteúdo, com mínimo 10 e máximo 50 (unidades do Excel ≈ caracteres).
  Coluna de descrição/observação: 40–50 com quebra de texto.
- **Dentro da tabela também:** célula com texto maior que a coluna (observações, descrições) quebra
  linha e a linha fica mais alta para o texto aparecer inteiro — nunca cortado, nunca invadindo a
  coluna ao lado. O `aplicar_estilo` faz isso sozinho.
- **Nada passa da borda direita da tabela.** Todo texto solto — nota de rodapé, observação, explicação,
  título, subtítulo, textos dos cartões — fica dentro da largura da tabela (coluna A até a última). Se
  não couber em uma linha, **quebra linha** (mescle a largura da tabela naquela linha, ligue "quebrar
  texto" e aumente a altura da linha). Texto vazando para as colunas vazias à direita quebra o
  "bloco" visual da aba. No script: `quebrar_na_largura(ws, linha, 1, ultima_col)` — o
  `aplicar_estilo` já faz isso com a nota, o título e o subtítulo.
- **Mesclar células é liberado e recomendado** para o layout ocupar os espaços certinho: faixa do topo,
  cartões de resultado, títulos de bloco, notas, textos explicativos, rótulos que precisam cobrir 2+
  colunas. Use sempre que deixar o bloco mais alinhado e preenchido.
  - Única exceção: **dentro das linhas de dados de uma tabela** com filtro/ordenação (cabeçalho e
    corpo), porque mesclagem ali quebra filtro, ordenação e fórmulas.

---

### Logo (Saldus)
- A logo é o ícone âmbar arredondado com o "S" + o nome **SALDUS** ao lado.
- **Arquivo-mestre em vetor:** `assets/logo-saldus.svg` (letras creme, para fundo escuro) e
  `assets/logo-saldus-escura.svg` (letras petróleo, para fundo branco). Use o SVG sempre que o programa
  aceitar (Figma, web, Google Slides).
- Para planilhas e Word (que não aceitam SVG pelos scripts): `assets/logo.png` (letras claras — **é a
  usada na faixa petróleo do topo**) e `assets/logo-escura.png` (letras em petróleo — só para fundo
  branco), gerados do SVG em alta resolução (1289 px de largura).
  Se o arquivo não estiver disponível, deixe o espaço livre e não recrie a logo.
- Posição: dentro da faixa do topo, canto **direito**, centralizada na altura das linhas 1–2, terminando
  ~12 px antes da borda direita da tabela; altura ≈ 40 px, sem distorcer a proporção.
- Logo **apenas na aba principal** (a primeira). As demais abas têm a faixa com título, mas sem logo.

---

## 4. Formatos de número (padrão Brasil)

| Tipo | Formato | Exemplo |
|---|---|---|
| Moeda | `"R$" #,##0.00;-"R$" #,##0.00` | R$ 1.234,56 |
| Inteiro | `#,##0` | 12.345 |
| Decimal | `#,##0.00` | 1.234,57 |
| Percentual | `0.0%` | 12,5% |
| Variação % | `+0.0%;-0.0%;0.0%` | +4,2% / -3,1% |
| Data | `dd/mm/yyyy` | 16/11/2026 |
| Mês | `mmm/yyyy` | nov/2026 |

(O separador `,`/`.` no código do formato é trocado automaticamente pela localidade pt-BR do programa.)

Números negativos sempre em vermelho `negativo` (`#C62828`) — aplique por **formatação condicional** ("menor que 0" → cor da fonte), não com `[Red]` no formato, porque o `[Red]` é o vermelho puro do programa e foge da paleta. Não misture formatos na mesma coluna.

---

## 5. Estados e destaques

- **Células de entrada** (que o usuário digita: premissas, metas, orçamento, parâmetros) → fundo
  `entrada`, fonte `entrada-texto`. Assim fica óbvio o que pode ser editado e o que é fórmula.
  Pode ser uma coluna inteira ou uma célula só; a cor de entrada vence a zebra. Use só para
  parâmetros — numa lista de registros comum (clientes, itens) não pinte tudo de âmbar.
- **Status** (coluna de situação) → formatação condicional só na cor da fonte, em negrito:
  - `positivo`: OK, Concluído, Pago, Ativo, Em dia, Aprovado, Novo
  - `alerta`: Pendente, Atenção, Em andamento, Em análise, Em renovação
  - `negativo`: Atrasado, Erro, Vencido, Reprovado
  - `texto-suave` (neutro): Cancelado, Inativo, Encerrado, Arquivado — não são "ruins", só saíram de cena
  - Termo novo → encaixe no grupo de significado mais próximo (no script: `status_extra`).
- **Itens cancelados/encerrados** continuam na lista (histórico), no fim da tabela, com a linha inteira
  em cinza `texto-suave` e riscada (formatação condicional só com cor + tachado, sem trocar a fonte).
  Ficam fora dos totais ("Total ativos") e, quando fizer sentido, o resumo mostra a **economia** que
  geraram (mensal e anual). Itens **novos** (contratados agora) ficam com status "Novo" e contam como
  ativos. O "antes" de uma comparação = ativos − novos + cancelados, e a economia mostrada é a
  **líquida** (cancelados − novos), não só o que foi cortado.
- **Linha de total:** some só colunas que fazem sentido somar. Saldo acumulado, %, médias e taxas não se
  somam: deixe vazio ou use a fórmula certa (último saldo, total/total etc.).
- **Destaque de um número-chave** (ex.: saldo final, total geral num resumo): fundo `destaque` (âmbar),
  texto `primaria` em negrito. No máximo um ou dois por aba.
- **Formatação condicional numérica** com moderação: no máximo uma escala/barras por tabela,
  usando tons de petróleo (de `#FFFFFF` a `#3E8A96`).

---

## 6. Abas

- Nomes curtos, sem acentos problemáticos nem espaços duplos: `Resumo`, `Dados`, `Cálculos`, `Listas`.
- Aba principal/resumo primeiro. **Não pinte as abas** (cor da guia): o Excel desenha essa cor com um
  degradê que foge do estilo chapado. A identidade já está na faixa do topo.
- **Nada de degradês** em lugar nenhum (células, gráficos, abas): só cores sólidas da paleta.
- Uma tabela principal por aba. Um bloco pequeno de premissas/parâmetros (células de entrada) acima ou
  ao lado é aceitável, com o mesmo estilo de cabeçalho.
- Planilha com várias abas: `Resumo` (indicadores, totais, gráfico — com logo) → abas de dados
  (`Custos`, `Clientes`…) → `Listas` (opções dos menus suspensos) por último.
- Use **validação de dados** (lista suspensa) em colunas de categoria, status e tipo, puxando de `Listas`.

---

## 7. Como aplicar

### Excel (.xlsx) com Python
Use o script pronto `scripts/estilo_excel.py` (biblioteca `openpyxl`). Ele aplica todo o padrão acima:

```python
import sys; sys.path.insert(0, "<pasta da skill>/scripts")
from estilo_excel import aplicar_estilo, estilizar_grafico, posicionar_grafico

# 1) escreva os cabeçalhos na linha 4 e os dados a partir da linha 5; depois:
aplicar_estilo(ws, titulo="Controle de Despesas", subtitulo="Jan–Dez 2026 · atualizado em 16/11/2026",
               formatos={"Valor": "moeda", "Data": "data", "Plano": "centro"},
               linha_total=True, colunas_total=["Valor"], colunas_entrada=["Meta"],
               coluna_status="Status", nota="Fonte: …")
# abas secundárias: logo=None

# 2) gráfico (openpyxl.chart): estilo + posição com a mesma largura da tabela
estilizar_grafico(grafico, titulo="Receita por plano", formato_valor='"R$" #,##0', rotulos=True)
posicionar_grafico(ws, grafico, linha=linha_da_nota + 2, coluna_final=ultima_coluna_da_tabela)
# dois gráficos lado a lado: parte=(1, 2) e parte=(2, 2)

# 3) salve com salvar() em vez de wb.save() — corrige o formato dos rótulos do gráfico
from estilo_excel import salvar
salvar(wb, "arquivo.xlsx")
```

Tipos de `formatos`: `moeda`, `inteiro`, `decimal`, `percentual`, `variacao`, `data`, `mes`,
`status`, `centro` (texto curto centralizado), `texto` (padrão). O filtro liga sozinho em listas com
8+ linhas (`filtro=True/False` força). As fontes Sora e Poppins precisam estar instaladas no computador
de quem abrir o .xlsx — mencione isso ao entregar.

Se o script não estiver disponível no ambiente, reproduza manualmente as regras das seções 1–6.

### Google Sheets
Use `scripts/estilo_sheets.gs` (Google Apps Script): na planilha, **Extensões → Apps Script**, cole o
arquivo, salve e rode `aplicarEstiloPadrao`. Ele formata a aba ativa assumindo o mesmo layout
(título A1, subtítulo A2, cabeçalho na linha 4).

### Sem código (manual ou outra IA)
Siga as seções 1–6 como checklist. Os valores exatos estão em `tokens.json`.

---

## 8. Checklist final (confira antes de entregar)

- [ ] Faixa petróleo nas linhas 1–2 (largura da tabela): título branco em A1, subtítulo em A2, logo clara à direita (só na 1ª aba); linha 3 vazia
- [ ] Cabeçalho na linha 4: fundo `#12424D`, texto branco, negrito, linha âmbar `#FFB745` embaixo
- [ ] Sora no título e no cabeçalho; Poppins no resto (10pt no corpo)
- [ ] Alinhamento: texto à esquerda, números à direita, datas/status ao centro; cabeçalho igual à coluna; vertical ao centro
- [ ] Listras `#F5F8F8` e bordas só horizontais `#D5E0E2`
- [ ] Linhas de grade ocultas; painéis congelados e filtro só nas listas de dados longas
- [ ] Nenhum texto passa da borda direita da tabela (notas e textos longos com quebra de linha)
- [ ] Formatos de número corretos por coluna; negativos em vermelho
- [ ] Linha de total (se houver) com fundo `#E7EFF0`, negrito e borda superior petróleo
- [ ] Células de entrada em âmbar-claro `#FFF1D6` com fonte petróleo
- [ ] Gráficos usando a paleta na ordem definida
