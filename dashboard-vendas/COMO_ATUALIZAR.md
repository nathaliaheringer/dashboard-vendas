# Como manter o Dashboard de Vendas atualizado

Este guia explica **(1)** como atualizar os números do dashboard, **(2)** como deixá-lo
no ar na internet e **(3)** como instruir o Claude Cowork a fazer isso sozinho todo dia.

Você recebeu 3 arquivos:

| Arquivo | O que é |
|---|---|
| `dashboard_vendas_NH.html` | O dashboard. Abre em qualquer navegador. |
| `dados.json` | Os números atuais do dashboard. É **este arquivo que se troca** para atualizar. |
| `COMO_ATUALIZAR.md` | Este guia. |

O dashboard e os dados são separados de propósito: o `.html` é o "molde" e nunca muda;
o `dados.json` é o que muda todo dia. Assim você atualiza sem nunca mexer no dashboard.

---

## 1. Como atualizar os números (2 formas)

### Forma A — Botão "Atualizar dados" (uso imediato, sem hospedagem)

1. Abra o `dashboard_vendas_NH.html` no navegador.
2. Clique em **"Atualizar dados"** (canto superior direito).
3. Escolha o arquivo `dados.json` novo.
4. Pronto — o dashboard recarrega com os números novos e **guarda essa versão naquele navegador**
   (na próxima vez que abrir, já aparece atualizado).

> Funciona offline. A limitação: a atualização vale só para aquele navegador/computador.
> Se você quiser que outras pessoas vejam sempre a versão mais recente, use a Forma B.

### Forma B — Dashboard hospedado (recomendado para uso contínuo)

Quando o dashboard está hospedado num link na internet, ele **busca o `dados.json`
automaticamente** toda vez que é aberto. Para atualizar, basta substituir o `dados.json`
no servidor — todo mundo que abrir o link vê o número novo, sem clicar em nada.

É essa a forma ideal para o fluxo com o Cowork.

---

## 2. Como deixar o dashboard no ar (hospedagem)

O dashboard é um único arquivo HTML, então a hospedagem é gratuita e simples.

### Opção mais fácil — Netlify Drop

1. Crie uma pasta no seu computador, por exemplo `dashboard-nh`.
2. Coloque dentro dela os **dois** arquivos, **renomeando o HTML para `index.html`**:
   - `index.html`  (que é o `dashboard_vendas_NH.html` renomeado)
   - `dados.json`
3. Acesse **https://app.netlify.com/drop**
4. Arraste a pasta inteira para a página. Em segundos você recebe um link público
   (ex.: `https://dashboard-nh.netlify.app`).
5. Para atualizar depois: entre no site do Netlify, vá no projeto e arraste a pasta de novo
   com o `dados.json` atualizado. **Ou** deixe o Cowork fazer isso (próxima seção).

### Alternativa — GitHub Pages

1. Crie um repositório no GitHub e suba `index.html` + `dados.json`.
2. Em *Settings → Pages*, ative o Pages na branch `main`.
3. O link fica `https://seu-usuario.github.io/nome-do-repo`.
4. Para atualizar: suba um novo `dados.json` (substituindo o antigo) no repositório.

> Em ambas as opções: o que importa é que `index.html` e `dados.json` fiquem
> **na mesma pasta**. O dashboard procura o `dados.json` "ao lado dele".

---

## 3. Fluxo com o Claude Cowork

A ideia: todo dia o Cowork **extrai os relatórios das plataformas**, **gera um `dados.json`
novo** e **substitui** o arquivo (no Netlify/GitHub, ou te entrega para você carregar pelo botão).

### O que o Cowork precisa extrair

| Plataforma | Relatório | Observação |
|---|---|---|
| Hubla | Relatório de Vendas | exportar em `.xlsx`, período do mês corrente |
| Hubla | Relatório de Carrinho Abandonado | exportar em `.csv` |
| **Hotmart** | **Histórico de vendas** | exportar em `.csv` (delimitado por `;`); período do mês corrente — **detecção automática**: qualquer `.csv` com a coluna `Nome do Produtor` é reconhecido como Hotmart |
| Meta Ads | Relatório de campanhas, com detalhamento **por dia** | exportar em `.csv`; colunas: campanha, dia, valor gasto, impressões, alcance, frequência, CPM |
| Planilha de controle (Google Sheets) | Aba **DADOS [produto]** (ex.: DADOS RE, DADOS PSI) | exportar em `.csv`; colunas: Day, Campaign Name, CPM, Link Clicks, CPC, CTR, Checkouts Initiated, Purchases, Amount Spent, Impressions, Landing Page Views — **usadas para o funil real por produto** |

### Prompt pronto para colar no Cowork

> Copie o texto abaixo e cole no Claude Cowork. Anexe o `dados.json` atual junto —
> ele serve de modelo do formato exato a ser seguido.

```
Preciso atualizar o dados.json do meu Dashboard de Vendas.

Passo 1 — Extraia estes relatórios:
- Hubla → Relatório de Vendas do mês corrente (.xlsx)
- Hubla → Relatório de Carrinho Abandonado (.csv)
- Meta Ads (conta "CA - Nathalia Heringer") → relatório de campanhas com
  detalhamento POR DIA, com as colunas: nome da campanha, dia, valor gasto,
  impressões, alcance, frequência e CPM.

Passo 2 — Gere um novo dados.json seguindo EXATAMENTE a mesma estrutura
do dados.json que estou anexando (mesmas chaves, mesmo formato). Use as
regras de cálculo descritas no arquivo COMO_ATUALIZAR.md, seção 4.

Passo 3 — Me devolva o dados.json pronto para download.
```

Se o dashboard estiver hospedado e você tiver dado ao Cowork acesso à sua conta
do Netlify/GitHub, pode acrescentar ao prompt: *"Depois, substitua o dados.json
no meu projeto do Netlify/GitHub."*

---

## 4. Regras de cálculo (referência para o Cowork)

Estas são as regras usadas para transformar os relatórios crus no `dados.json`.

**Vendas e faturamento (Relatório de Vendas Hubla)**
- Considerar apenas faturas com status **pago**.
- `Faturamento` = soma da coluna **Valor do produto**.
- `Valor líquido (NH)` = soma da coluna **Valor Líquido**.
- `Itens vendidos` = soma de itens (produto principal + order bumps).
- `Ticket médio` = faturamento ÷ nº de vendas.
- Vendas de **"Formação Primeira Infância • Turma 09"** são **parcelas de recorrência**
  (parcelamento inteligente) — marcar com `"recorrente": true`, não são vendas novas.

**Origem / canal de venda (coluna UTM Origem)**
- `facebookads` → **Facebook Ads**  |  `whatsapp` → **WhatsApp**
- `instagram` / `biografia` / mídia `bio` → **Instagram (orgânico/bio)**
- UTM vazia → **Sem origem registrada**
- Público **Frio** e **Quente** vêm sempre do Facebook Ads (definido pela UTM de campanha
  conter `[Frio]` ou `[Quente]`).

**Produto da campanha (coluna UTM Campanha)**
- Campanha com `[RE]` → produto **Regulando as Emoções**
- Campanha com `[PSI08]` → produto **Curso PSI 0-8**

**Investimento (Meta Ads)**
- `Investimento em vendas` = soma do gasto das campanhas de conversão (objetivo de venda).
- A campanha `[INSTA][Seguidores]` é de audiência — entra separada, **fora do ROAS**.
- `ROAS de vendas` = faturamento ÷ investimento em vendas.
- `Lucro do funil` = valor líquido − investimento em vendas.
- `ROI` = lucro ÷ investimento × 100.  `CPA` = investimento ÷ nº de vendas.

**Funil de conversão por produto (planilha de controle)**
- Para cada produto que tem planilha (aba `DADOS [produto]`), agregar do CSV:
  - `funnel_clicks` = soma de **Link Clicks** das campanhas do produto
  - `funnel_lpv` = soma de **Landing Page Views**
  - `funnel_ctr` = cliques ÷ impressões × 100 (recalculado)
  - `funnel_cpc` = gasto ÷ cliques (recalculado)
  - `funnel_source` = `"planilha"`
- **IGNORAR** as colunas `Purchases` e `Checkouts Initiated` da planilha — são contagens do Meta (pixel) e divergem das vendas/carrinhos reais.
- `checkouts` (checkouts iniciados no funil) = **sempre** vem do Hubla = compras pagas + carrinhos abandonados do produto no período.
- `compras` (compras pagas no funil) = **sempre** vem do Relatório Hubla (mesma fonte do total de vendas) e, quando houver, do relatório da Hotmart somado.
- Se o produto **não tiver planilha de controle**, deixar `funnel_source: "estimativa"` — o dashboard mostra campos editáveis de cliques/LPV.

**Semanas** — sempre fixas dentro do mês: S1 = dias 1–7, S2 = 8–14, S3 = 15–21,
S4 = 22–28, S5 = 29–31. Semanas sem dados ficam com `"has_data": false`.

**Métodos de pagamento (Relatório de Vendas Hubla)**
- Coluna: **"Método de pagamento"** (índice 5 no XLSX, base zero).
- Valores possíveis: `Cartão de Crédito`, `PIX`, `Boleto`.
- Contar por método globalmente (`payment_dist`) e também por produto (`products_paid.RE.pay_dist`, `products_paid.PSI.pay_dist`).
- A distribuição por produto filtra apenas as faturas cujo `Nome do produto` (col 14) pertença ao produto.

**Parcelamento (Relatório de Vendas Hubla)**
- Coluna: **"Parcelas"** (índice 32 no XLSX, base zero).
- Registrar distribuição **apenas para faturas de Cartão de Crédito** (agrupado por número de parcelas 1x–12x).
- Calcular globalmente (`parc_dist`) e por produto (`products_paid.RE.parc_dist`, `products_paid.PSI.parc_dist`).

**Order Bumps por produto (Relatório de Vendas Hubla)**
- Coluna do produto OB: **"Nome do produto de orderbump"** (índice 16 no XLSX).
- Coluna do produto principal: **"Nome do produto"** (índice 14 no XLSX).
- Uma fatura pertence ao funil RE ou PSI se o produto aparecer em **qualquer uma das duas colunas** ("Nome do produto" OU "Nome do produto de orderbump"). Isso captura tanto compras diretas quanto compras do produto feitas como order bump dentro de outro funil.
- **Faturamento** (`fat`) de RE/PSI: contar apenas linhas em que o produto é o `Nome do produto` (coluna principal) — evita dupla contagem.
- **Contagem de vendas** (`faturas`) e OBs: usar OR (produto em qualquer coluna).
- Calcular `ob_detail` global e `products_paid.RE.ob_detail` / `products_paid.PSI.ob_detail` com nome, contagem e valor estimado de cada OB.
- Calcular `products_paid.RE.ob_val` e `products_paid.PSI.ob_val` (soma do valor estimado dos OBs por produto).

**Região (Relatório de Vendas, coluna Endereço Estado)**
- Agrupar os estados em macrorregiões (Sudeste, Sul, Nordeste, Centro-Oeste, Norte).
- Vendas sem estado preenchido → "Não informado".

---

## 5. Estrutura do `dados.json` (resumo)

O arquivo é um objeto JSON com estas chaves principais. **A forma mais segura de acertar
o formato é abrir o `dados.json` atual e seguir o mesmo padrão**, apenas trocando os números.

| Chave | Conteúdo |
|---|---|
| `period` | datas de início/fim do período e rótulo |
| `totals` | totais consolidados (faturamento, investimento, lucro, ROI, ROAS, etc.) |
| `daily` | lista com um registro por dia (data, gasto, faturamento, vendas, etc.) |
| `weeks` | os 5 blocos semanais |
| `campaigns` | uma entrada por campanha, com detalhamento diário (`daily`) |
| `products_paid` | dados de funil por produto (RE e PSI) + benchmarks + `pay_dist`, `parc_dist`, `ob_val`, `ob_detail` por produto |
| `origins` / `fb_split` | receita por canal de venda |
| `products` | receita por produto |
| `regioes` / `estados` | receita por região |
| `followers` | dados da campanha de seguidores |
| `payment_dist` | distribuição global de métodos de pagamento |
| `parc_dist` | distribuição global de parcelamento (CC) |
| `ob_detail` | detalhe global dos order bumps (nome, contagem, valor) |
| `insights` | textos de recomendação |

> **Importante:** mantenha **todas** as chaves do arquivo modelo. Se o dashboard receber
> um `dados.json` incompleto, ele avisa na tela e mantém os dados anteriores — nada quebra.

---

## Resumo rápido

- **Atualizar agora, sem hospedar:** botão "Atualizar dados" → escolher o `dados.json`.
- **Deixar no ar:** Netlify Drop com `index.html` + `dados.json` na mesma pasta.
- **Automatizar:** Cowork extrai os relatórios → gera o `dados.json` → substitui no servidor.
- O dashboard **nunca precisa ser reconstruído** — só o `dados.json` muda.
