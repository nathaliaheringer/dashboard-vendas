# Como manter o Dashboard de Vendas atualizado

Este guia explica **(1)** como atualizar os números do dashboard, **(2)** como o gerador funciona
e **(3)** as regras de cálculo usadas.

**Arquivos do projeto:**

| Arquivo | O que é |
|---|---|
| `index.html` (GitHub Pages) | O dashboard. Nunca muda — é o molde visual. |
| `dados.json` (GitHub Pages) | Os números atuais. **Só este arquivo muda** a cada atualização. |
| `gerar_dados_nh.py` | Script Python que gera o `dados.json` a partir dos relatórios brutos. |
| `dashboard-vendas/` | Pasta com todos os relatórios exportados das plataformas. |

O gerador lê os relatórios brutos e produz o `dados.json` — o dashboard busca esse arquivo automaticamente ao ser aberto.

---

## 1. Como executar a atualização

### Pré-requisitos
- Python 3 com a biblioteca `openpyxl` instalada (`pip install openpyxl`)
- Git configurado com acesso ao repositório `nathaliaheringer/dashboard-vendas`

### Passos

1. **Exportar os relatórios** (veja seção 2) e salvar todos na pasta `dashboard-vendas/`

2. **Atualizar as constantes de período** no topo do `gerar_dados_nh.py`:
   ```python
   PERIOD_START  = datetime(2026, 5,  1)   # início do mês
   PERIOD_END    = datetime(2026, 5, 31)   # último dia com dados
   DAYS_ELAPSED  = 31                      # dia do mês de PERIOD_END
   PERIOD_LABEL  = "01 a 31 de maio de 2026"
   DAYS_MONTH    = 31                      # dias no mês
   ```

3. **Atualizar os nomes dos arquivos** no topo do script:
   ```python
   wb = openpyxl.load_workbook(f'{BASE_DATA}/NOVO-UUID-HUBLA.xlsx')
   CART_FILE = f'{BASE_DATA}/export-leads-NOVO-TIMESTAMP.csv'
   ```
   O gerador detecta automaticamente o CSV do Hotmart e o CSV de anúncios — não precisa alterar.

4. **Executar o gerador:**
   ```bash
   python3 /Users/guilhermebasso/Documents/Claude/Projects/NH/gerar_dados_nh.py
   ```
   A saída deve mostrar `✅ Salvo: .../dados.json`

5. **Publicar no GitHub:**
   ```bash
   cd /Users/guilhermebasso/Documents/Claude/Projects/NH
   git add dados.json
   git commit -m "Dados atualizados: DD/MMM — X faturas, R$Y faturados"
   git push
   ```

O dashboard em `https://nathaliaheringer.github.io/dashboard-vendas/` estará atualizado em ~30 segundos.

---

## 2. Relatórios necessários

Coloque todos os arquivos na pasta `dashboard-vendas/` antes de rodar o gerador.

| Plataforma | Relatório | Como exportar | Nome do arquivo |
|---|---|---|---|
| **Hubla** | Relatório de Vendas | Vendas → Exportar → `.xlsx`, período do mês | Qualquer nome `.xlsx` (UUID automático) |
| **Hubla** | Relatório de Carrinho Abandonado | Carrinhos → Exportar → `.csv` | `export-leads-TIMESTAMP.csv` |
| **Hotmart** | Histórico de Vendas | Vendas → Histórico → Exportar CSV | Qualquer `.csv` com `;` como separador — **detectado automaticamente** pela coluna `Nome do Produtor` |
| **Meta Ads** | Planilha de controle RE | Google Sheets → aba DADOS RE → Exportar CSV | `[Nathalia Heringer] Controle de Vendas Perpétuos. - DADOS RE (1).csv` |
| **Meta Ads** | Planilha de controle PSI | Google Sheets → aba DADOS PSI08 → Exportar CSV | `[Nathalia Heringer] Controle de Vendas Perpétuos. - DADOS PSI08 (2).csv` |
| **Meta Ads** | Relatório de anúncios *(opcional)* | Ads Manager → nível **Anúncio** → CSV com coluna `Ad name` | Qualquer `.csv` com coluna `Ad name` — **detectado automaticamente** |

> **Detecção automática:** O gerador identifica o tipo de cada CSV pela presença de colunas-chave. Não é necessário renomear os arquivos.

---

## 3. Regras de cálculo

### Faturamento e receita líquida

- Considerar apenas faturas com status **Paga**.
- `Faturamento bruto` = soma da coluna **Valor total** (inclui order bumps na fatura).
- `Receita líquida (NH)` = soma da coluna **Valor Líquido** (após taxa Hubla).
- Uma fatura pertence ao funil **RE ou PSI** se o produto aparecer em **qualquer** das colunas `Nome do produto` **OU** `Nome do produto de orderbump`. Isso captura tanto vendas diretas quanto compras feitas como OB dentro de outro funil.
- Revenue atribuída ao produto: soma do **Valor total da fatura inteira** para todas as faturas onde o produto aparece — idêntico ao filtro por produto no Hubla.

> **Importante:** o dashboard usa `Valor total` (não `Valor do produto`) para garantir que os números batam exatamente com o relatório de vendas do Hubla.

### Hotmart
- Vendas do Hotmart são somadas ao faturamento total e ao produto correspondente.
- Canal de venda: **Hotmart** (aparece separado no gráfico de origem das vendas).
- Métodos mapeados: `PIX Automático` → PIX, `Boleto Bancário` → Boleto.

### Origem / canal de venda
- `facebookads` → **Facebook Ads** | `whatsapp` → **WhatsApp**
- `instagram` / `bio` → **Instagram (orgânico)**
- Vendas Hotmart → **Hotmart**
- UTM vazia → **Sem origem**
- Público **Frio** e **Quente** vêm do Facebook Ads (UTM de campanha contém `[Frio]` ou `[Quente]`).

### Produto da campanha (UTM)
- `[RE]` → **Regulando as Emoções** | `[PSI08]` → **Curso PSI 0-8**

### Investimento (Meta Ads)
- `Investimento em vendas` = gasto das campanhas de conversão (objetivo venda).
- A campanha `[INSTA][Seguidores]` é de audiência — fora do ROAS de vendas.
- `ROAS` = faturamento ÷ investimento | `Lucro` = receita líquida − investimento
- `ROI` = lucro ÷ investimento × 100 | `CPA` = investimento ÷ nº de vendas

### Funil de conversão por produto
- Cliques e LPV **reais** vêm da planilha de controle (Meta Ads Google Sheets).
- Esses dados são salvos também **por dia** no `daily[]`, permitindo filtro por período.
- **IGNORAR** as colunas `Purchases` e `Checkouts Initiated` da planilha do Meta — divergem das vendas reais.
- `Checkouts` e `Compras` **sempre** vêm do Hubla (e Hotmart quando aplicável).

### Order Bumps
- Contagem e valor de cada OB rastreados por dia e por produto (RE/PSI).
- `ob_detail_re` / `ob_detail_psi` dentro de cada item do `daily[]` permite filtrar OBs por período.

### Semanas
- Fixas dentro do mês: S1 = 1–7, S2 = 8–14, S3 = 15–21, S4 = 22–28, S5 = 29–31.
- Semanas sem dados ficam com `"has_data": false`.

### Métodos de pagamento e parcelamento
- Rastreados globalmente e por produto (RE/PSI), também por dia no `daily[]`.
- Parcelamento registrado apenas para Cartão de Crédito (1x–12x).

### Anúncios (Meta Ads)
- Se um CSV com coluna `Ad name` estiver em `dashboard-vendas/`, o gerador popula a tabela de anúncios.
- Cópias de anúncio (nome contém `— Cópia`) são mescladas com o anúncio original.
- Vendas estimadas por anúncio = `vendas_campanha × (clicks_anúncio / clicks_campanha)`.
- **Fonte das vendas: sempre Hubla** — o pixel do Meta não é usado para vendas.

---

## 4. Estrutura do `dados.json`

| Chave | Conteúdo |
|---|---|
| `period` | datas de início/fim, rótulo e dias decorridos |
| `totals` | faturamento, receita líquida, investimento, lucro, ROI, ROAS, etc. |
| `daily[]` | um registro por dia com ~30 campos: fat, nh, spend, faturas, origens, regiões, pagamentos, OBs, funil RE/PSI por dia |
| `weeks[]` | os 5 blocos semanais com métricas consolidadas |
| `campaigns[]` | uma entrada por campanha com detalhamento diário |
| `ads[]` | anúncios com métricas de mídia e estimativas Hubla (vazio se não houver CSV) |
| `products_paid` | funil RE e PSI: fat, nh, spend, ROAS, CPA, ticket, funil, pay_dist, parc_dist, ob_detail, weeks |
| `origins` / `fb_split` | receita por canal de venda (Facebook, Instagram, WhatsApp, Hotmart, Sem origem) |
| `products[]` | receita por produto (todos individualmente, sem agrupamento) |
| `regioes` / `estados` | receita por macrorregião e por estado |
| `followers` | campanha de seguidores (investimento, alcance, impressões) |
| `payment_dist` | distribuição global de métodos de pagamento |
| `parc_dist` | distribuição global de parcelamento (CC) |
| `ob_detail[]` | order bumps globais (nome, contagem, valor estimado) |
| `insights[]` | análises e recomendações geradas automaticamente |

> O campo `daily[]` inclui por dia: `origins`, `fb_split`, `regioes`, `pay_dist`, `parc_dist`, `ob_detail_re`, `ob_detail_psi`, `funnel_clicks_re/psi`, `funnel_lpv_re/psi`. Isso permite que **todos os gráficos da Visão Geral respondam ao filtro de período**.

---

## 5. Checklist de atualização

- [ ] Exportar XLSX do Hubla (Vendas, período do mês)
- [ ] Exportar CSV de Carrinhos Abandonados do Hubla
- [ ] Exportar CSV do Histórico de Vendas do Hotmart (se houver vendas no mês)
- [ ] Exportar CSVs das abas DADOS RE e DADOS PSI08 do Google Sheets
- [ ] *(Opcional)* Exportar relatório de anúncios do Meta Ads Manager (nível Anúncio)
- [ ] Salvar todos os arquivos em `dashboard-vendas/`
- [ ] Atualizar `PERIOD_END`, `DAYS_ELAPSED` e `PERIOD_LABEL` no `gerar_dados_nh.py`
- [ ] Atualizar nome do XLSX e do CSV de carrinhos no script
- [ ] Executar `python3 gerar_dados_nh.py` — verificar saída sem erros
- [ ] Conferir na saída: faturas RE e PSI batem com o Hubla UI
- [ ] `git add dados.json && git commit -m "..." && git push`
