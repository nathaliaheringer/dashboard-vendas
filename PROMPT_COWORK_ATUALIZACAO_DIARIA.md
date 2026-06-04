# Prompt — Atualização diária do Dashboard NH (rodar no Cowork)

Cole o bloco abaixo no seu agente do Cowork. Ele é **autossuficiente** — explica
ao agente exatamente o que fazer, em qual ordem, com quais arquivos. Foi escrito
para rodar de forma autônoma todos os dias.

---

## 📋 Prompt pronto para colar

```
Você vai atualizar o Dashboard de Vendas da Nathalia Heringer. Siga os 6
passos abaixo EM ORDEM. Não pule etapas. Se algo falhar, pare e me avise
com o erro exato — não tente “arrumar” modificando o gerador.

═══════════════════════════════════════════════════════════════
CONTEXTO DO PROJETO
═══════════════════════════════════════════════════════════════

Repositório local:      /Users/guilhermebasso/Documents/Claude/Projects/NH/
Pasta de relatórios:    /Users/guilhermebasso/Documents/Claude/Projects/NH/dashboard-vendas/
Gerador (Python):       /Users/guilhermebasso/Documents/Claude/Projects/NH/gerar_dados_nh.py
Arquivo de saída:       /Users/guilhermebasso/Documents/Claude/Projects/NH/dados.json
Dashboard publicado:    https://nathaliaheringer.github.io/dashboard-vendas/

O gerador detecta AUTOMATICAMENTE:
  • O mês corrente (sempre puxa do datetime.now() do sistema)
  • Qual XLSX da Hubla está atualizado (vai pelo arquivo mais recente que
    contenha faturas do mês corrente)
  • Qual export-leads-*.csv usar (o mais recente)
  • CSVs de DADOS RE / DADOS PSI08 (pelo nome)
  • CSV do Hotmart (pelo cabeçalho que contém "Nome do Produtor")
  • CSV de anúncios diários (pelo cabeçalho que contém "Ad name" + "Day")

Por isso, você NÃO precisa editar o gerador. Apenas garanta que os
arquivos certos estão na pasta dashboard-vendas/ e execute.

═══════════════════════════════════════════════════════════════
PASSO 1 — Baixar Relatório de Vendas da Hubla (.xlsx)
═══════════════════════════════════════════════════════════════

a) Acesse: https://app.hub.la/dashboard/sales
b) Defina o filtro de período:
   - Início: dia 01 do mês corrente
   - Fim: dia atual
c) Status: "Pagas"
d) Clique em "Exportar" → escolha XLSX
e) O arquivo virá para ~/Downloads/ com nome tipo UUID.xlsx
f) MOVA o arquivo para a pasta dashboard-vendas/:
   mv ~/Downloads/UUID.xlsx /Users/guilhermebasso/Documents/Claude/Projects/NH/dashboard-vendas/

Importante: NÃO renomeie o arquivo. NÃO delete XLSX antigos — o gerador
escolhe o mais recente sozinho, e os antigos permitem o dashboard mostrar
o histórico via seletor de datas.

═══════════════════════════════════════════════════════════════
PASSO 2 — Baixar Relatório de Carrinho Abandonado da Hubla (.csv)
═══════════════════════════════════════════════════════════════

a) Acesse: https://app.hub.la/dashboard/abandoned-carts (ou seção
   "Carrinhos abandonados")
b) Defina período: dia 01 do mês corrente até hoje
c) Clique em "Exportar" → CSV
d) Arquivo virá como export-leads-AAAA-MM-DDTHH_MM_SS.csv
e) Mova para dashboard-vendas/:
   mv ~/Downloads/export-leads-*.csv /Users/guilhermebasso/Documents/Claude/Projects/NH/dashboard-vendas/

═══════════════════════════════════════════════════════════════
PASSO 3 — Atualizar a planilha de Meta Ads (Google Sheets)
═══════════════════════════════════════════════════════════════

Esta é a fonte de spend, impressões, cliques e LPV das campanhas Meta.

a) Acesse: https://docs.google.com/spreadsheets/d/[ID_DA_PLANILHA]
   (planilha "Nathalia Heringer - Controle de Vendas Perpétuos")
b) Verifique se as abas DADOS RE e DADOS PSI08 estão com dados até
   o dia ANTERIOR ao corrente (D-1 do Meta Ads — sempre há defasagem
   de 24h no relatório do Meta).
c) Se a planilha está atualizada (alguém da equipe atualiza), pule para 3.e.
d) Se não, adicione manualmente as linhas faltantes — uma linha por
   dia × campanha, com colunas:
   Day, Campaign Name, CPM, Link Clicks, CPC, CTR, Checkouts Initiated,
   Purchases, Amount Spent, Impressions, Landing Page Views
   (puxe do Meta Ads Manager: https://adsmanager.facebook.com)
e) Exporte CADA ABA como CSV:
   - Arquivo → Fazer download → CSV (separado por vírgulas)
   - Salve com o nome exato:
     "[Nathalia Heringer] Controle de Vendas Perpétuos. - DADOS RE.csv"
     "[Nathalia Heringer] Controle de Vendas Perpétuos. - DADOS PSI08.csv"
f) SUBSTITUA os CSVs antigos em dashboard-vendas/:
   cp ~/Downloads/[Nathalia*DADOS\ RE*.csv /Users/guilhermebasso/Documents/Claude/Projects/NH/dashboard-vendas/
   cp ~/Downloads/[Nathalia*DADOS\ PSI08*.csv /Users/guilhermebasso/Documents/Claude/Projects/NH/dashboard-vendas/

═══════════════════════════════════════════════════════════════
PASSO 4 — Atualizar dados de anúncios via Meta Ads MCP
═══════════════════════════════════════════════════════════════

Os anúncios individuais (criativos) vêm do MCP do Meta Ads — não há
exportação manual; você busca direto via API.

a) Identificar o período: do dia 01 do mês corrente até hoje.
b) Para cada uma das 9 campaign_ids abaixo, fazer UMA chamada separada
   ao tool ads_get_ad_entities (não tente buscar todos de uma vez —
   o resultado fica grande demais e é truncado):

   Conta: 1611608375662839 (CA - Nathalia Heringer)

   Campanhas:
   - 120243806928760031 — RE Frio Vídeos
   - 120243995594900031 — RE Frio Estáticos
   - 120243707784390031 — RE Quente Vídeos
   - 120243702144660031 — RE Quente Estáticos
   - 120245128902040031 — RE Frio Melhores Criativos (CBO)
   - 120242497775320031 — PSI Quente Estáticos
   - 120243049034730031 — PSI Frio
   - 120243048737440031 — PSI Initiate Checkout
   - 120241138795630031 — INSTA Seguidores

   Para cada chamada, use estes parâmetros:
   {
     "ad_account_id": "1611608375662839",
     "level": "ad",
     "fields": ["id","name","campaign_id","amount_spent","impressions","clicks"],
     "filtering": [{"field":"campaign.id","operator":"IN","value":["CAMPAIGN_ID"]}],
     "time_range": {"since":"AAAA-MM-01","until":"AAAA-MM-DD"},
     "time_increment": "1",
     "limit": 200
   }

c) Se alguma chamada retornar erro "exceeds maximum allowed tokens"
   (truncamento), o output é salvo automaticamente em um arquivo .txt
   na pasta da sessão — isso é OK, o passo 4.d processa esses arquivos.

d) Após coletar todas as 9 campanhas, consolide tudo num único CSV
   chamado "ads_daily.csv". Execute este Python (ajuste se necessário):

   python3 - <<'PYEOF'
   import json, re, glob, os, csv
   from collections import defaultdict

   # Diretório onde o MCP salva os outputs truncados
   MCP_DIR = '/Users/guilhermebasso/.claude/projects/-Users-guilhermebasso-Documents-Claude-Projects-NH/'

   CAMP_MAP = {
     '120242497775320031': '[PSI08] [Compra] [Quente] [Teste Criativos] [Estáticos] - ABO',
     '120243048737440031': '[PSI08] [Initiate Checkout] [Frio] [ADV+] - ABO',
     '120243049034730031': '[PSI08] [Compra] [Frio] - ABO',
     '120243702144660031': '[RE] [Compra] [Quente] [Validação] [Estáticos] - ABO',
     '120243806928760031': '[RE] [Compra] [Frio] [Validação] [Vídeos] - ABO',
     '120243707784390031': '[RE] [Compra] [Quente] [Validação] [Vídeos] - ABO',
     '120243995594900031': '[RE] [Compra] [Frio] [Validação] [Estáticos] - ABO',
     '120245128902040031': '[RE] [Compra] [Frio] [Melhores Criativos] - CBO',
     '120241138795630031': '[INSTA] [Seguidores] [Frio] - ABO',
   }
   MESES = {'janeiro':1,'fevereiro':2,'março':3,'abril':4,'maio':5,'junho':6,
            'julho':7,'agosto':8,'setembro':9,'outubro':10,'novembro':11,'dezembro':12}
   def parse_date_pt(s):
     if not s: return ''
     m = re.match(r'(\d+)\s+de\s+(\w+)\s+de\s+(\d{4})', str(s).strip(), re.IGNORECASE)
     if not m: return ''
     d,mes,y = int(m.group(1)), m.group(2).lower(), int(m.group(3))
     return f"{y}-{MESES.get(mes,0):02d}-{d:02d}" if mes in MESES else ''
   def parse_money(s):
     if not s or s in ('Not available',''): return 0.0
     s = re.sub(r'[^\d,.]', '', str(s))
     if ',' in s and '.' in s: s = s.replace('.','').replace(',','.')
     elif ',' in s: s = s.replace(',','.')
     try: return float(s)
     except: return 0.0
     try: return float(s)
     except: return 0.0
   def parse_int(s):
     if not s or s in ('Not available',''): return 0
     s = re.sub(r'[^\d]', '', str(s))
     try: return int(s)
     except: return 0

   # 1) Extrai TODAS as entradas de ads do transcript da sessão atual
   #    (cobre tanto o que veio inline quanto o que foi salvo em arquivo)
   transcript = sorted(glob.glob(MCP_DIR + '*.jsonl'),
                       key=lambda p: os.path.getmtime(p))[-1]
   tool_files = sorted(glob.glob(MCP_DIR + '**/tool-results/*ads_get_ad_entities*.txt',
                                  recursive=True))
   all_ads = []
   seen = set()
   import json as J
   with open(transcript, encoding='utf-8') as f:
     for line in f:
       try: msg = J.loads(line)
       except: continue
       texts = []
       cont = msg.get('message',{}).get('content',[])
       if isinstance(cont,list):
         for b in cont:
           if isinstance(b,dict):
             if 'text' in b: texts.append(b['text'])
             c = b.get('content')
             if isinstance(c,str): texts.append(c)
             elif isinstance(c,list):
               for cc in c:
                 if isinstance(cc,dict) and 'text' in cc: texts.append(cc['text'])
       for t in texts:
         for m in re.finditer(r'"ad_entities"\s*:\s*"(\[.*?\])"\s*,\s*"summary"', t, re.DOTALL):
           try: ads = J.loads(J.loads('"'+m.group(1)+'"'))
           except:
             try: ads = J.loads(m.group(1))
             except: continue
           for ad in ads:
             k = (ad.get('date_start',''), ad.get('id',''))
             if not k[0] or k in seen: continue
             seen.add(k); all_ads.append(ad)
   for tf in tool_files:
     try:
       with open(tf,encoding='utf-8') as f: data = J.load(f)
       ads = J.loads(data['ad_entities'])
       for ad in ads:
         k = (ad.get('date_start',''), ad.get('id',''))
         if not k[0] or k in seen: continue
         seen.add(k); all_ads.append(ad)
     except: continue

   rows = []
   for ad in all_ads:
     date_iso = parse_date_pt(ad.get('date_start',''))
     if not date_iso: continue
     sp = parse_money(ad.get('amount_spent','0'))
     imp = parse_int(ad.get('impressions','0'))
     cl = parse_int(ad.get('clicks','0'))
     if sp == 0 and imp == 0: continue
     rows.append({
       'Day': date_iso,
       'Campaign name': CAMP_MAP.get(ad.get('campaign_id',''), ad.get('campaign_id','')),
       'Ad name': ad.get('name',''),
       'Amount Spent': sp, 'Impressions': imp, 'Reach': 0, 'Frequency': 0,
       'CPM (Cost per 1,000 Impressions)': round(sp/imp*1000,2) if imp else 0,
       'Link Clicks': cl,
       'CPC (Cost per Link Click)': round(sp/cl,2) if cl else 0,
       'CTR (Link Click-Through Rate)': round(cl/imp*100,2) if imp else 0,
     })
   rows.sort(key=lambda x:(x['Day'],x['Campaign name'],x['Ad name']))
   out = '/Users/guilhermebasso/Documents/Claude/Projects/NH/dashboard-vendas/ads_daily.csv'
   with open(out,'w',encoding='utf-8',newline='') as f:
     w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
     w.writeheader(); w.writerows(rows)
   print(f"✅ ads_daily.csv: {len(rows)} linhas")
   PYEOF

═══════════════════════════════════════════════════════════════
PASSO 5 — Executar o gerador
═══════════════════════════════════════════════════════════════

cd /Users/guilhermebasso/Documents/Claude/Projects/NH/
python3 gerar_dados_nh.py

Saída esperada (resumo):
  📅 Período detectado: 01 a XX de <MÊS> de 2026
  [Hubla] X faturas no mês corrente | Y faturas no histórico
  [Carrinhos] X | RE:Y(FB:Z), PSI:W(FB:V)
  [RE CSV] mês corrente: clicks=X lpv=Y
  [PSI08 CSV] mês corrente: clicks=X lpv=Y
  [Anúncios] X anúncios únicos | Y registros diários
  === RESUMO ===
  Fat: R$X | NH: R$Y | Lucro: R$Z | ROAS: Wx | ROI: V%
  RE: X vendas R$Y ROAS=Z | PSI: A vendas R$B ROAS=C
  ✅ Salvo: /Users/guilhermebasso/.../dados.json

CONFIRMAR antes de prosseguir:
  ☐ Período detectado bate com o mês corrente
  ☐ Faturas RE/PSI batem com o que aparece no Hubla UI
    (filtrar por produto no Hubla para conferir)
  ☐ Spend RE/PSI bate com a planilha de controle (DADOS RE/PSI08)
  ☐ Sem erros nem stack traces

Se algum desses NÃO bater, pare e me avise. Não publique dados errados.

═══════════════════════════════════════════════════════════════
PASSO 6 — Publicar no GitHub Pages
═══════════════════════════════════════════════════════════════

cd /Users/guilhermebasso/Documents/Claude/Projects/NH/
git add dados.json dashboard-vendas/ads_daily.csv
git status   # confirme que só dados.json e ads_daily.csv estão staged
git commit -m "Dados atualizados: DD/MMM — X faturas, R$Y faturados"
git push

Após 30–60s, o dashboard em https://nathaliaheringer.github.io/dashboard-vendas/
estará com os dados novos.

═══════════════════════════════════════════════════════════════
REGRAS IMPORTANTES (não esqueça)
═══════════════════════════════════════════════════════════════

1. NÃO altere o gerador (gerar_dados_nh.py) sem me consultar.
2. NÃO delete arquivos XLSX antigos da pasta dashboard-vendas/ —
   eles permitem o seletor de datas mostrar o histórico de meses
   passados no dashboard.
3. NÃO renomeie os CSVs do Meta Ads — o gerador detecta pelo nome
   exato “DADOS RE” / “DADOS PSI08”.
4. NÃO commite os arquivos da pasta dashboard-vendas/ (exceto
   ads_daily.csv). Só dados.json + ads_daily.csv vão pro Git.
5. SE a Hubla mostrar 217 vendas RE e o resumo do gerador mostrar
   diferente, pare e me avise — é divergência de dados, não algo
   pra “arrumar” no script.
6. Vendas SEMPRE vêm da Hubla (não do pixel do Meta). Meta só dá
   spend, impressões e cliques. O dashboard já é construído assim.

═══════════════════════════════════════════════════════════════
EM CASO DE ERRO
═══════════════════════════════════════════════════════════════

• "ModuleNotFoundError: openpyxl" → rodar: pip3 install openpyxl
• "Nenhum .xlsx encontrado" → o passo 1 falhou; refaça o download.
• Faturas RE/PSI divergentes do Hubla → confira se o XLSX baixado
  cobre TODO o mês (de 01 a hoje); se não, baixe de novo com o
  período correto.
• Spend zerado → CSVs de DADOS RE/PSI08 não estão na pasta ou estão
  com nome errado.
• git push falhou → "git pull --rebase" antes; se houver conflito
  em dados.json, dê preferência ao seu local (acabou de gerar):
  git checkout --ours dados.json && git rebase --continue

Quando terminar com sucesso, responda apenas:
  ✅ Dashboard atualizado: DD/MMM, X vendas RE + Y vendas PSI
```

---

## ⏰ Como agendar no Cowork

1. No Cowork, crie uma **Routine** com gatilho de cronograma.
2. Sugestão de frequência: **diária às 09:30** (depois do horário em
   que o Meta Ads disponibiliza os dados de D-1).
3. Cole o prompt acima na rotina.
4. Marque para anexar:
   - `gerar_dados_nh.py` (o script)
   - `dados.json` atual (serve de referência de formato)

## 🔄 Manutenção

Sempre que o gerador mudar (ex.: novas métricas, novas campanhas), ATUALIZE
este arquivo aqui e refaça o paste no Cowork. Mudanças importantes no
fluxo de dados sempre exigem revisar o prompt.

## 📞 Quando NÃO usar a rotina

- Dia 01 de cada mês: rode manualmente para garantir que o gerador
  detectou o mês corrente certo e confira o resumo.
- Se a Hubla ou o Meta Ads ficar fora do ar: a rotina vai falhar; ignore
  o erro e rode no dia seguinte.
- Quando aparecer um produto novo na Hubla: o gerador vai listá-lo
  automaticamente, mas confira se entrou no donut de "Receita por produto".
