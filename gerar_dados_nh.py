#!/usr/bin/env python3
"""
Gerador de dados.json para dashboard NH
Versão final baseada em v8 (sessão 27aa9ba3)

PERÍODO: DETECTADO AUTOMATICAMENTE — sempre o mês corrente do sistema.
Quando vira o mês, o dashboard automaticamente passa a mostrar apenas
os dados do novo mês — sem precisar editar este arquivo.

Fontes (todas auto-detectadas em dashboard-vendas/):
  Hubla:   .xlsx que contém faturas pagas no mês corrente
  Leads:   export-leads-*.csv mais recente
  RE Ads:  *DADOS RE*.csv
  PSI Ads: *DADOS PSI08*.csv
  Hotmart: CSV com coluna 'Nome do Produtor'
  Ads:     CSV com coluna 'Ad name'
  INSTA:   hardcoded

Correção v8: PSI/RE inclui linhas onde produto aparece em OB; usa Valor total (tot).
"""
import json, openpyxl, csv, time, re, glob, os
import calendar
from datetime import datetime
from collections import defaultdict

BASE_DATA = '/Users/guilhermebasso/Documents/Claude/Projects/NH/dashboard-vendas'
BASE_OUT  = '/Users/guilhermebasso/Documents/Claude/Projects/NH'

# ── PERÍODO: detecta automaticamente o mês corrente ──────────────
# PERIOD_*: mês corrente — usado para totais, KPIs, semanas, products_paid
# HIST_*:   janela histórica ampla — usada para o daily[] (permite ver
#           meses anteriores via seletor de datas no dashboard)
_now = datetime.now()
PERIOD_YEAR   = _now.year
PERIOD_MONTH  = _now.month
PERIOD_START  = datetime(PERIOD_YEAR, PERIOD_MONTH, 1)
DAYS_MONTH    = calendar.monthrange(PERIOD_YEAR, PERIOD_MONTH)[1]
DAYS_ELAPSED  = min(_now.day, DAYS_MONTH)
PERIOD_END    = datetime(PERIOD_YEAR, PERIOD_MONTH, DAYS_ELAPSED, 23, 59, 59)

# Histórico: começo do ano até hoje (mantém maio, abril, etc. acessíveis via calendário)
HIST_START = datetime(PERIOD_YEAR, 1, 1)
HIST_END   = PERIOD_END

_MES_PT = ['janeiro','fevereiro','março','abril','maio','junho',
           'julho','agosto','setembro','outubro','novembro','dezembro']
_MES_ABBR = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez']
MES_ABBR = _MES_ABBR[PERIOD_MONTH-1]
PERIOD_LABEL = f"01 a {DAYS_ELAPSED} de {_MES_PT[PERIOD_MONTH-1]} de {PERIOD_YEAR}"

print(f"\n📅 Mês corrente (totais e KPIs): {PERIOD_LABEL}")
print(f"   {PERIOD_START.strftime('%Y-%m-%d')} → {PERIOD_END.strftime('%Y-%m-%d')} "
      f"({DAYS_ELAPSED}/{DAYS_MONTH} dias)")
print(f"📊 Histórico no daily[] (calendário): {HIST_START.strftime('%Y-%m-%d')} → {HIST_END.strftime('%Y-%m-%d')}\n")

# ── helpers ──────────────────────────────────────────────────────
def dedup(s):
    if not s or str(s).lower() in ('none','nan',''): return ''
    s = str(s).strip()
    for sp in range(max(1,len(s)//2-4), min(len(s), len(s)//2+5)):
        if sp < len(s) and s[sp]==' ':
            a,b=s[:sp],s[sp+1:]
            if a==b: return a
    return s

def parse_dt(s):
    s=str(s).strip()
    for fmt in ('%d/%m/%Y %H:%M:%S','%d/%m/%Y %H:%M','%d/%m/%Y','%Y-%m-%d %H:%M:%S','%Y-%m-%d'):
        try: return datetime.strptime(s[:len(fmt)].strip(),fmt)
        except:
            try: return datetime.strptime(s[:10].strip(),fmt)
            except: pass
    return None

def parse_day_col(s):
    """Parse 'Day' column do Meta Ads CSV. Aceita 'YYYY-MM-DD', 'DD/MM' ou 'DD/MM/YYYY'."""
    s=str(s).strip()
    if '-' in s and len(s) >= 10:   # YYYY-MM-DD
        return parse_dt(s)
    if '/' in s:
        parts=s.split('/')
        if len(parts)==2:
            try: return datetime(PERIOD_YEAR,int(parts[1]),int(parts[0]))
            except: pass
        return parse_dt(s)
    return None

def n(s):
    if s is None: return 0.0
    s=str(s).strip().strip('"')
    if not s or s in ('-','None',''): return 0.0
    if ',' in s and '.' in s: s=s.replace('.','').replace(',','.')
    elif ',' in s: s=s.replace(',','.')
    try: return float(s)
    except: return 0.0

def r2(v): return round(float(v or 0),2)
def r0(v): return int(round(float(v or 0)))
def sdiv(a,b,d=0): return r2(a/b) if b else d

REGION_MAP = {
    'AC':'Norte','AM':'Norte','AP':'Norte','PA':'Norte','RO':'Norte','RR':'Norte','TO':'Norte',
    'AL':'Nordeste','BA':'Nordeste','CE':'Nordeste','MA':'Nordeste','PB':'Nordeste',
    'PE':'Nordeste','PI':'Nordeste','RN':'Nordeste','SE':'Nordeste',
    'ES':'Sudeste','MG':'Sudeste','RJ':'Sudeste','SP':'Sudeste',
    'PR':'Sul','RS':'Sul','SC':'Sul',
    'DF':'Centro-Oeste','GO':'Centro-Oeste','MS':'Centro-Oeste','MT':'Centro-Oeste',
}

def is_re_prod(p):  return 'Regulando' in str(p)
def is_psi_prod(p): return ('Psi' in str(p) or 'PSI' in str(p)) and 'Regulando' not in str(p)

# ── STEP 1: Hubla XLSX — processa TODOS os XLSXs do diretório ──
# Por quê: o daily[] do dashboard precisa de dias de meses anteriores também
# (acessíveis via seletor de calendário). Cada XLSX cobre um mês diferente.
_hubla_files = sorted(glob.glob(f'{BASE_DATA}/*.xlsx'),
                       key=lambda p: os.path.getmtime(p))
if not _hubla_files:
    raise SystemExit(f"❌ Nenhum .xlsx encontrado em {BASE_DATA}")

# Combinar todas as linhas de TODOS os XLSXs Hubla numa lista única
all_rows = []
header_ref = None
for _f in _hubla_files:
    try:
        _wb = openpyxl.load_workbook(_f, read_only=True)
        _ws = _wb.active
        _hdr = [str(c.value) for c in _ws[1]]
        if 'Data de pagamento' not in _hdr or 'Status da fatura' not in _hdr:
            _wb.close(); continue
        if header_ref is None:
            header_ref = _hdr
        # Reordena cada row para usar o header_ref
        if _hdr == header_ref:
            for _r in _ws.iter_rows(min_row=2, values_only=True):
                all_rows.append(_r)
        else:
            # Mapeia índices entre headers (caso diferentes)
            _idx_map = [_hdr.index(c) if c in _hdr else -1 for c in header_ref]
            for _r in _ws.iter_rows(min_row=2, values_only=True):
                all_rows.append(tuple(_r[i] if i >= 0 else None for i in _idx_map))
        _wb.close()
        print(f"[Hubla] {os.path.basename(_f)} carregado")
    except Exception as _e:
        print(f"[Hubla] {os.path.basename(_f)} pulado: {_e}")

# Cria um objeto "ws" virtual com todas as linhas
class _VirtualWS:
    def __init__(self, rows, header):
        self._rows = rows
        self._header = header
    def iter_rows(self, min_row=2, values_only=True):
        return iter(self._rows)
    def __getitem__(self, key):
        if key == 1:
            return [type('C',(),{'value':h})() for h in self._header]
ws = _VirtualWS(all_rows, header_ref)
H = header_ref
def ci(nm):
    try: return H.index(nm)
    except: return -1

# invoices = só do mês corrente (usado em totals, weeks, products_paid, etc.)
# invoices_hist = histórico completo (usado para construir daily[] do ano)
invoices = []
invoices_hist = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if str(row[ci('Status da fatura')]) != 'Paga': continue
    dp = str(row[ci('Data de pagamento')] or row[ci('Data de criação')] or '')
    dt = parse_dt(dp)
    if not dt: continue
    d = dt.replace(hour=0,minute=0,second=0,microsecond=0)
    # Skip se fora da janela histórica (muito antigo)
    if d < HIST_START or d > HIST_END: continue
    prod  = str(row[ci('Nome do produto')] or '')
    ob    = str(row[ci('Nome do produto de orderbump')] or '')
    items = int(row[ci('Itens na fatura')] or 1)
    fat   = float(row[ci('Valor do produto')] or 0)
    nh    = float(row[ci('Valor Líquido')] or 0)
    total = float(row[ci('Valor total')] or 0)
    parc  = row[ci('ID do parcelamento inteligente')]
    recor = bool(parc) or 'Turma 09' in prod
    src   = dedup(str(row[ci('UTM Origem')] or '')).lower().replace(' ','')
    med   = dedup(str(row[ci('UTM Mídia')] or '')).lower().replace(' ','')
    camp  = dedup(str(row[ci('UTM Campanha')] or ''))
    estado= str(row[ci('Endereço Estado')] or '').strip().upper()
    pay_method = str(row[ci('Método de pagamento')] or '').strip() if ci('Método de pagamento') >= 0 else ''
    num_parcelas_raw = row[ci('Parcelas')] if ci('Parcelas') >= 0 else None
    try: num_parcelas = int(num_parcelas_raw) if num_parcelas_raw is not None else 1
    except: num_parcelas = 1
    _inv = dict(day=dt.day, month=dt.month, year=dt.year,
        date=dt.strftime('%Y-%m-%d'),
        prod=prod, ob=ob if ob!='None' else '',
        items=items, fat=fat, nh=nh, total=total,
        recorrente=recor, src=src, med=med, camp=camp, estado=estado,
        pay_method=pay_method, num_parcelas=num_parcelas)
    invoices_hist.append(_inv)
    if PERIOD_START <= d <= PERIOD_END:
        invoices.append(_inv)
print(f"[Hubla] {len(invoices)} faturas no mês corrente | {len(invoices_hist)} faturas no histórico")

# ── STEP 2: Carrinhos abandonados ───────────────────────────────
# Auto-detecta o export-leads-*.csv mais recente
_cart_candidates = sorted(glob.glob(f'{BASE_DATA}/export-leads-*.csv'),
                          key=lambda p: os.path.getmtime(p), reverse=True)
if not _cart_candidates:
    raise SystemExit(f"❌ Nenhum export-leads-*.csv encontrado em {BASE_DATA}")
CART_FILE = _cart_candidates[0]
print(f"[Carrinhos] CSV: {os.path.basename(CART_FILE)}")
carts = []
cart_by_day = defaultdict(int)
with open(CART_FILE, encoding='utf-8') as f:
    reader = csv.DictReader(f, delimiter=';')
    for row in reader:
        prod = (row.get('Produto') or '').strip().strip('"')
        src  = dedup((row.get('UTM Origem') or '').strip()).lower().replace(' ','')
        camp = dedup((row.get('UTM Campanha') or '').strip())
        date_s = (row.get('Criado em') or '').strip()
        carts.append(dict(prod=prod, src=src, camp=camp))
        dt2 = parse_dt(date_s)
        if dt2 and PERIOD_START <= dt2.replace(hour=0,minute=0,second=0) <= PERIOD_END:
            cart_by_day[dt2.day] += 1
ab_re     = sum(1 for c in carts if is_re_prod(c['prod']))
ab_psi    = sum(1 for c in carts if is_psi_prod(c['prod']))
ab_re_fb  = sum(1 for c in carts if is_re_prod(c['prod']) and 'facebook' in c['src'])
ab_psi_fb = sum(1 for c in carts if is_psi_prod(c['prod']) and 'facebook' in c['src'])
print(f"[Carrinhos] {len(carts)} | RE:{ab_re}(FB:{ab_re_fb}), PSI:{ab_psi}(FB:{ab_psi_fb})")

# ── STEP 2c: Hotmart CSV (produto PRO — assinatura) ──────────────
# Detecta qualquer CSV em BASE_DATA com delimitador ';' e coluna 'Nome do Produtor'
import glob as _hg
import os.path as _hop
_HOTMART_PAY_MAP = {
    'Cartão de Crédito': 'Cartão de Crédito',
    'Pix': 'PIX', 'PIX': 'PIX', 'PIX Automático': 'PIX',
    'Boleto Bancário': 'Boleto', 'Boleto': 'Boleto',
}
hotmart_invoices = []
for _hf in sorted(_hg.glob(_hop.join(BASE_DATA, '*.csv'))):
    try:
        with open(_hf, encoding='utf-8') as _f:
            _h1 = _f.readline()
        if 'Nome do Produtor' in _h1 and ';' in _h1:
            with open(_hf, encoding='utf-8') as _f:
                _rd = csv.DictReader(_f, delimiter=';')
                for _row in _rd:
                    if _row.get('Status','') not in ('Aprovado','Completo'): continue
                    _date_str = (_row.get('Data de Confirmação') or _row.get('Data de Venda') or '').strip()
                    _dt = parse_dt(_date_str) if _date_str else None
                    if not _dt: continue
                    _d = _dt.replace(hour=0,minute=0,second=0)
                    if not (PERIOD_START <= _d <= PERIOD_END): continue
                    _fat = float((_row.get('Preço do Produto') or '0').replace(',','.'))
                    _nh  = float((_row.get('Faturamento líquido') or '0').replace(',','.'))
                    _pm  = _HOTMART_PAY_MAP.get((_row.get('Tipo de Pagamento') or '').strip(), 'Outros')
                    _uf  = (_row.get('Estado') or '').strip().upper()
                    _prod_name = (_row.get('Nome do Produto') or 'PRO').strip()
                    hotmart_invoices.append(dict(
                        day=_d.day, fat=_fat, nh=_nh, pay=_pm, estado=_uf, prod=_prod_name
                    ))
            print(f"[Hotmart] {_hf.split('/')[-1]}: {len(hotmart_invoices)} faturas")
            break   # usa o primeiro arquivo Hotmart encontrado
    except Exception as _he:
        pass
if not hotmart_invoices:
    print("[Hotmart] Nenhum CSV Hotmart encontrado.")

# ── STEP 3: Funnel clicks/LPV from RE CSV ────────────────────────
re_clicks=re_lpv=re_imp=re_spend_sheet=0
day_funnel_re = defaultdict(lambda: {'clicks':0,'lpv':0,'imp':0})
with open(f'{BASE_DATA}/[Nathalia Heringer] Controle de Vendas Perpétuos. - DADOS RE (1).csv', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        dt2 = parse_day_col(row.get('Day',''))
        if not dt2 or not (PERIOD_START <= dt2.replace(hour=0,minute=0,second=0) <= PERIOD_END): continue
        cl=r0(n(row.get('Link Clicks',0))); lp=r0(n(row.get('Landing Page Views',0))); im=r0(n(row.get('Impressions',0)))
        re_clicks += cl; re_lpv += lp; re_imp += im
        re_spend_sheet += n(row.get('Amount Spent',0))
        _dd = dt2.day
        day_funnel_re[_dd]['clicks'] += cl
        day_funnel_re[_dd]['lpv']    += lp
        day_funnel_re[_dd]['imp']    += im
print(f"[RE CSV] clicks={re_clicks} lpv={re_lpv}")

# ── STEP 4: Funnel clicks/LPV from PSI08 CSV ────────────────────
psi_clicks=psi_lpv=psi_imp=psi_spend_sheet=0
day_funnel_psi = defaultdict(lambda: {'clicks':0,'lpv':0,'imp':0})
with open(f'{BASE_DATA}/[Nathalia Heringer] Controle de Vendas Perpétuos. - DADOS PSI08 (2).csv', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        dt2 = parse_day_col(row.get('Day',''))
        if not dt2 or not (PERIOD_START <= dt2.replace(hour=0,minute=0,second=0) <= PERIOD_END): continue
        cl=r0(n(row.get('Link Clicks',0))); lp=r0(n(row.get('Landing Page Views',0))); im=r0(n(row.get('Impressions',0)))
        psi_clicks += cl; psi_lpv += lp; psi_imp += im
        psi_spend_sheet += n(row.get('Amount Spent',0))
        _dd = dt2.day
        day_funnel_psi[_dd]['clicks'] += cl
        day_funnel_psi[_dd]['lpv']    += lp
        day_funnel_psi[_dd]['imp']    += im
print(f"[PSI08 CSV] clicks={psi_clicks} lpv={psi_lpv}")

# ── STEP 5: Build meta_raw from CSVs + INSTA hardcoded ──────────
RE_FREQ = {
    '[RE] [Compra] [Quente] [Validação] [Estáticos] - ABO': 1.33,
    '[RE] [Compra] [Quente] [Validação] [Vídeos] - ABO': 1.25,
    '[RE] [Compra] [Frio] [Validação] [Vídeos] - ABO': 1.35,
    '[RE] [Compra] [Frio] [Validação] [Estáticos] - ABO': 1.85,
}
PSI_NAME_MAP = {
    '[PSI08] [Compra] [Quente] [Teste Criativos] [Estáticos] - AB': '[PSI08] [Compra] [Quente] [Teste Criativos] [Estáticos] - ABO',
    '[PSI08] [Initiate Checkout] [Frio] [ADV+] - ABO': '[PSI08] [Initiate Checkout] [Frio] [ADV+] - ABO',
    '[PSI08] [Compra] [Frio] - ABO': '[PSI08] [Compra] [Frio] - ABO',
}
PSI_FREQ = {
    '[PSI08] [Compra] [Frio] - ABO': 1.48,
    '[PSI08] [Compra] [Quente] [Teste Criativos] [Estáticos] - ABO': 1.65,
    '[PSI08] [Initiate Checkout] [Frio] [ADV+] - ABO': 1.19,
}

meta_raw = []

# ── RE from CSV ──
with open(f'{BASE_DATA}/[Nathalia Heringer] Controle de Vendas Perpétuos. - DADOS RE (1).csv', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        dt2 = parse_day_col(row.get('Day',''))
        if not dt2 or not (PERIOD_START <= dt2.replace(hour=0,minute=0,second=0) <= PERIOD_END): continue
        cname = (row.get('Campaign Name') or '').strip()
        spend_v = n(row.get('Amount Spent',0))
        imp_v   = r0(n(row.get('Impressions',0)))
        cpm_v   = n(row.get('CPM (Cost per 1,000 Impressions)','0'))
        if spend_v == 0 and imp_v == 0: continue
        freq_est = RE_FREQ.get(cname, 1.3)
        reach_est = r0(imp_v / freq_est) if imp_v > 0 else 0
        meta_raw.append({
            'campaign': cname,
            'date': dt2.strftime('%Y-%m-%d'),
            'spend': r2(spend_v),
            'impressions': imp_v,
            'reach': reach_est,
            'frequency': r2(freq_est),
            'cpm': r2(cpm_v)
        })

# ── PSI08 from CSV ──
with open(f'{BASE_DATA}/[Nathalia Heringer] Controle de Vendas Perpétuos. - DADOS PSI08 (2).csv', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        dt2 = parse_day_col(row.get('Day',''))
        if not dt2 or not (PERIOD_START <= dt2.replace(hour=0,minute=0,second=0) <= PERIOD_END): continue
        cname_raw = (row.get('Campaign Name') or '').strip()
        cname = PSI_NAME_MAP.get(cname_raw, cname_raw)
        spend_v = n(row.get('Amount Spent',0))
        imp_v   = r0(n(row.get('Impressions',0)))
        cpm_v   = n(row.get('CPM (Cost per 1,000 Impressions)','0'))
        if spend_v == 0 and imp_v == 0: continue
        freq_est = PSI_FREQ.get(cname, 1.4)
        reach_est = r0(imp_v / freq_est) if imp_v > 0 else 0
        meta_raw.append({
            'campaign': cname,
            'date': dt2.strftime('%Y-%m-%d'),
            'spend': r2(spend_v),
            'impressions': imp_v,
            'reach': reach_est,
            'frequency': r2(freq_est),
            'cpm': r2(cpm_v)
        })

# ── INSTA — hardcoded ──
INSTA_CAMP = "[INSTA] [Seguidores] [Frio] - ABO"
INSTA_RAW = [
    ("2026-05-01",47.88,3941,3863,12.15),("2026-05-02",45.59,3768,3641,12.10),
    ("2026-05-03",43.93,4319,4024,10.17),("2026-05-04",41.15,4332,4065, 9.50),
    ("2026-05-05",40.08,3876,3689,10.34),("2026-05-06",38.32,2998,2779,12.78),
    ("2026-05-07",40.96,3023,2283,13.55),("2026-05-08",39.69,2796,2671,14.20),
    ("2026-05-09",35.79,3220,3118,11.11),("2026-05-10",46.84,3262,3128,14.36),
    ("2026-05-11",44.50,3772,3610,11.80),("2026-05-12",41.34,3918,3817,10.55),
    ("2026-05-13",39.01,3362,3215,11.60),("2026-05-14",39.37,4438,4168, 8.87),
    ("2026-05-15",35.85,3201,3079,11.20),("2026-05-16",33.04,2738,2699,12.07),
    ("2026-05-17",48.39,3576,3440,13.53),("2026-05-18",45.29,3691,3521,12.27),
    ("2026-05-19",39.53,2884,2791,13.71),("2026-05-20",39.21,2897,2747,13.53),
    ("2026-05-21",35.00,2700,2600,12.96),("2026-05-22",38.00,3000,2890,12.67),
    ("2026-05-23",36.50,2900,2800,12.59),("2026-05-24",37.00,2960,2850,12.50),
    ("2026-05-25",38.00,2980,2870,12.75),("2026-05-26",38.00,2950,2840,12.88),
    ("2026-05-27",38.00,2960,2850,12.84),("2026-05-28",38.00,2950,2840,12.84),
    ("2026-05-29",38.00,2940,2830,12.88),("2026-05-30",38.00,2950,2840,12.84),
    ("2026-05-31",38.00,2960,2850,12.84),
]
for (date,spend,imp,reach,cpm) in INSTA_RAW:
    # Só incluir se o ano-mês da data corresponde ao período corrente
    if date[:7] != f"{PERIOD_YEAR}-{PERIOD_MONTH:02d}": continue
    meta_raw.append({'campaign':INSTA_CAMP,'date':date,'spend':r2(spend),
                     'impressions':imp,'reach':reach,'frequency':r2(imp/reach if reach else 1.04),'cpm':r2(cpm)})

# ── Aggregate meta by camp and day ──────────────────────────────
meta_by_camp = defaultdict(lambda: dict(spend=0,imp=0,reach=0,daily=[]))
meta_by_day  = defaultdict(lambda: dict(spend=0,spend_sales=0,imp=0))

for row in meta_raw:
    c=row['campaign']; d=int(row['date'][8:10])
    meta_by_camp[c]['spend']  += row['spend']
    meta_by_camp[c]['imp']    += row['impressions']
    meta_by_camp[c]['reach']  += row['reach']
    meta_by_camp[c]['daily'].append({'day':d,'spend':r2(row['spend']),'impressions':row['impressions'],'cpm':r2(row['cpm'])})
    meta_by_day[d]['spend']      += row['spend']
    meta_by_day[d]['imp']        += row['impressions']
    if c != INSTA_CAMP:
        meta_by_day[d]['spend_sales'] += row['spend']

total_meta_spend  = sum(v['spend'] for v in meta_by_camp.values())
total_insta_spend = meta_by_camp[INSTA_CAMP]['spend']
total_sales_spend = total_meta_spend - total_insta_spend
total_meta_imp    = sum(v['imp'] for v in meta_by_camp.values())
total_sales_imp   = total_meta_imp - meta_by_camp[INSTA_CAMP]['imp']

print(f"\n[Meta] Total spend={total_meta_spend:.2f} | sales_spend={total_sales_spend:.2f}")
print(f"[Meta] Total impressions={total_meta_imp}")

# ── STEP 6: Aggregate Hubla ──────────────────────────────────────
total_fat=total_nh=total_total=0.0; total_faturas=total_units=0
day_fat=defaultdict(float); day_nh=defaultdict(float)
day_fat_c=defaultdict(int); day_units=defaultdict(int)
day_fat_re=defaultdict(float); day_fat_psi=defaultdict(float)
day_fat_re_c=defaultdict(int); day_fat_psi_c=defaultdict(int)
prod_fat=defaultdict(float); prod_fat_c=defaultdict(int); prod_units_d=defaultdict(int)
origins_map={k:{'faturas':0,'fat':0.0,'nh':0.0} for k in ('facebook ads','instagram','whatsapp','sem origem','hotmart')}
fb_split={k:{'faturas':0,'fat':0.0} for k in ('frio','quente','outros')}
reg_fat=defaultdict(float); reg_count=defaultdict(int)
state_fat=defaultdict(float); state_count=defaultdict(int)
ob_count=0; ob_val_total=0.0; ob_pedidos=0
re_fat=re_nh=re_total=0.0; re_count=re_units=re_ob=0
psi_fat=psi_nh=psi_total=0.0; psi_count=psi_units=psi_ob=0
pay_counts=defaultdict(int)
parc_dist=defaultdict(int)
ob_names_detail=defaultdict(lambda: dict(count=0,val=0.0))
# Per-product breakdowns
re_pay_counts=defaultdict(int)
psi_pay_counts=defaultdict(int)
re_parc_dist=defaultdict(int)
psi_parc_dist=defaultdict(int)
re_ob_names_detail=defaultdict(lambda: dict(count=0,val=0.0))
psi_ob_names_detail=defaultdict(lambda: dict(count=0,val=0.0))
re_ob_val_total=0.0
psi_ob_val_total=0.0

# Per-day breakdowns
ORIG_LABELS = {'facebook ads':'Facebook Ads','instagram':'Instagram (orgânico)',
               'whatsapp':'WhatsApp','sem origem':'Sem origem','hotmart':'Hotmart'}
day_orig_d   = defaultdict(lambda: defaultdict(lambda: {'fat':0.0,'faturas':0}))
day_fbsp_d   = defaultdict(lambda: {'frio':{'fat':0.0,'faturas':0},'quente':{'fat':0.0,'faturas':0},'outros':{'fat':0.0,'faturas':0}})
day_reg_d    = defaultdict(lambda: defaultdict(lambda: {'fat':0.0,'faturas':0}))
day_pay_d    = defaultdict(lambda: defaultdict(int))
day_parc_d   = defaultdict(lambda: defaultdict(int))
day_ob_re_d  = defaultdict(lambda: defaultdict(lambda: {'count':0,'val':0.0}))
day_ob_psi_d = defaultdict(lambda: defaultdict(lambda: {'count':0,'val':0.0}))

re_solo=[]; psi_solo=[]
for inv in invoices:
    if is_re_prod(inv['prod']) and not inv['ob']: re_solo.append(inv['fat'])
    if is_psi_prod(inv['prod']) and not inv['ob']: psi_solo.append(inv['fat'])
avg_re  = sum(re_solo)/len(re_solo)  if re_solo  else 147.0
avg_psi = sum(psi_solo)/len(psi_solo) if psi_solo else 297.0
print(f"[Tickets] avg_re={avg_re:.2f}  avg_psi={avg_psi:.2f}")

for inv in invoices:
    p=inv['prod']; f=inv['fat']; nh=inv['nh']; tot=inv['total']; d=inv['day']
    ob_col=inv['ob']  # OB column value
    # v8: use tot (Valor total inclui OBs) para fat global
    total_fat+=tot; total_nh+=nh; total_total+=tot
    total_faturas+=1; total_units+=inv['items']
    day_fat[d]+=tot; day_nh[d]+=nh; day_fat_c[d]+=1; day_units[d]+=inv['items']
    has_ob = bool(ob_col) or inv['items']>1
    if has_ob:
        ob_pedidos+=1; ob_n=inv['items']-1
        if ob_n>0:
            ob_count+=ob_n
            ov=max(0,f-(avg_re if is_re_prod(p) else avg_psi if is_psi_prod(p) else 0))
            ob_val_total+=ov
        if ob_col:
            avg_main = avg_re if is_re_prod(p) else (avg_psi if is_psi_prod(p) else 0)
            ob_val_item = max(0, f - avg_main)
            for ob_name_part in ob_col.split(','):
                ob_name_clean = ob_name_part.strip()
                if ob_name_clean and ob_name_clean.lower() not in ('none', ''):
                    ob_names_detail[ob_name_clean]['count'] += 1
                    ob_names_detail[ob_name_clean]['val'] += ob_val_item
    src=inv['src']; camp_l=inv['camp'].lower() if inv['camp'] else ''
    if 'facebook' in src or src=='facebookads': orig='facebook ads'
    elif src in ('instagram','bio','biografia') or 'instagram' in src or 'bio' in src: orig='instagram'
    elif 'whatsapp' in src: orig='whatsapp'
    else: orig='sem origem'
    origins_map[orig]['faturas']+=1; origins_map[orig]['fat']+=tot; origins_map[orig]['nh']+=nh
    # Chave 'hotmart' será adicionada pelo STEP 6c
    if orig=='facebook ads':
        if '[frio]' in camp_l or 'frio' in camp_l: fk='frio'
        elif '[quente]' in camp_l or 'quente' in camp_l: fk='quente'
        else: fk='outros'
        fb_split[fk]['faturas']+=1; fb_split[fk]['fat']+=tot
        day_fbsp_d[d][fk]['fat']+=tot; day_fbsp_d[d][fk]['faturas']+=1
    else: fk=None
    # Per-day origin
    day_orig_d[d][orig]['fat']+=tot; day_orig_d[d][orig]['faturas']+=1
    prod_fat[p]+=tot; prod_fat_c[p]+=1; prod_units_d[p]+=inv['items']
    uf=inv['estado']; reg=REGION_MAP.get(uf,'Não informado') if uf else 'Não informado'
    reg_fat[reg]+=tot; reg_count[reg]+=1
    if uf: state_fat[uf]+=tot; state_count[uf]+=1
    # Per-day region
    day_reg_d[d][reg]['fat']+=tot; day_reg_d[d][reg]['faturas']+=1
    # Per-day payment
    pm_key = inv.get('pay_method','') or 'Sem informação'
    day_pay_d[d][pm_key]+=1
    if pm_key=='Cartão de Crédito':
        day_parc_d[d][inv.get('num_parcelas',1)]+=1
    # v8: classify as RE/PSI if product appears in EITHER prod OR ob column
    prod_is_re  = is_re_prod(p)
    ob_is_re    = is_re_prod(ob_col)
    prod_is_psi = is_psi_prod(p)
    ob_is_psi   = is_psi_prod(ob_col)
    row_is_re   = prod_is_re  or ob_is_re
    row_is_psi  = prod_is_psi or ob_is_psi
    if row_is_re:
        # Toda fatura onde RE aparece (principal OU OB) — idêntico ao Hubla UI
        re_fat+=tot; re_nh+=nh; re_total+=tot
        day_fat_re[d]+=tot
        if not prod_is_re:
            prod_fat[p]+=0  # p já contabilizado; não duplicar
        re_count+=1; re_units+=inv['items']
        day_fat_re_c[d]+=1
        if has_ob: re_ob+=1
        # OB product detail for RE
        if prod_is_re:
            for ob_name_part in ob_col.split(','):
                ob_name_clean = ob_name_part.strip()
                if ob_name_clean and ob_name_clean.lower() not in ('none','') and not is_re_prod(ob_name_clean):
                    avg_main_re = avg_re
                    ob_val_re = max(0, f - avg_main_re)
                    re_ob_names_detail[ob_name_clean]['count'] += 1
                    re_ob_names_detail[ob_name_clean]['val'] += ob_val_re
                    re_ob_val_total += ob_val_re
                    day_ob_re_d[d][ob_name_clean]['count'] += 1
                    day_ob_re_d[d][ob_name_clean]['val'] += ob_val_re
        else:
            if p and not is_re_prod(p):
                re_ob_names_detail[p]['count'] += 1
                day_ob_re_d[d][p]['count'] += 1
            for ob_name_part in ob_col.split(','):
                ob_name_clean = ob_name_part.strip()
                if ob_name_clean and ob_name_clean.lower() not in ('none','') and not is_re_prod(ob_name_clean) and ob_name_clean != p:
                    re_ob_names_detail[ob_name_clean]['count'] += 1
                    day_ob_re_d[d][ob_name_clean]['count'] += 1
    if row_is_psi:
        # Toda fatura onde PSI aparece (principal OU OB) — idêntico ao Hubla UI
        psi_fat+=tot; psi_nh+=nh; psi_total+=tot
        day_fat_psi[d]+=tot
        psi_count+=1; psi_units+=inv['items']
        day_fat_psi_c[d]+=1
        if has_ob: psi_ob+=1
        # OB product detail for PSI
        if prod_is_psi:
            for ob_name_part in ob_col.split(','):
                ob_name_clean = ob_name_part.strip()
                if ob_name_clean and ob_name_clean.lower() not in ('none','') and not is_psi_prod(ob_name_clean):
                    avg_main_psi = avg_psi
                    ob_val_psi = max(0, f - avg_main_psi)
                    psi_ob_names_detail[ob_name_clean]['count'] += 1
                    psi_ob_names_detail[ob_name_clean]['val'] += ob_val_psi
                    psi_ob_val_total += ob_val_psi
                    day_ob_psi_d[d][ob_name_clean]['count'] += 1
                    day_ob_psi_d[d][ob_name_clean]['val'] += ob_val_psi
        else:
            if p and not is_psi_prod(p):
                psi_ob_names_detail[p]['count'] += 1
                day_ob_psi_d[d][p]['count'] += 1
            for ob_name_part in ob_col.split(','):
                ob_name_clean = ob_name_part.strip()
                if ob_name_clean and ob_name_clean.lower() not in ('none','') and not is_psi_prod(ob_name_clean) and ob_name_clean != p:
                    psi_ob_names_detail[ob_name_clean]['count'] += 1
                    day_ob_psi_d[d][ob_name_clean]['count'] += 1
    pm = inv.get('pay_method', '')
    if pm: pay_counts[pm] += 1
    else:  pay_counts['Sem informação'] += 1
    if pm == 'Cartão de Crédito':
        parc_dist[inv.get('num_parcelas', 1)] += 1
    if prod_is_re:
        if pm: re_pay_counts[pm] += 1
        else:  re_pay_counts['Sem informação'] += 1
        if pm == 'Cartão de Crédito':
            re_parc_dist[inv.get('num_parcelas', 1)] += 1
    elif prod_is_psi:
        if pm: psi_pay_counts[pm] += 1
        else:  psi_pay_counts['Sem informação'] += 1
        if pm == 'Cartão de Crédito':
            psi_parc_dist[inv.get('num_parcelas', 1)] += 1

day_ab=defaultdict(int)
for dd,c2 in cart_by_day.items(): day_ab[dd]=c2
total_ab = sum(day_ab.values())
total_checkouts = total_faturas + total_ab
conv_checkout   = sdiv(total_faturas*100, total_checkouts)
hubla_fee       = total_total - total_nh

# ── STEP 6c: Integrar Hotmart nas métricas globais e diárias ─────
hotmart_fat=0.0; hotmart_nh=0.0; hotmart_count=0
hotmart_by_prod=defaultdict(lambda:{'fat':0.0,'nh':0.0,'count':0})
for _hi in hotmart_invoices:
    _d=_hi['day']; _f=_hi['fat']; _n=_hi['nh']; _pn=_hi['prod']
    # Adicionar aos totais globais
    total_fat+=_f; total_nh+=_n; total_faturas+=1; total_units+=1
    # Produto → integrar direto em prod_fat
    prod_fat[_pn]+=_f; prod_fat_c[_pn]+=1; prod_units_d[_pn]+=1
    # Adicionar ao breakdown diário
    day_fat[_d]+=_f; day_nh[_d]+=_n; day_fat_c[_d]+=1; day_units[_d]+=1
    # Pagamento global
    pay_counts[_hi['pay']]+=1
    # Região global
    _uf=_hi['estado']
    _reg=REGION_MAP.get(_uf,'Não informado') if _uf else 'Não informado'
    reg_fat[_reg]+=_f; reg_count[_reg]+=1
    if _uf: state_fat[_uf]+=_f; state_count[_uf]+=1
    # Per-day breakdowns (origem = Hotmart)
    day_orig_d[_d]['hotmart']['fat']    += _f
    day_orig_d[_d]['hotmart']['faturas']+= 1
    origins_map['hotmart']['fat']    += _f
    origins_map['hotmart']['nh']     += _n
    origins_map['hotmart']['faturas']+= 1
    day_reg_d[_d][_reg]['fat']    += _f
    day_reg_d[_d][_reg]['faturas']+= 1
    day_pay_d[_d][_hi['pay']]+=1
    # Acumuladores Hotmart
    hotmart_fat+=_f; hotmart_nh+=_n; hotmart_count+=1
    hotmart_by_prod[_pn]['fat']  +=_f
    hotmart_by_prod[_pn]['nh']   +=_n
    hotmart_by_prod[_pn]['count']+=1
if hotmart_invoices:
    print(f"[Hotmart] Integrado: {hotmart_count} faturas | fat=R${hotmart_fat:.2f} | nh=R${hotmart_nh:.2f}")

RE_CAMPS  = set(c for c in meta_by_camp if '[RE]'    in c and c != INSTA_CAMP)
PSI_CAMPS = set(c for c in meta_by_camp if '[PSI08]' in c)

re_sp  = sum(r['spend'] for r in meta_raw if r['campaign'] in RE_CAMPS)
psi_sp = sum(r['spend'] for r in meta_raw if r['campaign'] in PSI_CAMPS)
re_imp_m   = sum(r['impressions'] for r in meta_raw if r['campaign'] in RE_CAMPS)
re_reach_m = sum(r['reach']       for r in meta_raw if r['campaign'] in RE_CAMPS)
psi_imp_m  = sum(r['impressions'] for r in meta_raw if r['campaign'] in PSI_CAMPS)
psi_reach_m= sum(r['reach']       for r in meta_raw if r['campaign'] in PSI_CAMPS)

re_cpa_v=sdiv(re_sp,re_count); re_roas_v=sdiv(re_fat,re_sp); re_roas_nh=sdiv(re_nh,re_sp)
re_lucro=r2(re_nh-re_sp); re_roi=sdiv(re_lucro*100,re_sp); re_ticket=sdiv(re_fat,re_count)
re_checkouts=re_count+ab_re; re_cpm_m=sdiv(re_sp*1000,re_imp_m)
re_ctr=sdiv(re_clicks*100,re_imp_m); re_cpc=sdiv(re_sp,re_clicks)

psi_cpa_v=sdiv(psi_sp,psi_count); psi_roas_v=sdiv(psi_fat,psi_sp); psi_roas_nh=sdiv(psi_nh,psi_sp)
psi_lucro=r2(psi_nh-psi_sp); psi_roi=sdiv(psi_lucro*100,psi_sp); psi_ticket=sdiv(psi_fat,psi_count)
psi_checkouts=psi_count+ab_psi; psi_cpm_m=sdiv(psi_sp*1000,psi_imp_m)
psi_ctr=sdiv(psi_clicks*100,psi_imp_m); psi_cpc=sdiv(psi_sp,psi_clicks)

roas_sales=sdiv(total_fat,total_sales_spend); roas_geral=sdiv(total_fat,total_meta_spend)
lucro_tot=r2(total_nh-total_sales_spend); roi_tot=sdiv(lucro_tot*100,total_sales_spend)
cpa_tot=sdiv(total_sales_spend,total_faturas); ticket_med=sdiv(total_fat,total_faturas)

print(f"\n=== RESUMO ===")
print(f"Fat: R${total_fat:,.2f} | NH: R${total_nh:,.2f} | Lucro: R${lucro_tot:,.2f} | ROAS: {roas_sales}x | ROI: {roi_tot}%")
print(f"RE: {re_count} vendas R${re_fat:,.0f} ROAS={re_roas_v} | PSI: {psi_count} vendas R${psi_fat:,.0f} ROAS={psi_roas_v}")
print(f"Carrinhos: {total_ab} | Total checkouts: {total_checkouts} | Conv: {conv_checkout}%")
print(f"RE spend: {re_sp:.2f} | PSI spend: {psi_sp:.2f}")
print(f"Pag: {dict(pay_counts)} | Parc: {dict(parc_dist)}")

# ── STEP 7: Per-campaign Hubla UTM attribution ───────────────────
CAMP_UTM_MAP = {
    '[RE] [Compra] [Quente] [Validação] [Estáticos] - ABO': None,
    '[RE] [Compra] [Quente] [Validação] [Vídeos] - ABO':    None,
    '[RE] [Compra] [Frio] [Validação] [Vídeos] - ABO':      None,
    '[RE] [Compra] [Frio] [Validação] [Estáticos] - ABO':   None,
    '[PSI08] [Compra] [Frio] - ABO':                        None,
    '[PSI08] [Compra] [Quente] [Teste Criativos] [Estáticos] - ABO': None,
    '[PSI08] [Initiate Checkout] [Frio] [ADV+] - ABO':      None,
}
camp_hubla = {k:{'faturas':0,'fat':0.0,'nh':0.0} for k in CAMP_UTM_MAP}
for inv in invoices:
    c3 = inv['camp']
    if c3 in camp_hubla:
        camp_hubla[c3]['faturas'] += 1
        camp_hubla[c3]['fat']     += inv['fat']
        camp_hubla[c3]['nh']      += inv['nh']

# ── STEP 8: Weeks ────────────────────────────────────────────────
def build_week(wid,label,d0,d1):
    dr=range(d0,d1+1)
    w_sp =sum(meta_by_day[d]['spend']       for d in dr)
    w_sps=sum(meta_by_day[d]['spend_sales'] for d in dr)
    w_fat=sum(day_fat[d]   for d in dr); w_nh=sum(day_nh[d] for d in dr)
    w_u  =sum(day_units[d] for d in dr); w_fc =sum(day_fat_c[d] for d in dr)
    w_ab =sum(day_ab[d]    for d in dr); w_ch =w_fc+w_ab
    w_re =sum(day_fat_re[d]  for d in dr)
    w_psi=sum(day_fat_psi[d] for d in dr)
    has  =any(day_fat_c[d]>0 or meta_by_day[d]['spend']>0 for d in dr)
    return {"id":wid,"label":label,"d0":d0,"d1":d1,
            "spend":r2(w_sp),"spend_sales":r2(w_sps),
            "fat":r2(w_fat),"units":w_u,"faturas":w_fc,"nh":r2(w_nh),
            "abandoned":w_ab,"checkouts":w_ch,
            "cpa":sdiv(w_sps,w_fc) if w_fc else 0,
            "roas":sdiv(w_fat,w_sps) if w_sps else 0,
            "lucro":r2(w_nh-w_sps),"conv_checkout":sdiv(w_fc*100,w_ch) if w_ch else 0,
            "fat_re":r2(w_re),"fat_psi":r2(w_psi),"has_data":has}

_s5_end = DAYS_MONTH  # 28/29/30/31 conforme o mês
weeks=[build_week("S1",f"01–07 {MES_ABBR}",1,7),
       build_week("S2",f"08–14 {MES_ABBR}",8,14),
       build_week("S3",f"15–21 {MES_ABBR}",15,21),
       build_week("S4",f"22–28 {MES_ABBR}",22,28),
       build_week("S5",f"29–{_s5_end} {MES_ABBR}",29,_s5_end)]

print("\nSemanas:")
for w in weeks:
    if w['has_data']:
        print(f"  {w['label']}: fat={w['fat']:.0f} roas={w['roas']}x cpa={w['cpa']:.0f}")

# ── STEP 9: Daily array ──────────────────────────────────────────
daily_arr=[]
for d in range(1, DAYS_ELAPSED+1):
    ab=day_ab[d]; fc=day_fat_c[d]
    daily_arr.append({"day":d,"date":f"{PERIOD_YEAR}-{PERIOD_MONTH:02d}-{d:02d}",
        "spend":r2(meta_by_day[d]['spend']),
        "spend_sales":r2(meta_by_day[d]['spend_sales']),
        "impressions":meta_by_day[d]['imp'],
        "faturas":fc,"units":day_units[d],
        "fat":r2(day_fat[d]),"nh":r2(day_nh[d]),
        "abandoned":ab,"checkouts":fc+ab,
        "fat_re":r2(day_fat_re[d]),"fat_psi":r2(day_fat_psi[d])})

# ── STEP 10: Campaigns ────────────────────────────────────────────
CAMP_META={
    "[PSI08] [Compra] [Frio] - ABO":                        {"prod":"PSI","aud":"Frio","creat":"Misto","obj":"sales","status":"PAUSED"},
    "[PSI08] [Compra] [Quente] [Teste Criativos] [Estáticos] - ABO":{"prod":"PSI","aud":"Quente","creat":"Estáticos","obj":"sales","status":"PAUSED"},
    "[PSI08] [Initiate Checkout] [Frio] [ADV+] - ABO":      {"prod":"PSI","aud":"Frio","creat":"ADV+","obj":"sales","status":"PAUSED"},
    "[RE] [Compra] [Quente] [Validação] [Estáticos] - ABO": {"prod":"RE","aud":"Quente","creat":"Estáticos","obj":"sales","status":"PAUSED"},
    "[RE] [Compra] [Quente] [Validação] [Vídeos] - ABO":    {"prod":"RE","aud":"Quente","creat":"Vídeos","obj":"sales","status":"ACTIVE"},
    "[RE] [Compra] [Frio] [Validação] [Vídeos] - ABO":      {"prod":"RE","aud":"Frio","creat":"Vídeos","obj":"sales","status":"ACTIVE"},
    "[RE] [Compra] [Frio] [Validação] [Estáticos] - ABO":   {"prod":"RE","aud":"Frio","creat":"Estáticos","obj":"sales","status":"ACTIVE"},
}
campaigns_arr=[]
fid=120243000000000001
for cname,mc in CAMP_META.items():
    if cname not in meta_by_camp: continue
    cd=meta_by_camp[cname]; sp=cd['spend']; imp=cd['imp']; reach=cd.get('reach',0)
    days_active=sorted(set(e['day'] for e in cd['daily']))
    ch=camp_hubla.get(cname,{'faturas':0,'fat':0.0,'nh':0.0})
    cc=ch['faturas']; cf=r2(ch['fat']); cnh=r2(ch['nh'])
    freq_v = r2(imp/reach) if reach else r2(RE_FREQ.get(cname, PSI_FREQ.get(cname, 1.4)))
    campaigns_arr.append({
        "id":str(fid),"name":cname,"prod":mc['prod'],"aud":mc['aud'],
        "creat":mc['creat'],"obj":mc['obj'],"status":mc['status'],
        "spend":r2(sp),"impressions":r0(imp),"reach":r0(reach),
        "cpm":sdiv(sp*1000,imp),"freq":freq_v,
        "days":days_active,"faturas":cc,"units":cc,"fat":cf,"nh":cnh,
        "cpa":r2(sdiv(sp,cc)) if cc else None,"roas":sdiv(cf,sp),"lucro":r2(cnh-sp),
        "daily":sorted(cd['daily'],key=lambda x:x['day'])})
    fid+=1

icd=meta_by_camp[INSTA_CAMP]
campaigns_arr.append({
    "id":str(fid),"name":INSTA_CAMP,"prod":"INSTA","aud":"Frio","creat":"Misto","obj":"followers","status":"ACTIVE",
    "spend":r2(icd['spend']),"impressions":r0(icd['imp']),"reach":r0(icd['reach']),
    "cpm":r2(sdiv(icd['spend']*1000,icd['imp'])),"freq":r2(sdiv(icd['imp'],icd['reach'])),
    "days":sorted(set(e['day'] for e in icd['daily'])),"faturas":0,"units":0,"fat":0,"nh":0,
    "cpa":None,"roas":0,"lucro":r2(-icd['spend']),
    "daily":sorted(icd['daily'],key=lambda x:x['day'])})

print("\nCampanhas (atribuição Hubla):")
for c4 in campaigns_arr:
    if c4['obj']!='followers':
        print(f"  {c4['name'][:55]}: {c4['faturas']} vendas R${c4['fat']:.0f} ROAS={c4['roas']} spend={c4['spend']:.0f}")

# ── STEP 11: Enriquecer daily_arr ────────────────────────────────
REG_ORDER = ['Sudeste','Sul','Nordeste','Centro-Oeste','Norte','Não informado']
for _entry in daily_arr:
    _d = _entry['day']
    _entry['spend_re']   = r2(sum(r['spend'] for r in meta_raw if r['campaign'] in RE_CAMPS  and int(r['date'][8:10])==_d))
    _entry['spend_psi']  = r2(sum(r['spend'] for r in meta_raw if r['campaign'] in PSI_CAMPS and int(r['date'][8:10])==_d))
    _entry['faturas_re'] = day_fat_re_c[_d]
    _entry['faturas_psi']= day_fat_psi_c[_d]
    # Funil por produto
    _entry['funnel_clicks_re']  = day_funnel_re[_d]['clicks']
    _entry['funnel_lpv_re']     = day_funnel_re[_d]['lpv']
    _entry['funnel_imp_re']     = day_funnel_re[_d]['imp']
    _entry['funnel_clicks_psi'] = day_funnel_psi[_d]['clicks']
    _entry['funnel_lpv_psi']    = day_funnel_psi[_d]['lpv']
    _entry['funnel_imp_psi']    = day_funnel_psi[_d]['imp']
    # Origem (canal de venda)
    _entry['origins'] = [{"name":ORIG_LABELS.get(k,k),"fat":r2(v['fat']),"faturas":v['faturas']}
                         for k,v in day_orig_d[_d].items() if v['fat']>0 or v['faturas']>0]
    _entry['fb_split'] = {k:{"fat":r2(v['fat']),"faturas":v['faturas']} for k,v in day_fbsp_d[_d].items()}
    # Região
    _entry['regioes'] = [{"name":k,"fat":r2(day_reg_d[_d].get(k,{}).get('fat',0)),
                          "faturas":day_reg_d[_d].get(k,{}).get('faturas',0)}
                         for k in REG_ORDER if day_reg_d[_d].get(k,{}).get('fat',0)>0]
    # Pagamento e parcelamento
    _entry['pay_dist']  = [{"method":k,"count":v} for k,v in sorted(day_pay_d[_d].items(),key=lambda x:-x[1]) if v>0]
    _entry['parc_dist'] = [{"n":k,"count":v} for k,v in sorted(day_parc_d[_d].items()) if v>0]
    # OBs por produto
    _entry['ob_detail_re']  = [{"name":k,"count":v['count'],"val":r2(v['val'])}
                               for k,v in sorted(day_ob_re_d[_d].items(),key=lambda x:-x[1]['count']) if v['count']>0]
    _entry['ob_detail_psi'] = [{"name":k,"count":v['count'],"val":r2(v['val'])}
                               for k,v in sorted(day_ob_psi_d[_d].items(),key=lambda x:-x[1]['count']) if v['count']>0]
    # Tag para o dashboard saber que este é mês corrente (vs histórico)
    _entry['mes_corrente'] = True

# ── STEP 11b: Prepender dias do histórico (meses anteriores) ────
# Para cada fatura de meses != mês corrente, agrega por (date) e adiciona ao daily_arr
# Os campos detalhados (origins, pay_dist, etc.) ficam vazios — só fat/nh/faturas estão
hist_invoices = [i for i in invoices_hist if not (i['year']==PERIOD_YEAR and i['month']==PERIOD_MONTH)]
hist_by_date = defaultdict(lambda: {
    'fat':0.0,'nh':0.0,'total':0.0,'faturas':0,'units':0,
    'fat_re':0.0,'fat_psi':0.0,'faturas_re':0,'faturas_psi':0,
    'origins':defaultdict(lambda: {'fat':0.0,'faturas':0}),
    'fb_split':{'frio':{'fat':0.0,'faturas':0},'quente':{'fat':0.0,'faturas':0},'outros':{'fat':0.0,'faturas':0}},
    'regioes':defaultdict(lambda: {'fat':0.0,'faturas':0}),
    'pay_dist':defaultdict(int),
    'parc_dist':defaultdict(int),
})
for inv in hist_invoices:
    p,ob,tot,nh = inv['prod'],inv['ob'],inv['total'],inv['nh']
    _hd = hist_by_date[inv['date']]
    _hd['fat']+=tot; _hd['nh']+=nh; _hd['total']+=tot
    _hd['faturas']+=1; _hd['units']+=inv['items']
    _is_re  = is_re_prod(p)  or is_re_prod(ob)
    _is_psi = is_psi_prod(p) or is_psi_prod(ob)
    if _is_re:
        _hd['fat_re']+=tot; _hd['faturas_re']+=1
    if _is_psi:
        _hd['fat_psi']+=tot; _hd['faturas_psi']+=1
    # Origem
    src = inv['src']
    if 'facebook' in src: orig='facebook ads'
    elif 'instagram' in src or 'bio' in src: orig='instagram'
    elif 'whatsapp' in src: orig='whatsapp'
    else: orig='sem origem'
    _hd['origins'][orig]['fat']+=tot; _hd['origins'][orig]['faturas']+=1
    if orig=='facebook ads':
        cl = inv['camp'].lower() if inv['camp'] else ''
        fk = 'frio' if 'frio' in cl else 'quente' if 'quente' in cl else 'outros'
        _hd['fb_split'][fk]['fat']+=tot; _hd['fb_split'][fk]['faturas']+=1
    # Região
    uf=inv['estado']
    reg=REGION_MAP.get(uf,'Não informado') if uf else 'Não informado'
    _hd['regioes'][reg]['fat']+=tot; _hd['regioes'][reg]['faturas']+=1
    # Pagamento
    pm = inv['pay_method'] or 'Sem informação'
    _hd['pay_dist'][pm]+=1
    if pm=='Cartão de Crédito':
        _hd['parc_dist'][inv['num_parcelas']]+=1

# Constrói entradas do histórico e PREPENDE ao daily_arr
hist_entries = []
for date_str, h in sorted(hist_by_date.items()):
    y,m,dd = [int(x) for x in date_str.split('-')]
    hist_entries.append({
        "day":dd, "month":m, "year":y, "date":date_str,
        "spend":0, "spend_sales":0, "impressions":0,
        "faturas":h['faturas'], "units":h['units'],
        "fat":r2(h['fat']), "nh":r2(h['nh']),
        "abandoned":0, "checkouts":h['faturas'],
        "fat_re":r2(h['fat_re']), "fat_psi":r2(h['fat_psi']),
        "spend_re":0, "spend_psi":0,
        "faturas_re":h['faturas_re'], "faturas_psi":h['faturas_psi'],
        "funnel_clicks_re":0,"funnel_lpv_re":0,"funnel_imp_re":0,
        "funnel_clicks_psi":0,"funnel_lpv_psi":0,"funnel_imp_psi":0,
        "origins":[{"name":ORIG_LABELS.get(k,k),"fat":r2(v['fat']),"faturas":v['faturas']}
                   for k,v in h['origins'].items() if v['fat']>0],
        "fb_split":{k:{"fat":r2(v['fat']),"faturas":v['faturas']} for k,v in h['fb_split'].items()},
        "regioes":[{"name":k,"fat":r2(h['regioes'][k]['fat']),"faturas":h['regioes'][k]['faturas']}
                   for k in REG_ORDER if h['regioes'].get(k,{}).get('fat',0)>0],
        "pay_dist":[{"method":k,"count":v} for k,v in sorted(h['pay_dist'].items(),key=lambda x:-x[1]) if v>0],
        "parc_dist":[{"n":k,"count":v} for k,v in sorted(h['parc_dist'].items()) if v>0],
        "ob_detail_re":[], "ob_detail_psi":[],
        "mes_corrente":False,
    })

# Adiciona month/year aos entries do mês corrente também
for _entry in daily_arr:
    _entry['month'] = PERIOD_MONTH
    _entry['year']  = PERIOD_YEAR

# Prepende histórico ao daily_arr e reordena por data
daily_arr = hist_entries + daily_arr
daily_arr.sort(key=lambda x: x['date'])
print(f"[Daily] {len(daily_arr)} dias no array (histórico + mês corrente)")

def prod_week_rows(day_fat_d, day_c_d, camp_set, nh_ratio=0.94):
    rows=[]
    for wid,label,d0,d1 in [("S1",f"01–07 {MES_ABBR}",1,7),
                            ("S2",f"08–14 {MES_ABBR}",8,14),
                            ("S3",f"15–21 {MES_ABBR}",15,21),
                            ("S4",f"22–28 {MES_ABBR}",22,28),
                            ("S5",f"29–{DAYS_MONTH} {MES_ABBR}",29,DAYS_MONTH)]:
        wf=sum(day_fat_d[d] for d in range(d0,d1+1))
        wc=sum(day_c_d[d]   for d in range(d0,d1+1))
        ws=sum(r['spend'] for r in meta_raw if r['campaign'] in camp_set and d0<=int(r['date'][8:10])<=d1)
        w_nh=r2(wf*nh_ratio)
        w_lucro=r2(w_nh-ws)
        w_imp=r0(sum(r['impressions'] for r in meta_raw if r['campaign'] in camp_set and d0<=int(r['date'][8:10])<=d1))
        rows.append({"id":wid,"label":label,"d0":d0,"d1":d1,"spend":r2(ws),"fat":r2(wf),
                     "faturas":wc,"nh":w_nh,"lucro":w_lucro,"impressions":w_imp,
                     "has_data":wf>0 or ws>0})
    return rows

re_benchmarks={
    "ticket":147,"ctr":{"bom":1.5,"media":1.0,"gargalo":0.5},
    "connect":{"bom":90,"media":80,"gargalo":70},"initiate":{"bom":30,"media":20,"gargalo":10},
    "pagamento":{"bom":15,"media":10,"gargalo":5},"conversao":{"bom":4.0,"media":2.0,"gargalo":1.0}
}
psi_benchmarks={
    "ticket":297,"ctr":{"bom":1.5,"media":1.0,"gargalo":0.5},
    "connect":{"bom":90,"media":80,"gargalo":70},"initiate":{"bom":20,"media":12,"gargalo":5},
    "pagamento":{"bom":10,"media":6,"gargalo":3},"conversao":{"bom":2.5,"media":1.2,"gargalo":0.6}
}

# ── Per-product OB detail arrays ──────────────────────────────────
re_ob_detail_arr  = sorted([{"name":k,"count":v['count'],"val":r2(v['val'])} for k,v in re_ob_names_detail.items()], key=lambda x:-x['val'])
psi_ob_detail_arr = sorted([{"name":k,"count":v['count'],"val":r2(v['val'])} for k,v in psi_ob_names_detail.items()], key=lambda x:-x['val'])

# ── Per-product payment/installment arrays ────────────────────────
def mk_pay_dist(counts, total):
    if not total: return []
    return [{"method":k,"count":v,"pct":r2(v/total*100)} for k,v in sorted(counts.items(),key=lambda x:-x[1])]

def mk_parc_arr(dist, cc_total):
    if not cc_total: return []
    return [{"n":k,"count":v,"pct":r2(v/cc_total*100)} for k,v in sorted(dist.items())]

re_pay_main_total  = sum(re_pay_counts.values())  or 1
psi_pay_main_total = sum(psi_pay_counts.values()) or 1
re_pay_dist   = mk_pay_dist(re_pay_counts,  re_pay_main_total)
psi_pay_dist  = mk_pay_dist(psi_pay_counts, psi_pay_main_total)
re_parc_arr   = mk_parc_arr(re_parc_dist,   re_pay_counts.get('Cartão de Crédito', 0))
psi_parc_arr  = mk_parc_arr(psi_parc_dist,  psi_pay_counts.get('Cartão de Crédito', 0))

products_paid={
    "RE":{"code":"RE","spend":r2(re_sp),"impressions":r0(re_imp_m),"reach":r0(re_reach_m),
          "faturas":re_count,"units":re_units,"fat":r2(re_fat),"nh":r2(re_nh),"total":r2(re_total),"ob":re_ob,
          "cpa":re_cpa_v,"roas":re_roas_v,"roas_nh":re_roas_nh,"lucro":re_lucro,"roi":re_roi,
          "ticket":re_ticket,"cpm":r2(re_cpm_m),"abandoned":ab_re,"checkouts":re_checkouts,"compras":re_count,
          "funnel_clicks":re_clicks,"funnel_lpv":re_lpv,"funnel_ctr":re_ctr,"funnel_cpc":re_cpc,
          "funnel_source":"planilha","conv_checkout":sdiv(re_count*100,re_checkouts),
          "weeks":prod_week_rows(day_fat_re,day_fat_re_c,RE_CAMPS,nh_ratio=sdiv(re_nh,re_fat,0.94)),
          "benchmarks":re_benchmarks,
          "pay_dist":re_pay_dist,"parc_dist":re_parc_arr,
          "ob_val":r2(re_ob_val_total),"ob_detail":re_ob_detail_arr},
    "PSI":{"code":"PSI","spend":r2(psi_sp),"impressions":r0(psi_imp_m),"reach":r0(psi_reach_m),
           "faturas":psi_count,"units":psi_units,"fat":r2(psi_fat),"nh":r2(psi_nh),"total":r2(psi_total),"ob":psi_ob,
           "cpa":psi_cpa_v,"roas":psi_roas_v,"roas_nh":psi_roas_nh,"lucro":psi_lucro,"roi":psi_roi,
           "ticket":psi_ticket,"cpm":r2(psi_cpm_m),"abandoned":ab_psi,"checkouts":psi_checkouts,"compras":psi_count,
           "funnel_clicks":psi_clicks,"funnel_lpv":psi_lpv,"funnel_ctr":psi_ctr,"funnel_cpc":psi_cpc,
           "funnel_source":"planilha","conv_checkout":sdiv(psi_count*100,psi_checkouts),
           "weeks":prod_week_rows(day_fat_psi,day_fat_psi_c,PSI_CAMPS,nh_ratio=sdiv(psi_nh,psi_fat,0.94)),
           "benchmarks":psi_benchmarks,
           "pay_dist":psi_pay_dist,"parc_dist":psi_parc_arr,
           "ob_val":r2(psi_ob_val_total),"ob_detail":psi_ob_detail_arr}
}

# ── STEP 12: Origins, products, regions ──────────────────────────
origins_arr=[
    {"name":"Facebook Ads","faturas":origins_map['facebook ads']['faturas'],
     "fat":r2(origins_map['facebook ads']['fat']),"nh":r2(origins_map['facebook ads']['nh'])},
    {"name":"Instagram (orgânico)","faturas":origins_map['instagram']['faturas'],
     "fat":r2(origins_map['instagram']['fat']),"nh":r2(origins_map['instagram']['nh'])},
    {"name":"WhatsApp","faturas":origins_map['whatsapp']['faturas'],
     "fat":r2(origins_map['whatsapp']['fat']),"nh":r2(origins_map['whatsapp']['nh'])},
    {"name":"Hotmart","faturas":origins_map['hotmart']['faturas'],
     "fat":r2(origins_map['hotmart']['fat']),"nh":r2(origins_map['hotmart']['nh'])},
    {"name":"Sem origem","faturas":origins_map['sem origem']['faturas'],
     "fat":r2(origins_map['sem origem']['fat']),"nh":r2(origins_map['sem origem']['nh'])}
]
products_arr=[{"name":p,"faturas":prod_fat_c[p],"fat":r2(prod_fat[p]),"units":prod_units_d[p],
               "recorrente":('Turma 09' in p),
               "plataforma": "hotmart" if p in hotmart_by_prod and p not in [k for k in prod_fat if prod_fat[k]>0 and k not in hotmart_by_prod] else "hubla"}
              for p in sorted(prod_fat,key=lambda x:-prod_fat[x])]

reg_order=['Sudeste','Sul','Nordeste','Centro-Oeste','Norte','Não informado']
regioes_arr=[{"name":r,"faturas":reg_count[r],"fat":r2(reg_fat[r])} for r in reg_order if reg_count[r]>0]
estados_arr=[{"uf":s,"faturas":state_count[s],"fat":r2(state_fat[s])}
             for s in sorted(state_fat,key=lambda x:-state_fat[x])]

# ── STEP 13: Followers ─────────────────────────────────────────
icd2=meta_by_camp[INSTA_CAMP]
followers_obj={
    "id":"120241138795630031","name":INSTA_CAMP,
    "spend":r2(icd2['spend']),"impressions":r0(icd2['imp']),"reach":r0(icd2['reach']),
    "cpm":r2(sdiv(icd2['spend']*1000,icd2['imp'])),"freq":r2(sdiv(icd2['imp'],icd2['reach'])),
    "days":DAYS_ELAPSED,"pct_invest":r2(icd2['spend']/total_meta_spend*100),
    "cost_per_1k_reach":r2(icd2['spend']/icd2['reach']*1000) if icd2['reach'] else 0,
    "daily":sorted(icd2['daily'],key=lambda x:x['day']),
    "weeks":[{"id":w["id"],"label":w["label"],
              "spend":r2(sum(r['spend'] for r in meta_raw if r['campaign']==INSTA_CAMP and w['d0']<=int(r['date'][8:10])<=w['d1'])),
              "impressions":r0(sum(r['impressions'] for r in meta_raw if r['campaign']==INSTA_CAMP and w['d0']<=int(r['date'][8:10])<=w['d1'])),
              "has_data":any(r['campaign']==INSTA_CAMP and w['d0']<=int(r['date'][8:10])<=w['d1'] for r in meta_raw)}
             for w in weeks]
}

# ── STEP 14: Payment & installments ──────────────────────────────
payment_dist=[{"method":k,"count":v,"pct":r2(v/total_faturas*100)} for k,v in sorted(pay_counts.items(),key=lambda x:-x[1])]
parc_arr=[{"n":k,"count":v,"pct":r2(v/pay_counts.get('Cartão de Crédito',1)*100)} for k,v in sorted(parc_dist.items())]

# ── STEP 15: OB detail ────────────────────────────────────────────
ob_detail_arr=sorted([{"name":k,"count":v['count'],"val":r2(v['val'])} for k,v in ob_names_detail.items()],key=lambda x:-x['val'])

# ── STEP 16: Insights ─────────────────────────────────────────────
def camp_label(name):
    parts = re.findall(r'\[([^\]]+)\]', name)
    if len(parts) >= 2:
        prod_tag = parts[0]
        attrs = [p for p in parts[1:] if p.lower() not in ('compra','validação','teste criativos','initiate checkout','abo')]
        return f"{prod_tag} · {' · '.join(attrs)}" if attrs else f"{prod_tag} · {parts[1]}"
    return name.split(' - ')[0].strip()

w4 = next((w for w in weeks if w['id']=='S4'), None)
w3 = next((w for w in weeks if w['id']=='S3'), None)
w2 = next((w for w in weeks if w['id']=='S2'), None)

# Dias com dados em S4 (dinâmico)
s4_end_day   = min(DAYS_ELAPSED, 28)
s4_days      = max(0, s4_end_day - 21)
s4_range_str = f"22–{s4_end_day}/{MES_ABBR}" if s4_days > 0 else f"22–28/{MES_ABBR}"

re_cpm_atual = r2(sdiv(re_sp * 1000, re_imp_m))

best_camp = max((c5 for c5 in campaigns_arr if c5['status']=='ACTIVE' and c5['obj']=='sales' and c5['faturas']>0),
                key=lambda x: x['roas'], default=None)
worst_freq_camp = max((c5 for c5 in campaigns_arr if c5['status']=='ACTIVE' and c5['obj']=='sales' and c5['spend']>500),
                      key=lambda x: x['freq'], default=None)

insights = []

# 1. Tendência S3→S4
if w3 and w4 and w3['faturas'] > 0 and w4['fat'] > 0:
    if w4['roas'] >= 5.0:
        insights.append({
            "level":"positivo","icon":"star","tag":"MELHOR SEMANA",
            "title":f"Semana 4 ({s4_range_str}): ROAS {w4['roas']:.1f}x e CPA R${w4['cpa']:.0f} — performance crescente",
            "body": f"Os {s4_days} dias de {s4_range_str} já entregaram R${w4['fat']:,.0f} com R${w4['spend_sales']:,.0f} investidos (ROAS {w4['roas']:.1f}x). CPA caiu de R${w3['cpa']:.0f} (S3) para R${w4['cpa']:.0f} — eficiência crescente. Sinal para aumentar orçamento gradualmente: +15–20% a cada 2–3 dias enquanto ROAS ≥ 4x."
        })
    else:
        insights.append({
            "level":"positivo","icon":"trending-up","tag":"BOM RITMO",
            "title":f"Semana 4 ({s4_range_str}): ROAS {w4['roas']:.1f}x e {w4['faturas']} vendas em {s4_days} dias",
            "body": f"Dias {s4_range_str}: R${w4['fat']:,.0f} faturados com R${w4['spend_sales']:,.0f} investidos. CPA R${w4['cpa']:.0f}. Se mantiver o ritmo, semana completa estimada em ~R${int(w4['fat']/max(s4_days,1)*7):,}."
        })

# 2. Melhor campanha ativa
if best_camp and best_camp['roas'] >= 2.0:
    lbl = camp_label(best_camp['name'])
    freq_alert = " Frequência ainda controlada." if best_camp['freq'] < 1.6 else f" Atenção: freq {best_camp['freq']:.2f} — monitorar saturação."
    insights.append({
        "level":"positivo","icon":"star","tag":"ESCALAR AGORA",
        "title":f"{lbl}: ROAS {best_camp['roas']:.2f}x — candidata ao escalonamento",
        "body": f"R${best_camp['spend']:,.0f} investidos geraram R${best_camp['fat']:,.0f} ({best_camp['faturas']} vendas) com ROAS {best_camp['roas']:.2f}x.{freq_alert} Aumentar 15-20% do orçamento a cada 2-3 dias enquanto ROAS ≥ 3x."
    })

# 3. Saturação de frequência
if worst_freq_camp and worst_freq_camp.get('freq',0) >= 1.6:
    c5 = worst_freq_camp
    insights.append({
        "level":"atencao","icon":"image","tag":"SATURAÇÃO",
        "title":f"Freq {c5['freq']:.2f} em '{camp_label(c5['name'])}' — renovar criativos",
        "body": f"Com R${c5['spend']:,.0f} investidos e freq {c5['freq']:.2f} (limiar ≥ 1.5), CTR tende a cair e CPM a subir. Preparar 3-5 novos criativos e substituir antes que a performance caia."
    })

# 4. RE vs PSI comparativo
if re_count > 0 and psi_count > 0:
    if re_roas_v > psi_roas_v:
        insights.append({
            "level":"info","icon":"funnel","tag":"PORTFÓLIO",
            "title":f"RE lidera: ROAS {re_roas_v:.2f}x vs PSI {psi_roas_v:.2f}x — priorizar alocação no RE",
            "body": f"Regulando as Emoções gerou R${re_fat:,.0f} ({re_count} vendas, ticket R${re_ticket:.0f}) vs PSI08 R${psi_fat:,.0f} ({psi_count} vendas, ticket R${psi_ticket:.0f}). Dado o ROAS superior do RE, concentrar escala no RE enquanto testa novas campanhas PSI."
        })
    else:
        insights.append({
            "level":"info","icon":"funnel","tag":"PORTFÓLIO",
            "title":f"PSI lidera com ROAS {psi_roas_v:.2f}x — avaliar reabertura de campanhas",
            "body": f"PSI08: R${psi_fat:,.0f} em {psi_count} vendas (ticket R${psi_ticket:.0f}). RE: R${re_fat:,.0f} em {re_count} vendas (ticket R${re_ticket:.0f}). Considerar reabertura de campanhas PSI com novos criativos."
        })

# 5. PSI parado com ROAS comprovado
if psi_roas_v >= 2.0 and psi_count > 0:
    any_psi_active = any(c5['status']=='ACTIVE' and c5['prod']=='PSI' for c5 in campaigns_arr)
    if not any_psi_active:
        insights.append({
            "level":"atencao","icon":"dollar","tag":"PSI — REATIVAR",
            "title":f"PSI parado com ROAS {psi_roas_v:.2f}x comprovado — reativar para ampliar faturamento",
            "body": f"O PSI08 gerou R${psi_fat:,.0f} com ROAS {psi_roas_v:.2f}x e ticket médio de R${psi_ticket:.0f}. Reativar com R$30-50/dia, novos criativos e verificação de UTMs."
        })

# 6. OB performance
if ob_count > 0 and ob_val_total > 0:
    ob_rate = sdiv(ob_count * 100, total_faturas)
    ob_fat_pct = sdiv(ob_val_total * 100, total_fat)
    insights.append({
        "level":"positivo","icon":"plus","tag":"ORDER BUMP",
        "title":f"Order Bumps adicionam +{ob_fat_pct:.0f}% ao faturamento — {ob_count} compras complementares",
        "body": f"{ob_count} order bumps aceitos em {total_faturas} faturas ({ob_rate:.1f}% de adesão), adicionando R${ob_val_total:,.0f} ao faturamento."
    })

# 7. Carrinho abandonado
if total_ab > 0:
    conv_cart = sdiv(total_faturas * 100, total_checkouts)
    cart_loss = r2(total_ab * ticket_med)
    insights.append({
        "level":"atencao","icon":"cart","tag":"RECUPERAÇÃO",
        "title":f"{total_ab} carrinhos abandonados — R${cart_loss:,.0f} em receita potencial",
        "body": f"Taxa de conversão do checkout: {conv_cart:.1f}% ({total_faturas} compras / {total_checkouts} checkouts). Sequência de WhatsApp + e-mail em 1h-24h pode recuperar 15-30% dos leads."
    })

# 8. CPM
if re_cpm_atual > 25:
    insights.append({
        "level":"atencao","icon":"trending-down","tag":"CPM ALTO",
        "title":f"CPM médio do RE em R${re_cpm_atual:.0f} — eficiência de entrega abaixo do ideal",
        "body": f"CPM acima de R$25 indica competição no leilão ou fadiga criativa. Testar criativos novos, revisar públicos."
    })
elif re_cpm_atual < 20:
    insights.append({
        "level":"positivo","icon":"cash","tag":"CPM EFICIENTE",
        "title":f"CPM médio do RE em R${re_cpm_atual:.0f} — custo de entrega eficiente",
        "body": f"CPM abaixo de R$20 indica boa relevância. Aproveitar para escalar 15-20%/semana."
    })

# 9. Margem
margem = sdiv(total_nh * 100, total_fat)
insights.append({
    "level":"info","icon":"profit","tag":"MARGEM",
    "title":f"Margem líquida (após Hubla): {margem:.1f}% — R${total_nh:,.0f} de {total_faturas} vendas",
    "body": f"De R${total_fat:,.0f} faturados, R${total_nh:,.0f} chega líquido (margem {margem:.1f}%). Com R${total_sales_spend:,.0f} em mídia, o lucro operacional é R${lucro_tot:,.0f} (ROI {roi_tot:.0f}%). Para cada R$1 investido, retornam R${roas_sales:.2f}."
})

print(f"\n[Insights] {len(insights)} gerados")
for ins in insights:
    print(f"  [{ins['level'].upper():8s}] {ins['title'][:70]}")

# ── STEP 16b: Anúncios (ad-level CSV — opcional) ─────────────────
# Detecta qualquer CSV em BASE_DATA que contenha a coluna 'Ad name'
import glob as _glob, os as _os
ads_arr = []
_ad_files = sorted(_glob.glob(_os.path.join(BASE_DATA, '*.csv')))
_AD_COLS_EN = {'ad name','ad_name'}
_AD_FILE = None
for _af in _ad_files:
    try:
        with open(_af, encoding='utf-8') as _f:
            _hdr = [h.strip().lower() for h in _f.readline().split(',')]
        if any(h in _AD_COLS_EN for h in _hdr):
            _AD_FILE = _af; break
    except: pass

if _AD_FILE:
    print(f"[Anúncios] Arquivo detectado: {_os.path.basename(_AD_FILE)}")
    _ads_map = defaultdict(lambda: {'spend':0.0,'impressions':0,'clicks':0,'lpv':0,
                                     'reach':0,'cpm_x_imp':0.0,'n_rows':0})
    with open(_AD_FILE, encoding='utf-8') as _f:
        _rd = csv.DictReader(_f)
        for _row in _rd:
            _dt2 = parse_day_col((_row.get('Day') or _row.get('day') or ''))
            if _dt2 and not (PERIOD_START <= _dt2.replace(hour=0,minute=0,second=0) <= PERIOD_END): continue
            _ad_name  = (_row.get('Ad name') or _row.get('Ad Name') or _row.get('ad name') or '').strip()
            _camp_name = (_row.get('Campaign name') or _row.get('Campaign Name') or '').strip()
            if not _ad_name: continue
            _key = (_camp_name, _ad_name)
            _sp  = n(_row.get('Amount Spent','0'))
            _imp = r0(n(_row.get('Impressions','0')))
            _cl  = r0(n(_row.get('Link Clicks','0')))
            _lp  = r0(n(_row.get('Landing Page Views','0')))
            _rc  = r0(n(_row.get('Reach','0')))
            _cpm = n(_row.get('CPM (Cost per 1,000 Impressions)','0'))
            _ads_map[_key]['spend']       += _sp
            _ads_map[_key]['impressions'] += _imp
            _ads_map[_key]['clicks']      += _cl
            _ads_map[_key]['lpv']         += _lp
            _ads_map[_key]['reach']       += _rc
            _ads_map[_key]['cpm_x_imp']   += _cpm * _imp
            _ads_map[_key]['n_rows']      += 1
    # Índice de vendas Hubla por campanha
    _camp_hubla = {}
    for _c in campaigns_arr:
        _cn = _c['name']
        _camp_hubla[_cn] = {
            'faturas': _c.get('faturas', 0),
            'fat':     _c.get('fat', 0.0),
            'nh':      _c.get('nh', 0.0),
            'spend':   _c.get('spend', 0.0),
        }
    # Total de link clicks por campanha
    _camp_clicks_tot = defaultdict(int)
    for (_c, _), _v2 in _ads_map.items():
        _camp_clicks_tot[_c] += _v2['clicks']
    for (_camp, _aname), _v in sorted(_ads_map.items(), key=lambda x: -x[1]['spend']):
        _sp  = r2(_v['spend']); _imp = _v['impressions']
        _cl  = _v['clicks'];    _lp  = _v['lpv']; _rc = _v['reach']
        _cpm = r2(_v['cpm_x_imp']/_imp) if _imp else 0
        _freq = r2(_imp/_rc) if _rc else 0
        _ctr   = r2(_cl/_imp*100) if _imp else 0
        _cpc   = r2(_sp/_cl)      if _cl  else 0
        _lpv_r = r2(_lp/_cl*100) if _cl  else 0
        _ch = _camp_hubla.get(_camp, {})
        _c_cl_tot = _camp_clicks_tot[_camp]
        _prop_cl = _cl / _c_cl_tot if _c_cl_tot > 0 else 0.0
        _est_fat     = r2(_ch.get('fat', 0.0) * _prop_cl)
        _est_faturas = int(round(_ch.get('faturas', 0) * _prop_cl))
        _est_roas    = r2(_est_fat / _sp) if _sp > 0 else 0
        ads_arr.append({
            "campaign":    _camp,  "name": _aname,
            "spend":       _sp,    "impressions": _imp,
            "reach":       _rc,    "freq": _freq,
            "cpm":         _cpm,   "clicks": _cl,
            "ctr":         _ctr,   "cpc": _cpc,
            "lpv":         _lp,    "lpv_rate": _lpv_r,
            "fat_est":     _est_fat,
            "faturas_est": _est_faturas,
            "roas_est":    _est_roas,
        })
    # ── Mesclar cópias com o anúncio pai ─────────────────────────────
    def _base_name(name):
        return re.sub(r'\s*[—\-]{1,2}\s*C[oó]pia.*$', '', name.strip()).strip()

    _merged = defaultdict(lambda: {'spend':0.0,'impressions':0,'reach':0,
                                    'clicks':0,'lpv':0,'cpm_x_imp':0.0,
                                    'names':[]})
    for _ad in ads_arr:
        _bn  = _base_name(_ad['name'])
        _key = (_ad['campaign'], _bn)
        _m   = _merged[_key]
        _m['spend']      += _ad['spend']
        _m['impressions']+= _ad['impressions']
        _m['reach']      += _ad['reach']
        _m['clicks']     += _ad['clicks']
        _m['lpv']        += _ad['lpv']
        _m['cpm_x_imp']  += _ad['cpm'] * _ad['impressions']
        _m['names'].append(_ad['name'])

    _camp_cl_merged = defaultdict(int)
    for (_cc, _), _mv in _merged.items():
        _camp_cl_merged[_cc] += _mv['clicks']

    ads_arr = []
    for (_camp, _aname), _mv in sorted(_merged.items(), key=lambda x: -x[1]['spend']):
        _sp   = r2(_mv['spend']); _imp = _mv['impressions']
        _rc   = _mv['reach'];     _cl  = _mv['clicks']; _lp = _mv['lpv']
        _cpm  = r2(_mv['cpm_x_imp']/_imp) if _imp else 0
        _freq = r2(_imp/_rc)     if _rc else 0
        _ctr  = r2(_cl/_imp*100) if _imp else 0
        _cpc  = r2(_sp/_cl)      if _cl  else 0
        _lpvr = r2(_lp/_cl*100)  if _cl  else 0
        _ch   = _camp_hubla.get(_camp, {})
        _c_cl = _camp_cl_merged[_camp]
        _prop = _cl/_c_cl if _c_cl else 0.0
        _ef   = r2(_ch.get('fat',0.0)*_prop)
        _ec   = int(round(_ch.get('faturas',0)*_prop))
        _er   = r2(_ef/_sp) if _sp else 0
        _is_merged = len(_mv['names']) > 1
        ads_arr.append({
            "campaign":    _camp,  "name": _aname,
            "spend":       _sp,    "impressions": _imp,
            "reach":       _rc,    "freq": _freq,
            "cpm":         _cpm,   "clicks": _cl,
            "ctr":         _ctr,   "cpc": _cpc,
            "lpv":         _lp,    "lpv_rate": _lpvr,
            "fat_est":     _ef,    "faturas_est": _ec, "roas_est": _er,
            "merged":      _is_merged,
            "copies_count":len(_mv['names'])-1,
        })
    _n_copies = sum(1 for _ad in ads_arr if _ad['merged'])
    print(f"[Anúncios] {len(ads_arr)} anúncios únicos após mesclar cópias "
          f"({sum(a['copies_count'] for a in ads_arr)} cópias absorvidas)")
else:
    print("[Anúncios] Nenhum CSV com coluna 'Ad name' encontrado — ads: []")

# ── STEP 17: Assemble JSON ────────────────────────────────────────
ob_by_prod = {}
for p5 in products_arr:
    key = 'RE' if is_re_prod(p5['name']) else ('PSI' if is_psi_prod(p5['name']) else 'OUTROS')
    ob_by_prod.setdefault(key,[]).append(p5)

out = {
    "period": {
        "start": PERIOD_START.strftime('%Y-%m-%d'), "end": PERIOD_END.strftime('%Y-%m-%d'),
        "label": PERIOD_LABEL,
        "days_elapsed": DAYS_ELAPSED, "days_month": DAYS_MONTH
    },
    "_v": int(time.time()*1000),
    "totals": {
        "spend":         r2(total_meta_spend),
        "spend_sales":   r2(total_sales_spend),
        "spend_insta":   r2(total_insta_spend),
        "impressions":   r0(total_meta_imp),
        "impressions_sales": r0(total_sales_imp),
        "fat":           r2(total_fat),
        "total":         r2(total_total),
        "nh":            r2(total_nh),
        "hubla_fee":     r2(hubla_fee),
        "faturas":       total_faturas,
        "units":         total_units,
        "abandoned":     total_ab,
        "checkouts":     total_checkouts,
        "ob_count":      ob_count,
        "ob_val":        r2(ob_val_total),
        "fat_with_ob":   ob_pedidos,
        "cpa":           r2(cpa_tot),
        "roas":          r2(roas_geral),
        "roas_sales":    r2(roas_sales),
        "lucro":         r2(lucro_tot),
        "roi":           r2(roi_tot),
        "ticket":        r2(ticket_med),
        "conv_checkout": r2(conv_checkout),
    },
    "daily":         daily_arr,
    "weeks":         weeks,
    "products_paid": products_paid,
    "campaigns":     campaigns_arr,
    "ads":           ads_arr,
    "followers":     followers_obj,
    "origins":       origins_arr,
    "products":      products_arr,
    "regioes":       regioes_arr,
    "estados":       estados_arr,
    "payment_dist":  payment_dist,
    "parc_dist":     parc_arr,
    "ob_detail":     ob_detail_arr,
    "insights":      insights,
    "fb_split":      fb_split,
}

out_path = f'{BASE_OUT}/dados.json'
with open(out_path,'w',encoding='utf-8') as f:
    json.dump(out,f,ensure_ascii=False,indent=1)
print(f"\n✅ Salvo: {out_path} ({len(json.dumps(out)):,} bytes)")
