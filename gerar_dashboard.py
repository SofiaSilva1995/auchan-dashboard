#!/usr/bin/env python3
"""
Gerador Dashboard Auchan Sell-Out — Pierre Fabre Portugal
Uso: python3 gerar_auchan_dashboard.py [caminho_excel] [ficheiro_output]
"""

import pandas as pd
import json
import sys
import base64
import unicodedata
from pathlib import Path
from datetime import datetime

PF_LOGO_B64 = 'data:image/jpeg;base64,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'

# ─── CONFIG ──────────────────────────────────────────────────────────────────
import glob as _glob
_xlsx = sorted(_glob.glob('*.xlsx') + _glob.glob('*.xlsb') + _glob.glob('*.xls'))
_csv  = sorted(_glob.glob('*.csv'))
EXCEL = sys.argv[1] if len(sys.argv) > 1 else (_xlsx[0] if _xlsx else (_csv[0] if _csv else None))
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else f'Dashboard_Auchan_{datetime.now().strftime("%d_%m_%Y")}.html'

BRAND_COLORS = {
    'AVENE':       '#4a9eda',
    'DUCRAY':      '#2c7a4b',
    'A-DERMA':     '#e85d04',
    'FURTERER':    '#7c3aed',
    'KLORANE':     '#0891b2',
    'ORAL CARE':   '#0d9488',
    'HEALTH CARE': '#dc2626',
    'PFD':         '#b45309',
    'PFM':         '#4f46e5',
}
DEFAULT_BRAND_COLOR = '#64748b'

CAT_COLORS = {
    'SOLARES':       '#f59e0b',
    'ROSTO':         '#ec4899',
    'CORPO':         '#06b6d4',
    'CAPILARES':     '#8b5cf6',
    'BEBÉ':          '#10b981',
    'HIGIENE ORAL':  '#3b82f6',
    'ARTICULAÇÃO':   '#ef4444',
    'MÁ CIRCULAÇÃO': '#f97316',
    'REPELENTES':    '#84cc16',
    'LÁBIOS / MÃOS': '#e879f9',
    'DESINFETANTE':  '#6b7280',
}
DEFAULT_CAT_COLOR = '#94a3b8'

MONTH_PT = {1:'Jan',2:'Fev',3:'Mar',4:'Abr',5:'Mai',6:'Jun',
            7:'Jul',8:'Ago',9:'Set',10:'Out',11:'Nov',12:'Dez'}

# ─── 1. LOAD & CLEAN ─────────────────────────────────────────────────────────
if not EXCEL:
    raise FileNotFoundError("Nenhum ficheiro .xlsx/.xlsb/.csv encontrado no directório actual.")
print(f"A carregar {Path(EXCEL).name} ...")
_ext = Path(EXCEL).suffix.lower()
if _ext == '.csv':
    try:
        raw = pd.read_csv(EXCEL, encoding='utf-8', low_memory=False)
    except UnicodeDecodeError:
        raw = pd.read_csv(EXCEL, encoding='latin-1', low_memory=False)
elif _ext == '.xlsb':
    raw = pd.read_excel(EXCEL, sheet_name=0, header=0, engine='pyxlsb')
    raw.columns = raw.iloc[0].tolist()
    raw = raw.iloc[1:].reset_index(drop=True)
else:
    raw = pd.read_excel(EXCEL, sheet_name=0, header=0)
    raw.columns = raw.iloc[0].tolist()
    raw = raw.iloc[1:].reset_index(drop=True)

# Normalizar nomes de colunas (strip whitespace duplo)
raw.columns = [str(c).strip().replace('  ', ' ') for c in raw.columns]

df = raw.copy()
df['TPVP']    = pd.to_numeric(df.get('TPVP',    df.get('TPVP', 0)), errors='coerce').fillna(0)
df['TQT VND'] = pd.to_numeric(df.get('TQT VND', df.get('TQT VND', 0)), errors='coerce').fillna(0)
df['Mês']     = pd.to_numeric(df['Mês'], errors='coerce')
df['Ano']     = pd.to_numeric(df['Ano'], errors='coerce')

# Remover linhas sem vendas (zero-quantity rows)
df = df[(df['TQT VND'] > 0) & (df['TPVP'] > 0)].copy()
print(f"  {len(df):,} linhas válidas")

df['YM'] = df['Ano'].astype(int).astype(str) + '-' + df['Mês'].astype(int).apply(lambda x: f'{x:02d}')

# Colunas chave
CV   = 'TPVP'
CU   = 'TQT VND'
CMARCA  = 'Marca PF'
CCAT    = 'CATEGORIA2'
CLOJA   = 'Loja'
CCNP    = 'EAN/CNP'
CDESC   = 'Descrição PF'  # after normalize
CGAMA   = 'GAMA'

# Fallback para nome de coluna descrição
if CDESC not in df.columns:
    for c in df.columns:
        if 'descrição' in c.lower() or 'descricao' in c.lower():
            CDESC = c
            break

for col in [CMARCA, CCAT, CLOJA]:
    df[col] = df[col].fillna('OUTROS').astype(str).str.strip()
df[CLOJA] = df[CLOJA].apply(lambda s: unicodedata.normalize('NFD', str(s)).encode('ascii', errors='ignore').decode())
df[CCNP]  = df[CCNP].astype(str).str.strip().str.rstrip('.0')
df[CDESC] = df[CDESC].fillna('').astype(str).str.strip() if CDESC in df.columns else ''
if CGAMA in df.columns:
    df[CGAMA] = df[CGAMA].fillna('').astype(str).str.strip()
else:
    df[CGAMA] = ''

# ─── 2. DIMENSÕES ─────────────────────────────────────────────────────────────
months = sorted(df['YM'].unique().tolist())
brands = sorted(df[CMARCA].unique().tolist())
cats   = sorted(df[CCAT].unique().tolist())
stores = sorted(df[CLOJA].unique().tolist())

ym_2025 = [m for m in months if m.startswith('2025')]
ym_2026 = [m for m in months if m.startswith('2026')]

def ym_label(ym):
    y, m = ym.split('-')
    return f"{MONTH_PT[int(m)]} {y}"

month_labels = {ym: ym_label(ym) for ym in months}

print(f"  Marcas: {brands}")
print(f"  Categorias: {cats}")
print(f"  Lojas: {len(stores)}, Produtos: {df[CCNP].nunique()}, Meses: {months[0]}..{months[-1]}")

# ─── 3. AGREGAÇÕES ───────────────────────────────────────────────────────────
def agg(g):
    return {'v': round(float(g[CV].sum()), 2), 'u': int(g[CU].sum())}

print("A agregar dados...")

# monthly total
monthly = {}
for ym, g in df.groupby('YM'):
    monthly[ym] = agg(g)

# brand × month
bm = {}
for (b, ym), g in df.groupby([CMARCA, 'YM']):
    bm[f"{b}|{ym}"] = agg(g)

# store × month
sm = {}
for (s, ym), g in df.groupby([CLOJA, 'YM']):
    sm[f"{s}|{ym}"] = agg(g)

# cat × month
cm = {}
for (c, ym), g in df.groupby([CCAT, 'YM']):
    cm[f"{c}|{ym}"] = agg(g)

# brand × cat × month
bcm = {}
for (b, c, ym), g in df.groupby([CMARCA, CCAT, 'YM']):
    bcm[f"{b}|{c}|{ym}"] = agg(g)

# store × brand × month
sbm = {}
for (s, b, ym), g in df.groupby([CLOJA, CMARCA, 'YM']):
    sbm[f"{s}|{b}|{ym}"] = agg(g)

# store × cat × month
scm = {}
for (s, c, ym), g in df.groupby([CLOJA, CCAT, 'YM']):
    scm[f"{s}|{c}|{ym}"] = agg(g)

# products
print("A processar produtos...")
prod_meta = df.groupby(CCNP).agg(
    desc  =(CDESC,  'first'),
    marca =(CMARCA, 'first'),
    cat   =(CCAT,   'first'),
    gama  =(CGAMA,  'first'),
).reset_index().rename(columns={CCNP: 'cnp'})

cnp_months = {}
for (cnp, ym), g in df.groupby([CCNP, 'YM']):
    if cnp not in cnp_months:
        cnp_months[cnp] = {}
    cnp_months[cnp][ym] = agg(g)

products = []
for _, row in prod_meta.iterrows():
    cnp = str(row['cnp'])
    by_m = cnp_months.get(cnp, {})
    has_25 = any(m in by_m for m in ym_2025)
    has_26 = any(m in by_m for m in ym_2026)
    products.append({
        'cnp': cnp,
        'desc': str(row['desc']),
        'marca': str(row['marca']),
        'cat': str(row['cat']),
        'gama': str(row['gama']),
        'm': by_m,
        'innov': has_26 and not has_25
    })

# top CNPs per store (para filtro loja na tabela produtos)
store_top_cnps = {}
for store in stores:
    g = df[df[CLOJA] == store].groupby(CCNP)[CV].sum().nlargest(100)
    store_top_cnps[store] = g.index.tolist()

print(f"  {len(products)} produtos agregados")

# ─── 4. MONTAR JSON DATA ──────────────────────────────────────────────────────
period_options = []
# YTD 2026 (opção default)
last_2026_m = max(int(m.split('-')[1]) for m in ym_2026) if ym_2026 else 5
period_options.append({'val': 'ytd2026', 'label': f'YTD 2026 (Jan-{MONTH_PT[last_2026_m]})'})
# Meses 2026 desc
for ym in sorted(ym_2026, reverse=True):
    period_options.append({'val': ym, 'label': month_labels[ym]})
# Separador
period_options.append({'val': '__sep__', 'label': '── 2025 ──'})
# Meses 2025 desc
for ym in sorted(ym_2025, reverse=True):
    period_options.append({'val': ym, 'label': month_labels[ym]})
# 2025 completo
period_options.append({'val': 'all2025', 'label': '2025 Completo'})

DATA = {
    'months': months,
    'ym_2025': ym_2025,
    'ym_2026': ym_2026,
    'brands': brands,
    'cats': cats,
    'stores': stores,
    'month_labels': month_labels,
    'period_options': period_options,
    'brand_colors': {b: BRAND_COLORS.get(b, DEFAULT_BRAND_COLOR) for b in brands},
    'cat_colors': {c: CAT_COLORS.get(c, DEFAULT_CAT_COLOR) for c in cats},
    'monthly': monthly,
    'bm': bm,
    'sm': sm,
    'cm': cm,
    'bcm': bcm,
    'sbm': sbm,
    'scm': scm,
    'products': products,
    'store_top_cnps': store_top_cnps,
    'generated': datetime.now().strftime('%d/%m/%Y %H:%M'),
}

data_json = json.dumps(DATA, ensure_ascii=False, separators=(',', ':'))
print(f"  JSON: {len(data_json)/1024:.0f} KB")

# ─── 5. HTML ──────────────────────────────────────────────────────────────────
HTML = f"""<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Auchan Dashboard — Pierre Fabre</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
:root{{
  --pf:#E2001A; --pf2:#b5001b; --amber:#F7A600; --green:#10b981;
  --red:#ef4444; --card:#fff; --bg:#f1f5f9; --border:#e2e8f0;
  --text:#1e293b; --muted:#64748b;
}}
body{{background:var(--bg);color:var(--text);font-family:'Segoe UI',sans-serif;font-size:14px}}
.topbar{{background:#1e3a5f;padding:12px 20px;display:flex;align-items:center;justify-content:space-between}}
.topbar-left{{display:flex;align-items:center;gap:12px}}
.pf-logo{{width:36px;height:36px;border-radius:50%;object-fit:cover;box-shadow:0 0 0 2px rgba(255,255,255,.25)}}
.topbar h1{{color:#fff;font-size:16px;font-weight:700;margin:0}}
.topbar .sub{{color:#94a3b8;font-size:11px}}
.filter-bar{{background:#fff;border-bottom:1px solid var(--border);padding:8px 20px;display:flex;flex-wrap:wrap;gap:8px;align-items:center}}
.filter-bar label{{font-size:11px;font-weight:600;color:var(--muted);margin-bottom:2px;display:block}}
.filter-bar select,.filter-bar input{{font-size:12px;border:1px solid var(--border);border-radius:6px;padding:4px 8px;background:#f8fafc;color:var(--text)}}
.filter-bar select:focus,.filter-bar input:focus{{outline:none;border-color:var(--pf)}}
.tab-nav{{background:#fff;border-bottom:2px solid var(--border);padding:0 20px;display:flex;gap:0;overflow-x:auto}}
.tab-btn{{padding:10px 16px;font-size:12px;font-weight:600;color:var(--muted);border:none;background:none;border-bottom:2px solid transparent;margin-bottom:-2px;cursor:pointer;white-space:nowrap;transition:.15s}}
.tab-btn.active{{color:var(--pf);border-bottom-color:var(--pf)}}
.tab-btn:hover{{color:var(--pf)}}
.content{{padding:16px 20px;min-height:calc(100vh - 140px)}}
.tab-panel{{display:none}}.tab-panel.active{{display:block}}
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:16px}}
.kpi{{background:var(--card);border-radius:10px;padding:16px;border:1px solid var(--border);position:relative;overflow:hidden}}
.kpi::after{{content:'';position:absolute;right:-10px;top:-10px;width:60px;height:60px;border-radius:50%;background:var(--pf);opacity:.05}}
.kpi .label{{font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}}
.kpi .value{{font-size:22px;font-weight:700;color:#1e3a5f;margin:4px 0 2px}}
.kpi .value.kv-pos{{color:#10b981}}.kpi .value.kv-neg{{color:#ef4444}}
.kpi .sub{{font-size:11px;color:var(--muted)}}
.kpi .badge-up{{color:var(--green);font-weight:700}}
.kpi .badge-dn{{color:var(--red);font-weight:700}}
.btn-evol{{padding:4px 12px;border:1px solid #cbd5e1;border-radius:6px;background:#f8fafc;color:#475569;font-size:12px;font-weight:600;cursor:pointer;transition:all .15s}}
.btn-evol:hover{{background:#e2e8f0}}
.btn-evol-active{{background:#1e3a5f!important;color:#fff!important;border-color:#1e3a5f!important}}
.chart-card{{background:var(--card);border-radius:10px;padding:16px;border:1px solid var(--border);margin-bottom:14px}}
.chart-card h5{{font-size:13px;font-weight:700;color:var(--pf);margin-bottom:12px}}
.row-charts{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
@media(max-width:768px){{.row-charts{{grid-template-columns:1fr}}}}
.table-wrap{{overflow-x:auto}}
table.dt{{width:100%;border-collapse:collapse;font-size:12px}}
table.dt th{{background:var(--pf);color:#fff;padding:8px 10px;text-align:left;font-weight:600;white-space:nowrap;cursor:pointer;user-select:none}}
table.dt th:hover{{background:var(--pf2)}}
table.dt td{{padding:7px 10px;border-bottom:1px solid var(--border);white-space:nowrap}}
table.dt tr:hover td{{background:#f8fafc}}
.badge-innov{{background:#fef3c7;color:#92400e;border-radius:4px;padding:1px 5px;font-size:10px;font-weight:700}}
.badge-grow{{color:var(--green);font-weight:700}}
.badge-drop{{color:var(--red);font-weight:700}}
.search-box{{position:relative;margin-bottom:10px}}
.search-box input{{width:100%;padding:7px 12px 7px 32px;border:1px solid var(--border);border-radius:8px;font-size:13px}}
.search-box::before{{content:'';position:absolute;left:10px;top:50%;transform:translateY(-50%);font-size:12px}}
.section-title{{font-size:13px;font-weight:700;color:var(--pf);margin-bottom:10px;padding-bottom:6px;border-bottom:2px solid var(--pf);display:inline-block}}
.alert-card{{background:var(--card);border-radius:10px;border:1px solid var(--border);overflow:hidden;margin-bottom:14px}}
.alert-card .ac-header{{background:var(--pf);color:#fff;padding:10px 14px;font-size:12px;font-weight:700}}
.alert-card .ac-row{{display:flex;align-items:center;padding:8px 14px;border-bottom:1px solid var(--border);gap:10px}}
.alert-card .ac-row:last-child{{border-bottom:none}}
.ac-rank{{width:20px;height:20px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;flex-shrink:0}}
.ac-rank.grow{{background:#d1fae5;color:#065f46}}
.ac-rank.drop{{background:#fee2e2;color:#991b1b}}
.ac-name{{flex:1;font-size:12px;color:var(--text)}}
.ac-val{{font-size:12px;font-weight:700}}
.innov-weight{{background:linear-gradient(135deg,#fef3c7,#fde68a);border-radius:10px;padding:16px;text-align:center;margin-bottom:14px}}
.innov-weight .pct{{font-size:32px;font-weight:800;color:#92400e}}
.innov-weight .lbl{{font-size:12px;color:#78350f}}
.pagination{{display:flex;gap:6px;align-items:center;padding:8px 0}}
.pagination button{{padding:4px 10px;border:1px solid var(--border);border-radius:6px;background:#fff;font-size:12px;cursor:pointer}}
.pagination button.active{{background:var(--pf);color:#fff;border-color:var(--pf)}}
.pagination button:hover:not(.active){{background:var(--bg)}}
.pagination .page-info{{font-size:12px;color:var(--muted)}}
</style>
</head>
<body>

<!-- TOP BAR -->
<div class="topbar">
  <div class="topbar-left">
    <img class="pf-logo" src="{PF_LOGO_B64}" alt="Pierre Fabre">
    <div>
      <h1>Auchan Sell-Out</h1>
      <div class="sub">Dados até 15 Mai 2026 · Gerado em {DATA['generated']}</div>
    </div>
  </div>
  <div style="color:#94a3b8;font-size:11px;text-align:right">Filtros aplicados<br>a todos os painéis</div>
</div>

<!-- FILTER BAR -->
<div class="filter-bar">
  <div>
    <label>Período</label>
    <select id="fPeriod" onchange="applyFilters()">
    </select>
  </div>
  <div>
    <label>Marca PF</label>
    <select id="fBrand" onchange="applyFilters()">
      <option value="">Todas as Marcas</option>
    </select>
  </div>
  <div>
    <label>Categoria 2</label>
    <select id="fCat" onchange="applyFilters()">
      <option value="">Todas as Categorias</option>
    </select>
  </div>
  <div>
    <label>Loja</label>
    <select id="fStore" onchange="applyFilters()">
      <option value="">Todas as Lojas</option>
    </select>
  </div>
  <div style="min-width:160px">
    <label>Pesquisar CNP/Produto</label>
    <input type="text" id="fCNP" placeholder="CNP ou descrição..." oninput="applyFilters()" style="width:180px">
  </div>
  <div style="margin-left:auto;font-size:11px;color:var(--muted)">
    <span id="fSummary"></span>
  </div>
</div>

<!-- TABS -->
<div class="tab-nav">
  <button class="tab-btn active" onclick="switchTab('exec',this)">Resumo Executivo</button>
  <button class="tab-btn" onclick="switchTab('stores',this)">Lojas</button>
  <button class="tab-btn" onclick="switchTab('brands',this)">Marcas</button>
  <button class="tab-btn" onclick="switchTab('cats',this)">Categorias</button>
  <button class="tab-btn" onclick="switchTab('products',this)">Produtos</button>
  <button class="tab-btn" onclick="switchTab('growth',this)">Crescimento</button>
  <button class="tab-btn" onclick="switchTab('innov',this)">Inovação</button>
  <button class="tab-btn" onclick="switchTab('alerts',this)">Alertas</button>
</div>

<div class="content">

<!-- TAB: EXEC SUMMARY -->
<div id="tab-exec" class="tab-panel active">
  <div class="kpi-grid" id="kpiGrid"></div>
  <div class="row-charts">
    <div class="chart-card" style="grid-column:1/-1">
      <h5>Evolução Mensal de Vendas (€) — 2025 vs 2026</h5>
      <canvas id="chartMonthly" height="80"></canvas>
    </div>
  </div>
  <div class="row-charts">
    <div class="chart-card">
      <h5>Mix por Marca — Período Seleccionado</h5>
      <div style="position:relative;height:260px"><canvas id="chartBrandMix"></canvas></div>
    </div>
    <div class="chart-card">
      <h5>Mix por Categoria — Período Seleccionado</h5>
      <div style="position:relative;height:260px"><canvas id="chartCatMix"></canvas></div>
    </div>
  </div>
</div>

<!-- TAB: LOJAS -->
<div id="tab-stores" class="tab-panel">
  <div class="row-charts">
    <div class="chart-card">
      <h5>Top 15 Lojas por Vendas (€) — CY vs LY</h5>
      <canvas id="chartTopStores" height="320"></canvas>
    </div>
    <div class="chart-card">
      <h5>Crescimento vs Ano Anterior (%) — Top/Bottom 10</h5>
      <canvas id="chartStoreGrowth" height="320"></canvas>
    </div>
  </div>
  <div class="chart-card">
    <h5>Ranking Completo de Lojas</h5>
    <div class="table-wrap">
      <table class="dt" id="tblStores">
        <thead><tr>
          <th onclick="sortTable('tblStores',0)">#</th>
          <th onclick="sortTable('tblStores',1)">Loja</th>
          <th onclick="sortTable('tblStores',2)">Vendas CY (€)</th>
          <th onclick="sortTable('tblStores',3)">Vendas LY (€)</th>
          <th onclick="sortTable('tblStores',4)">Crescimento</th>
          <th onclick="sortTable('tblStores',5)">Unid. CY</th>
        </tr></thead>
        <tbody id="tblStoresBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- TAB: MARCAS -->
<div id="tab-brands" class="tab-panel">
  <div class="chart-card">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
      <h5 style="margin:0">Vendas por Marca — Evolução Mensal (€)</h5>
      <div id="brandEvolToggle" style="display:flex;gap:6px">
        <button onclick="setBrandEvolView('2025')" class="btn-evol" id="bevol-2025">2025</button>
        <button onclick="setBrandEvolView('2026')" class="btn-evol" id="bevol-2026">2026</button>
        <button onclick="setBrandEvolView('all')" class="btn-evol btn-evol-active" id="bevol-all">Todos</button>
      </div>
    </div>
    <canvas id="chartBrandEvol" height="90"></canvas>
  </div>
  <div class="row-charts">
    <div class="chart-card">
      <h5>Comparação CY vs LY por Marca (€)</h5>
      <canvas id="chartBrandCmp" height="250"></canvas>
    </div>
    <div class="chart-card">
      <h5>Crescimento por Marca (%)</h5>
      <canvas id="chartBrandGrowth" height="250"></canvas>
    </div>
  </div>
  <div class="chart-card">
    <h5>Detalhe por Marca e Categoria</h5>
    <div class="table-wrap">
      <table class="dt" id="tblBrandCat">
        <thead><tr>
          <th onclick="sortTable('tblBrandCat',0)">Marca</th>
          <th onclick="sortTable('tblBrandCat',1)">Categoria</th>
          <th onclick="sortTable('tblBrandCat',2)">Vendas CY (€)</th>
          <th onclick="sortTable('tblBrandCat',3)">Vendas LY (€)</th>
          <th onclick="sortTable('tblBrandCat',4)">Crescimento</th>
          <th onclick="sortTable('tblBrandCat',5)">Unid. CY</th>
        </tr></thead>
        <tbody id="tblBrandCatBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- TAB: CATEGORIAS -->
<div id="tab-cats" class="tab-panel">
  <div class="row-charts">
    <div class="chart-card">
      <h5>Vendas por Categoria (€) — CY vs LY</h5>
      <canvas id="chartCatCmp" height="280"></canvas>
    </div>
    <div class="chart-card">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
        <h5 style="margin:0">Evolução Mensal por Categoria (€)</h5>
        <button onclick="document.getElementById('fCat').value='';applyFilters()" style="font-size:11px;padding:4px 12px;border:1px solid var(--border);border-radius:6px;background:#f8fafc;color:var(--muted);cursor:pointer;font-weight:600">Resetar</button>
      </div>
      <canvas id="chartCatEvol" height="280"></canvas>
    </div>
  </div>
  <div class="chart-card">
    <h5>Crescimento por Categoria vs Ano Anterior (%)</h5>
    <canvas id="chartCatGrowth" height="80"></canvas>
  </div>
  <div class="chart-card">
    <h5>Performance por Categoria e Marca</h5>
    <div class="table-wrap">
      <table class="dt" id="tblCatBrand">
        <thead><tr>
          <th onclick="sortTable('tblCatBrand',0)">Categoria</th>
          <th onclick="sortTable('tblCatBrand',1)">Marca</th>
          <th onclick="sortTable('tblCatBrand',2)">Vendas CY (€)</th>
          <th onclick="sortTable('tblCatBrand',3)">Vendas LY (€)</th>
          <th onclick="sortTable('tblCatBrand',4)">Crescimento</th>
          <th onclick="sortTable('tblCatBrand',5)">Unid. CY</th>
        </tr></thead>
        <tbody id="tblCatBrandBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- TAB: PRODUTOS -->
<div id="tab-products" class="tab-panel">
  <div class="kpi-grid" id="kpiProds"></div>
  <div class="search-box">
    <input type="text" id="prodSearch" placeholder="Pesquisar CNP, descrição, marca, gama..." oninput="renderProductsTable()">
  </div>
  <div class="chart-card" style="padding:12px">
    <div class="table-wrap">
      <table class="dt" id="tblProds">
        <thead><tr>
          <th onclick="sortProdTable(0)">CNP</th>
          <th onclick="sortProdTable(1)">Descrição</th>
          <th onclick="sortProdTable(2)">Marca</th>
          <th onclick="sortProdTable(3)">Categoria</th>
          <th onclick="sortProdTable(4)">Gama</th>
          <th onclick="sortProdTable(5)">Vendas CY (€)</th>
          <th onclick="sortProdTable(6)">Vendas LY (€)</th>
          <th onclick="sortProdTable(7)">Cresc. %</th>
          <th onclick="sortProdTable(8)">Unid. CY</th>
          <th onclick="sortProdTable(9)">Unid. LY</th>
        </tr></thead>
        <tbody id="tblProdsBody"></tbody>
      </table>
    </div>
    <div class="pagination" id="prodPagination"></div>
  </div>
</div>

<!-- TAB: CRESCIMENTO -->
<div id="tab-growth" class="tab-panel">
  <div class="row-charts">
    <div class="chart-card">
      <h5>Top 15 Produtos — Maior Crescimento Absoluto (€)</h5>
      <canvas id="chartGrowthTop" height="350"></canvas>
    </div>
    <div class="chart-card">
      <h5>Top 15 Produtos — Maior Quebra Absoluta (€)</h5>
      <canvas id="chartGrowthBot" height="350"></canvas>
    </div>
  </div>
  <div class="row-charts">
    <div class="chart-card">
      <h5>Top 15 Produtos — Maior Crescimento % vs LY</h5>
      <canvas id="chartGrowthTopPct" height="350"></canvas>
    </div>
    <div class="chart-card">
      <h5>Top 15 Produtos — Maior Quebra % vs LY</h5>
      <canvas id="chartGrowthBotPct" height="350"></canvas>
    </div>
  </div>
</div>

<!-- TAB: INOVAÇÃO -->
<div id="tab-innov" class="tab-panel">
  <div class="innov-weight">
    <div class="pct" id="innovPct">—</div>
    <div class="lbl">Peso da Inovação nas Vendas 2026 (produtos sem histórico 2025)</div>
  </div>
  <div class="kpi-grid" id="kpiInnov"></div>
  <div class="chart-card">
    <h5>Top 20 Produtos de Inovação — Vendas 2026 (€)</h5>
    <canvas id="chartInnov" height="130"></canvas>
  </div>
  <div class="chart-card">
    <h5>Lista Completa Produtos de Inovação</h5>
    <div class="table-wrap">
      <table class="dt" id="tblInnov">
        <thead><tr>
          <th onclick="sortTable('tblInnov',0)">CNP</th>
          <th onclick="sortTable('tblInnov',1)">Descrição</th>
          <th onclick="sortTable('tblInnov',2)">Marca</th>
          <th onclick="sortTable('tblInnov',3)">Categoria</th>
          <th onclick="sortTable('tblInnov',4)">Vendas 2026 (€)</th>
          <th onclick="sortTable('tblInnov',5)">Unidades 2026</th>
        </tr></thead>
        <tbody id="tblInnovBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- TAB: ALERTAS -->
<div id="tab-alerts" class="tab-panel">
  <div class="row-charts">
    <div>
      <div class="section-title">Produtos</div>
      <div class="alert-card">
        <div class="ac-header">Top 10 Produtos com Maior Crescimento (€)</div>
        <div id="alertProdUp"></div>
      </div>
      <div class="alert-card">
        <div class="ac-header">Top 10 Produtos com Maior Quebra (€)</div>
        <div id="alertProdDn"></div>
      </div>
    </div>
    <div>
      <div class="section-title">Lojas</div>
      <div class="alert-card">
        <div class="ac-header">Top 10 Lojas com Maior Crescimento (€)</div>
        <div id="alertStoreUp"></div>
      </div>
      <div class="alert-card">
        <div class="ac-header">Top 10 Lojas com Maior Quebra (€)</div>
        <div id="alertStoreDn"></div>
      </div>
    </div>
  </div>
</div>

</div><!-- /content -->

<script>
const DATA = {data_json};

// ─── CHART REGISTRY ───────────────────────────────────────────────────────────
const CH = {{}};
function destroyChart(id) {{ if (CH[id]) {{ CH[id].destroy(); delete CH[id]; }} }}

// Desactivar datalabels globalmente; só activos por chart explícito (plugins:[ChartDataLabels])
if (typeof ChartDataLabels !== 'undefined') {{
  try {{ Chart.unregister(ChartDataLabels); }} catch(e) {{}}
}}

// ─── FILTERS ──────────────────────────────────────────────────────────────────
let F = {{ period: 'ytd2026', brand: '', cat: '', store: '', cnp: '' }};
let activeTab = 'exec';
let _brandEvolView = 'all';

// ─── INIT DROPDOWNS ──────────────────────────────────────────────────────────
function initFilters() {{
  const sel = document.getElementById('fPeriod');
  DATA.period_options.forEach(o => {{
    const opt = document.createElement('option');
    opt.value = o.val;
    opt.text = o.label;
    if (o.val === '__sep__') opt.disabled = true;
    sel.appendChild(opt);
  }});
  const fb = document.getElementById('fBrand');
  DATA.brands.forEach(b => {{
    const opt = document.createElement('option');
    opt.value = b; opt.text = b; fb.appendChild(opt);
  }});
  const fc = document.getElementById('fCat');
  DATA.cats.forEach(c => {{
    const opt = document.createElement('option');
    opt.value = c; opt.text = c; fc.appendChild(opt);
  }});
  const fs = document.getElementById('fStore');
  DATA.stores.forEach(s => {{
    const opt = document.createElement('option');
    opt.value = s; opt.text = s; fs.appendChild(opt);
  }});
}}

function applyFilters() {{
  F.period = document.getElementById('fPeriod').value;
  F.brand  = document.getElementById('fBrand').value;
  F.cat    = document.getElementById('fCat').value;
  F.store  = document.getElementById('fStore').value;
  F.cnp    = document.getElementById('fCNP').value.toLowerCase().trim();
  renderTab(activeTab);
  updateFilterSummary();
}}

function updateFilterSummary() {{
  const parts = [];
  if (F.brand) parts.push(F.brand);
  if (F.cat)   parts.push(F.cat);
  if (F.store) parts.push(F.store);
  if (F.cnp)   parts.push('CNP:"'+F.cnp+'"');
  document.getElementById('fSummary').textContent = parts.length ? parts.join(' · ') : '';
}}

// ─── PERIOD HELPERS ──────────────────────────────────────────────────────────
function getPeriodMonths() {{
  const p = F.period;
  if (p === 'ytd2026') return {{ cy: DATA.ym_2026, ly: DATA.ym_2025.slice(0, DATA.ym_2026.length) }};
  if (p === 'all2025') return {{ cy: DATA.ym_2025, ly: [] }};
  if (/^\d{{4}}-\d{{2}}$/.test(p)) {{
    const yr = parseInt(p.slice(0,4)), mo = p.slice(4);
    const lyYm = (yr-1) + mo;
    return {{ cy: [p], ly: DATA.months.includes(lyYm) ? [lyYm] : [] }};
  }}
  return {{ cy: DATA.months, ly: [] }};
}}

function fmt(v, dec=0) {{
  if (v === null || v === undefined || isNaN(v)) return '—';
  return new Intl.NumberFormat('pt-PT',{{minimumFractionDigits:dec,maximumFractionDigits:dec}}).format(v);
}}
function fmtEur(v) {{ return '€' + fmt(v); }}
function fmtPct(v) {{
  if (v === null || isNaN(v) || !isFinite(v)) return '—';
  const s = v >= 0 ? '+' : '';
  return s + fmt(v,1) + '%';
}}
function growthClass(v) {{
  if (!isFinite(v) || isNaN(v)) return '';
  return v >= 0 ? 'badge-grow' : 'badge-drop';
}}

// ─── DATA HELPERS ─────────────────────────────────────────────────────────────
function sumMonthly(yms) {{
  let v = 0, u = 0;
  yms.forEach(ym => {{
    const d = DATA.monthly[ym];
    if (d) {{ v += d.v; u += d.u; }}
  }});
  return {{ v, u }};
}}

function sumBM(brand, yms) {{
  let v = 0, u = 0;
  yms.forEach(ym => {{
    const d = DATA.bm[brand+'|'+ym];
    if (d) {{ v += d.v; u += d.u; }}
  }});
  return {{ v, u }};
}}

function sumSM(store, yms) {{
  let v = 0, u = 0;
  yms.forEach(ym => {{
    const d = DATA.sm[store+'|'+ym];
    if (d) {{ v += d.v; u += d.u; }}
  }});
  return {{ v, u }};
}}

function sumCM(cat, yms) {{
  let v = 0, u = 0;
  yms.forEach(ym => {{
    const d = DATA.cm[cat+'|'+ym];
    if (d) {{ v += d.v; u += d.u; }}
  }});
  return {{ v, u }};
}}

function sumBCM(brand, cat, yms) {{
  let v = 0, u = 0;
  yms.forEach(ym => {{
    const d = DATA.bcm[brand+'|'+cat+'|'+ym];
    if (d) {{ v += d.v; u += d.u; }}
  }});
  return {{ v, u }};
}}

function sumSBM(store, brand, yms) {{
  let v = 0, u = 0;
  yms.forEach(ym => {{
    const d = DATA.sbm[store+'|'+brand+'|'+ym];
    if (d) {{ v += d.v; u += d.u; }}
  }});
  return {{ v, u }};
}}

function sumSCM(store, cat, yms) {{
  let v = 0, u = 0;
  yms.forEach(ym => {{
    const d = DATA.scm[store+'|'+cat+'|'+ym];
    if (d) {{ v += d.v; u += d.u; }}
  }});
  return {{ v, u }};
}}

function sumProd(prod, yms) {{
  let v = 0, u = 0;
  yms.forEach(ym => {{
    const d = prod.m[ym];
    if (d) {{ v += d.v; u += d.u; }}
  }});
  return {{ v, u }};
}}

// Get effective CY total for exec with all active filters
function getFilteredTotal(yms) {{
  if (F.brand && F.cat && F.store) {{
    // brand × cat × store: approximation via products
    return sumFilteredProds(yms);
  }}
  if (F.brand && F.cat) return sumBCM(F.brand, F.cat, yms);
  if (F.brand && F.store) return sumSBM(F.store, F.brand, yms);
  if (F.cat && F.store) return sumSCM(F.store, F.cat, yms);
  if (F.brand) return sumBM(F.brand, yms);
  if (F.cat)  return sumCM(F.cat, yms);
  if (F.store) return sumSM(F.store, yms);
  return sumMonthly(yms);
}}

function sumFilteredProds(yms) {{
  const prods = getFilteredProducts();
  let v = 0, u = 0;
  prods.forEach(p => {{ const d = sumProd(p, yms); v += d.v; u += d.u; }});
  return {{ v, u }};
}}

function getFilteredProducts() {{
  return DATA.products.filter(p => {{
    if (F.brand && p.marca !== F.brand) return false;
    if (F.cat   && p.cat   !== F.cat)  return false;
    if (F.store) {{
      const tops = DATA.store_top_cnps[F.store] || [];
      if (!tops.includes(p.cnp)) return false;
    }}
    return true;
  }});
}}

// ─── TAB SWITCHING ───────────────────────────────────────────────────────────
function switchTab(id, btn) {{
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  if (btn) btn.classList.add('active');
  document.getElementById('tab-'+id).classList.add('active');
  activeTab = id;
  renderTab(id);
}}

function renderTab(id) {{
  switch(id) {{
    case 'exec':     renderExec(); break;
    case 'stores':   renderStores(); break;
    case 'brands':   renderBrands(); break;
    case 'cats':     renderCats(); break;
    case 'products': renderProductsTable(); break;
    case 'growth':   renderGrowth(); break;
    case 'innov':    renderInnov(); break;
    case 'alerts':   renderAlerts(); break;
  }}
}}

// ─── TAB: EXEC ────────────────────────────────────────────────────────────────
function renderExec() {{
  const pm = getPeriodMonths();
  const cy = getFilteredTotal(pm.cy);
  const ly = getFilteredTotal(pm.ly);
  const growV = ly.v > 0 ? ((cy.v - ly.v) / ly.v * 100) : null;
  const growU = ly.u > 0 ? ((cy.u - ly.u) / ly.u * 100) : null;

  const isYTD = F.period === 'ytd2026';
  const cyLabel = isYTD ? 'YTD 2026' : (F.period === 'all2025' ? '2025' : DATA.month_labels[F.period] || 'CY');
  const lyLabel = isYTD ? 'YTD 2025' : 'LY';

  // Build KPIs
  const kpis = [
    {{
      label: `Vendas ${{cyLabel}} (€)`,
      value: fmtEur(cy.v),
      sub: ly.v > 0 ? `${{lyLabel}}: ${{fmtEur(ly.v)}}` : 'Sem comparativo',
      badge: growV !== null ? fmtPct(growV) : null,
      grow: growV
    }},
    {{
      label: `Unidades ${{cyLabel}}`,
      value: fmt(cy.u),
      sub: ly.u > 0 ? `${{lyLabel}}: ${{fmt(ly.u)}} un.` : 'Sem comparativo',
      badge: growU !== null ? fmtPct(growU) : null,
      grow: growU
    }},
    {{
      label: 'Preço Médio CY (€)',
      value: cy.u > 0 ? fmtEur(cy.v / cy.u) : '—',
      sub: ly.u > 0 ? `LY: ${{fmtEur(ly.v/ly.u)}}` : '',
      badge: null,
      grow: null
    }},
    {{
      label: 'Crescimento Absoluto (€)',
      value: ly.v > 0 ? (cy.v - ly.v >= 0 ? '+' : '') + fmtEur(cy.v - ly.v) : '—',
      sub: 'vs período homólogo',
      badge: growV !== null ? fmtPct(growV) : null,
      grow: growV
    }},
  ];
  document.getElementById('kpiGrid').innerHTML = kpis.map(k => `
    <div class="kpi">
      <div class="label">${{k.label}}</div>
      <div class="value ${{k.grow !== null ? (k.grow >= 0 ? 'kv-pos' : 'kv-neg') : ''}}">${{k.value}}</div>
      <div class="sub">
        ${{k.sub}}
        ${{k.badge !== null ? `<span class="${{k.grow >= 0 ? 'badge-up' : 'badge-dn'}}"> ${{k.badge}}</span>` : ''}}
      </div>
    </div>`).join('');

  // Monthly trend chart
  destroyChart('chartMonthly');
  const allMonthsSorted = [...DATA.ym_2025, ...DATA.ym_2026];
  const m25 = DATA.ym_2025.map(ym => {{
    const d = getFilteredTotal([ym]);
    return d.v;
  }});
  const m26 = DATA.ym_2026.map(ym => {{
    const d = getFilteredTotal([ym]);
    return d.v;
  }});
  const labels25 = DATA.ym_2025.map(ym => DATA.month_labels[ym]);
  const labels26 = DATA.ym_2026.map(ym => DATA.month_labels[ym]);
  const allLabels = [...new Set([...labels25.map((l,i)=>l.replace('2025','')), ...labels26.map((l,i)=>l.replace('2026',''))])];
  const maxLen = Math.max(m25.length, m26.length);
  const labelsShort = DATA.ym_2025.slice(0,maxLen).map(ym => DATA.month_labels[ym].replace(' 2025',''));

  CH['chartMonthly'] = new Chart(document.getElementById('chartMonthly'), {{
    type: 'bar',
    data: {{
      labels: DATA.ym_2025.map(ym => DATA.month_labels[ym].replace(' 2025','')),
      datasets: [
        {{ label: '2025', data: m25, backgroundColor: '#94a3b880', borderColor: '#94a3b8', borderWidth: 1.5 }},
        {{ label: '2026', data: [...m26, ...Array(Math.max(0,m25.length-m26.length)).fill(null)],
           backgroundColor: '#1e3a5fcc', borderColor: '#1e3a5f', borderWidth: 1.5 }},
      ]
    }},
    options: {{
      responsive: true, maintainAspectRatio: true,
      plugins: {{ legend: {{ position: 'top' }}, tooltip: {{ callbacks: {{ label: c => ' €' + fmt(c.parsed.y) }} }} }},
      scales: {{ y: {{ ticks: {{ callback: v => '€'+fmt(v) }} }} }}
    }}
  }});

  // Brand mix donut
  destroyChart('chartBrandMix');
  const brands = F.brand ? [F.brand] : DATA.brands;
  const brandVals = brands.map(b => {{
    if (F.cat && F.store) return sumFilteredProds(pm.cy);
    if (F.cat) return sumBCM(b, F.cat, pm.cy).v;
    if (F.store) return sumSBM(F.store, b, pm.cy).v;
    return sumBM(b, pm.cy).v;
  }});
  CH['chartBrandMix'] = new Chart(document.getElementById('chartBrandMix'), {{
    plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
    type: 'doughnut',
    data: {{
      labels: brands,
      datasets: [{{ data: brandVals, backgroundColor: brands.map(b => DATA.brand_colors[b] || '#64748b'), borderWidth: 2 }}]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{
        legend: {{ position: 'right', labels: {{ font: {{ size: 11 }} }} }},
        tooltip: {{ callbacks: {{ label: c => ' €' + fmt(c.parsed) }} }},
        datalabels: {{
          display: true,
          formatter: (value, ctx) => {{
            const sum = ctx.dataset.data.reduce((a,b) => a+(b||0), 0);
            const pct = sum > 0 ? value/sum*100 : 0;
            return pct >= 4 ? pct.toFixed(1)+'%' : '';
          }},
          color: '#fff',
          font: {{ weight: 'bold', size: 11 }}
        }}
      }}
    }}
  }});

  // Cat mix donut
  destroyChart('chartCatMix');
  const cats = F.cat ? [F.cat] : DATA.cats;
  const catVals = cats.map(c => {{
    if (F.brand && F.store) return sumFilteredProds(pm.cy);
    if (F.brand) return sumBCM(F.brand, c, pm.cy).v;
    if (F.store) return sumSCM(F.store, c, pm.cy).v;
    return sumCM(c, pm.cy).v;
  }});
  CH['chartCatMix'] = new Chart(document.getElementById('chartCatMix'), {{
    plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
    type: 'doughnut',
    data: {{
      labels: cats,
      datasets: [{{ data: catVals, backgroundColor: cats.map(c => DATA.cat_colors[c] || '#94a3b8'), borderWidth: 2 }}]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{
        legend: {{ position: 'right', labels: {{ font: {{ size: 11 }} }} }},
        tooltip: {{ callbacks: {{ label: c => ' €' + fmt(c.parsed) }} }},
        datalabels: {{
          display: true,
          formatter: (value, ctx) => {{
            const sum = ctx.dataset.data.reduce((a,b) => a+(b||0), 0);
            const pct = sum > 0 ? value/sum*100 : 0;
            return pct >= 4 ? pct.toFixed(1)+'%' : '';
          }},
          color: '#fff',
          font: {{ weight: 'bold', size: 11 }}
        }}
      }}
    }}
  }});
}}

// ─── TAB: STORES ─────────────────────────────────────────────────────────────
function renderStores() {{
  const pm = getPeriodMonths();
  const stores = DATA.stores;

  const rows = stores.map(s => {{
    let cy, ly;
    if (F.brand && F.cat) {{
      // Approximate via products
      const prods = DATA.products.filter(p => p.marca===F.brand && p.cat===F.cat);
      cy = {{ v:0,u:0 }}; ly = {{ v:0,u:0 }};
      prods.forEach(p => {{
        const tops = DATA.store_top_cnps[s] || [];
        if (tops.includes(p.cnp)) {{
          const cd = sumProd(p, pm.cy); cy.v += cd.v; cy.u += cd.u;
          const ld = sumProd(p, pm.ly); ly.v += ld.v; ly.u += ld.u;
        }}
      }});
    }} else if (F.brand) {{
      cy = sumSBM(s, F.brand, pm.cy);
      ly = sumSBM(s, F.brand, pm.ly);
    }} else if (F.cat) {{
      cy = sumSCM(s, F.cat, pm.cy);
      ly = sumSCM(s, F.cat, pm.ly);
    }} else {{
      cy = sumSM(s, pm.cy);
      ly = sumSM(s, pm.ly);
    }}
    const grow = ly.v > 0 ? (cy.v - ly.v)/ly.v*100 : null;
    return {{ store:s, cy:cy.v, ly:ly.v, u:cy.u, grow }};
  }}).filter(r => r.cy > 0 || r.ly > 0);

  rows.sort((a,b) => b.cy - a.cy);

  // Top 15 chart
  const top15 = rows.slice(0,15).reverse();
  destroyChart('chartTopStores');
  CH['chartTopStores'] = new Chart(document.getElementById('chartTopStores'), {{
    plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
    type: 'bar',
    data: {{
      labels: top15.map(r => r.store.replace('Auchan ','').replace('Auchan','').trim()),
      datasets: [
        {{ label: 'CY', data: top15.map(r => r.cy), backgroundColor: '#1e3a5fcc' }},
        {{ label: 'LY', data: top15.map(r => r.ly), backgroundColor: '#94a3b860' }},
      ]
    }},
    options: {{
      indexAxis: 'y', responsive: true,
      plugins: {{
        tooltip: {{ callbacks: {{ label: c => ' €' + fmt(c.parsed.x) }} }},
        datalabels: {{
          display: true,
          anchor: 'center', align: 'center',
          formatter: (v, ctx) => {{
            if (!v || v < 100) return '';
            return v >= 10000 ? '€'+(v/1000).toFixed(0)+'k' : '€'+fmt(v,0);
          }},
          font: {{ size: 9, weight: 'bold' }},
          color: ctx => ctx.datasetIndex === 0 ? '#fff' : '#334155'
        }}
      }},
      scales: {{ x: {{ ticks: {{ callback: v => '€'+fmt(v) }} }} }}
    }}
  }});

  // Growth chart (top5 + bottom5)
  const withGrow = rows.filter(r => r.grow !== null);
  withGrow.sort((a,b) => b.grow - a.grow);
  const topG = withGrow.slice(0,5);
  const botG = withGrow.slice(-5).reverse();
  const combined = [...topG, ...botG];
  destroyChart('chartStoreGrowth');
  CH['chartStoreGrowth'] = new Chart(document.getElementById('chartStoreGrowth'), {{
    plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
    type: 'bar',
    data: {{
      labels: combined.map(r => r.store.replace('Auchan ','').trim()),
      datasets: [{{
        label: 'Crescimento %',
        data: combined.map(r => r.grow),
        backgroundColor: combined.map(r => r.grow >= 0 ? '#10b98180' : '#ef444480'),
        borderColor:     combined.map(r => r.grow >= 0 ? '#10b981' : '#ef4444'),
        borderWidth: 1.5
      }}]
    }},
    options: {{
      indexAxis: 'y', responsive: true,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{ callbacks: {{ label: c => ' ' + fmtPct(c.parsed.x) }} }},
        datalabels: {{
          display: true,
          anchor: 'center', align: 'center',
          formatter: v => v !== null ? fmtPct(v) : '',
          color: '#fff',
          font: {{ size: 10, weight: 'bold' }},
          textStrokeColor: '#00000040',
          textStrokeWidth: 2
        }}
      }},
      scales: {{ x: {{ ticks: {{ callback: v => v.toFixed(0)+'%' }} }} }}
    }}
  }});

  // Table
  const tbody = document.getElementById('tblStoresBody');
  tbody.innerHTML = rows.map((r,i) => `
    <tr>
      <td>${{i+1}}</td>
      <td>${{r.store}}</td>
      <td>${{fmtEur(r.cy)}}</td>
      <td>${{fmtEur(r.ly)}}</td>
      <td class="${{growthClass(r.grow)}}">${{fmtPct(r.grow)}}</td>
      <td>${{fmt(r.u)}}</td>
    </tr>`).join('');
}}

// ─── TAB: BRANDS ─────────────────────────────────────────────────────────────
function setBrandEvolView(v) {{
  _brandEvolView = v;
  renderBrands();
}}

function renderBrands() {{
  const pm = getPeriodMonths();
  const brands = F.brand ? [F.brand] : DATA.brands;

  // Monthly evolution line chart
  destroyChart('chartBrandEvol');
  const allMonths = [...DATA.ym_2025, ...DATA.ym_2026];
  const evolMonths = _brandEvolView === '2025' ? DATA.ym_2025
                   : _brandEvolView === '2026' ? DATA.ym_2026
                   : allMonths;
  CH['chartBrandEvol'] = new Chart(document.getElementById('chartBrandEvol'), {{
    type: 'line',
    data: {{
      labels: evolMonths.map(ym => DATA.month_labels[ym]),
      datasets: brands.map(b => {{
        const vals = evolMonths.map(ym => {{
          let d;
          if (F.cat && F.store) d = {{v:0,u:0}};
          else if (F.cat) d = sumBCM(b, F.cat, [ym]);
          else if (F.store) d = sumSBM(F.store, b, [ym]);
          else d = sumBM(b, [ym]);
          return d.v || null;
        }});
        return {{
          label: b, data: vals,
          borderColor: DATA.brand_colors[b] || '#64748b',
          backgroundColor: (DATA.brand_colors[b] || '#64748b') + '20',
          tension: 0.3, pointRadius: 3, borderWidth: 2
        }};
      }})
    }},
    options: {{
      responsive: true, maintainAspectRatio: true,
      plugins: {{
        tooltip: {{
          mode: 'index',
          callbacks: {{
            label: c => {{
              const ym = evolMonths[c.dataIndex];
              const val = c.parsed.y || 0;
              const brand = c.dataset.label;
              const base = ` ${{brand}}: €${{fmt(val)}}`;
              if (!ym || !ym.startsWith('2026')) return base;
              const lyYm = '2025' + ym.slice(4);
              let lyV = 0;
              if (F.cat && !F.store) {{
                const d = DATA.bcm[brand+'|'+F.cat+'|'+lyYm]; if (d) lyV = d.v;
              }} else if (F.store && !F.cat) {{
                const d = DATA.sbm[F.store+'|'+brand+'|'+lyYm]; if (d) lyV = d.v;
              }} else if (!F.cat && !F.store) {{
                const d = DATA.bm[brand+'|'+lyYm]; if (d) lyV = d.v;
              }}
              if (!lyV) return base;
              const delta = val - lyV;
              const pct = delta / lyV * 100;
              const sign = delta >= 0 ? '+' : '';
              return ` ${{brand}}: €${{fmt(val)}} | vs LY ${{sign}}${{pct.toFixed(1)}}% (${{sign}}€${{fmt(Math.abs(delta))}})`;
            }}
          }}
        }}
      }},
      scales: {{ y: {{ ticks: {{ callback: v => '€'+fmt(v) }} }} }}
    }}
  }});
  // Actualizar estado visual dos botões
  ['all','2025','2026'].forEach(v => {{
    const btn = document.getElementById('bevol-'+v);
    if (btn) {{
      btn.classList.toggle('btn-evol-active', v === _brandEvolView);
    }}
  }});

  // Brand comparison bar
  const bRows = brands.map(b => {{
    let cy, ly;
    if (F.cat && F.store) {{ cy={{v:0,u:0}}; ly={{v:0,u:0}}; }}
    else if (F.cat) {{ cy = sumBCM(b,F.cat,pm.cy); ly = sumBCM(b,F.cat,pm.ly); }}
    else if (F.store) {{ cy = sumSBM(F.store,b,pm.cy); ly = sumSBM(F.store,b,pm.ly); }}
    else {{ cy = sumBM(b,pm.cy); ly = sumBM(b,pm.ly); }}
    const grow = ly.v > 0 ? (cy.v-ly.v)/ly.v*100 : null;
    return {{ brand:b, cy:cy.v, ly:ly.v, u:cy.u, grow }};
  }}).filter(r => r.cy > 0 || r.ly > 0).sort((a,b) => b.cy - a.cy);

  destroyChart('chartBrandCmp');
  CH['chartBrandCmp'] = new Chart(document.getElementById('chartBrandCmp'), {{
    type: 'bar',
    data: {{
      labels: bRows.map(r => r.brand),
      datasets: [
        {{ label: 'CY', data: bRows.map(r => r.cy), backgroundColor: bRows.map(r => DATA.brand_colors[r.brand] || '#64748b'), borderWidth: 0 }},
        {{ label: 'LY', data: bRows.map(r => r.ly), backgroundColor: bRows.map(r => (DATA.brand_colors[r.brand]||'#64748b')+'60'), borderWidth: 0 }},
      ]
    }},
    options: {{
      responsive: true,
      plugins: {{ tooltip: {{ callbacks: {{ label: c => ' €' + fmt(c.parsed.y) }} }} }},
      scales: {{ y: {{ ticks: {{ callback: v => '€'+fmt(v) }} }} }}
    }}
  }});

  destroyChart('chartBrandGrowth');
  CH['chartBrandGrowth'] = new Chart(document.getElementById('chartBrandGrowth'), {{
    type: 'bar',
    data: {{
      labels: bRows.map(r => r.brand),
      datasets: [{{
        label: 'Crescimento %',
        data: bRows.map(r => r.grow),
        backgroundColor: bRows.map(r => r.grow >= 0 ? '#10b98180' : '#ef444480'),
        borderColor:     bRows.map(r => r.grow >= 0 ? '#10b981' : '#ef4444'),
        borderWidth: 1.5
      }}]
    }},
    options: {{
      responsive: true,
      plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: c => ' ' + fmtPct(c.parsed.y) }} }} }},
      scales: {{ y: {{ ticks: {{ callback: v => v.toFixed(0)+'%' }} }} }}
    }}
  }});

  // Brand × cat table
  const bcRows = [];
  const bList = F.brand ? [F.brand] : DATA.brands;
  const cList = F.cat ? [F.cat] : DATA.cats;
  bList.forEach(b => cList.forEach(c => {{
    let cy, ly;
    if (F.store) {{
      // approximate
      cy = {{v:0,u:0}}; ly = {{v:0,u:0}};
      DATA.products.filter(p => p.marca===b && p.cat===c).forEach(p => {{
        const tops = DATA.store_top_cnps[F.store]||[];
        if (tops.includes(p.cnp)) {{
          const cd = sumProd(p,pm.cy); cy.v+=cd.v; cy.u+=cd.u;
          const ld = sumProd(p,pm.ly); ly.v+=ld.v; ly.u+=ld.u;
        }}
      }});
    }} else {{
      cy = sumBCM(b,c,pm.cy);
      ly = sumBCM(b,c,pm.ly);
    }}
    if (cy.v > 0 || ly.v > 0) {{
      const grow = ly.v > 0 ? (cy.v-ly.v)/ly.v*100 : null;
      bcRows.push({{ brand:b, cat:c, cy:cy.v, ly:ly.v, u:cy.u, grow }});
    }}
  }}));
  bcRows.sort((a,b) => b.cy - a.cy);

  document.getElementById('tblBrandCatBody').innerHTML = bcRows.map(r => `
    <tr>
      <td><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${{DATA.brand_colors[r.brand]||'#64748b'}};margin-right:6px"></span>${{r.brand}}</td>
      <td>${{r.cat}}</td>
      <td>${{fmtEur(r.cy)}}</td>
      <td>${{fmtEur(r.ly)}}</td>
      <td class="${{growthClass(r.grow)}}">${{fmtPct(r.grow)}}</td>
      <td>${{fmt(r.u)}}</td>
    </tr>`).join('');
}}

// ─── TAB: CATS ────────────────────────────────────────────────────────────────
function renderCats() {{
  const pm = getPeriodMonths();
  const cats = F.cat ? [F.cat] : DATA.cats;

  const cRows = cats.map(c => {{
    let cy, ly;
    if (F.brand && F.store) {{ cy={{v:0,u:0}}; ly={{v:0,u:0}}; }}
    else if (F.brand) {{ cy=sumBCM(F.brand,c,pm.cy); ly=sumBCM(F.brand,c,pm.ly); }}
    else if (F.store) {{ cy=sumSCM(F.store,c,pm.cy); ly=sumSCM(F.store,c,pm.ly); }}
    else {{ cy=sumCM(c,pm.cy); ly=sumCM(c,pm.ly); }}
    const grow = ly.v>0 ? (cy.v-ly.v)/ly.v*100 : null;
    return {{ cat:c, cy:cy.v, ly:ly.v, u:cy.u, grow }};
  }}).filter(r => r.cy>0||r.ly>0).sort((a,b)=>b.cy-a.cy);

  destroyChart('chartCatCmp');
  CH['chartCatCmp'] = new Chart(document.getElementById('chartCatCmp'), {{
    plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
    type: 'bar',
    data: {{
      labels: cRows.map(r=>r.cat),
      datasets: [
        {{ label:'CY', data:cRows.map(r=>r.cy), backgroundColor: cRows.map(r=>DATA.cat_colors[r.cat]||'#94a3b8'), borderWidth:0 }},
        {{ label:'LY', data:cRows.map(r=>r.ly), backgroundColor: cRows.map(r=>(DATA.cat_colors[r.cat]||'#94a3b8')+'60'), borderWidth:0 }},
      ]
    }},
    options: {{
      indexAxis:'y', responsive:true,
      plugins:{{
        tooltip:{{ callbacks:{{ label: c=>' €'+fmt(c.parsed.x) }} }},
        datalabels: {{
          display: true,
          anchor: 'center', align: 'center',
          formatter: (v, ctx) => {{
            if (!v || v < 100) return '';
            return v >= 10000 ? '€'+(v/1000).toFixed(0)+'k' : '€'+fmt(v,0);
          }},
          font: {{ size: 9, weight: 'bold' }},
          color: ctx => ctx.datasetIndex === 0 ? '#fff' : '#334155'
        }}
      }},
      scales:{{ x:{{ ticks:{{ callback: v=>'€'+fmt(v) }} }} }}
    }}
  }});

  // Monthly evol by cat
  const allMonths = [...DATA.ym_2025,...DATA.ym_2026];
  destroyChart('chartCatEvol');
  CH['chartCatEvol'] = new Chart(document.getElementById('chartCatEvol'), {{
    type: 'line',
    data: {{
      labels: allMonths.map(ym=>DATA.month_labels[ym]),
      datasets: cats.map(c => {{
        const vals = allMonths.map(ym => {{
          let d;
          if (F.brand) d=sumBCM(F.brand,c,[ym]);
          else if (F.store) d=sumSCM(F.store,c,[ym]);
          else d=sumCM(c,[ym]);
          return d.v||null;
        }});
        return {{ label:c, data:vals, borderColor:DATA.cat_colors[c]||'#94a3b8', tension:.3, pointRadius:2, borderWidth:2 }};
      }})
    }},
    options: {{
      responsive:true,
      plugins:{{ tooltip:{{ callbacks:{{ label: c=>` ${{c.dataset.label}}: €${{fmt(c.parsed.y)}}` }} }} }},
      scales:{{ y:{{ ticks:{{ callback: v=>'€'+fmt(v) }} }} }}
    }}
  }});

  destroyChart('chartCatGrowth');
  CH['chartCatGrowth'] = new Chart(document.getElementById('chartCatGrowth'), {{
    type: 'bar',
    data: {{
      labels: cRows.map(r=>r.cat),
      datasets: [{{
        label:'Crescimento %',
        data: cRows.map(r=>r.grow),
        backgroundColor: cRows.map(r=>r.grow>=0?'#10b98180':'#ef444480'),
        borderColor: cRows.map(r=>r.grow>=0?'#10b981':'#ef4444'),
        borderWidth:1.5
      }}]
    }},
    options: {{
      responsive:true, maintainAspectRatio:true,
      plugins:{{ legend:{{display:false}}, tooltip:{{ callbacks:{{ label: c=>' '+fmtPct(c.parsed.y) }} }} }},
      scales:{{ y:{{ ticks:{{ callback: v=>v.toFixed(0)+'%' }} }} }}
    }}
  }});

  // Cat × brand table
  const cbRows = [];
  const cList = F.cat ? [F.cat] : DATA.cats;
  const bList = F.brand ? [F.brand] : DATA.brands;
  cList.forEach(c => bList.forEach(b => {{
    let cy, ly;
    if (F.store) {{
      cy={{v:0,u:0}}; ly={{v:0,u:0}};
      DATA.products.filter(p=>p.marca===b&&p.cat===c).forEach(p=>{{
        const tops=DATA.store_top_cnps[F.store]||[];
        if (tops.includes(p.cnp)) {{
          const cd=sumProd(p,pm.cy); cy.v+=cd.v; cy.u+=cd.u;
          const ld=sumProd(p,pm.ly); ly.v+=ld.v; ly.u+=ld.u;
        }}
      }});
    }} else {{ cy=sumBCM(b,c,pm.cy); ly=sumBCM(b,c,pm.ly); }}
    if (cy.v>0||ly.v>0) {{
      const grow=ly.v>0?(cy.v-ly.v)/ly.v*100:null;
      cbRows.push({{cat:c,brand:b,cy:cy.v,ly:ly.v,u:cy.u,grow}});
    }}
  }}));
  cbRows.sort((a,b)=>b.cy-a.cy);
  document.getElementById('tblCatBrandBody').innerHTML = cbRows.map(r=>`
    <tr>
      <td><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${{DATA.cat_colors[r.cat]||'#94a3b8'}};margin-right:6px"></span>${{r.cat}}</td>
      <td>${{r.brand}}</td>
      <td>${{fmtEur(r.cy)}}</td>
      <td>${{fmtEur(r.ly)}}</td>
      <td class="${{growthClass(r.grow)}}">${{fmtPct(r.grow)}}</td>
      <td>${{fmt(r.u)}}</td>
    </tr>`).join('');
}}

// ─── TAB: PRODUCTS ────────────────────────────────────────────────────────────
let prodSort = {{ col: 5, dir: -1 }};
let prodPage = 1;
const PROD_PER_PAGE = 50;
let prodRows = [];

function renderProductsTable() {{
  const pm = getPeriodMonths();
  const searchQ = (document.getElementById('prodSearch').value || F.cnp || '').toLowerCase();
  let prods = getFilteredProducts();

  if (searchQ) {{
    prods = prods.filter(p =>
      p.cnp.toLowerCase().includes(searchQ) ||
      p.desc.toLowerCase().includes(searchQ) ||
      p.marca.toLowerCase().includes(searchQ) ||
      p.gama.toLowerCase().includes(searchQ)
    );
  }}

  prodRows = prods.map(p => {{
    const cy = sumProd(p, pm.cy);
    const ly = sumProd(p, pm.ly);
    const grow = ly.v > 0 ? (cy.v-ly.v)/ly.v*100 : null;
    return {{ ...p, cyV:cy.v, lyV:ly.v, cyU:cy.u, lyU:ly.u, grow }};
  }}).filter(r => r.cyV > 0 || r.lyV > 0);

  // sort
  prodRows.sort((a,b) => {{
    const vals = [r => r.cnp, r=>r.desc, r=>r.marca, r=>r.cat, r=>r.gama,
                  r=>r.cyV, r=>r.lyV, r=>r.grow||0, r=>r.cyU, r=>r.lyU];
    const fn = vals[prodSort.col] || (r=>r.cyV);
    const av=fn(a), bv=fn(b);
    if (typeof av==='string') return av.localeCompare(bv) * prodSort.dir;
    return ((av||0)-(bv||0)) * prodSort.dir;
  }});

  // KPIs
  const totCY = prodRows.reduce((s,r)=>s+r.cyV, 0);
  const totLY = prodRows.reduce((s,r)=>s+r.lyV, 0);
  const grow = totLY>0?(totCY-totLY)/totLY*100:null;
  document.getElementById('kpiProds').innerHTML = `
    <div class="kpi"><div class="label">Produtos Activos (CY)</div><div class="value">${{prodRows.filter(r=>r.cyV>0).length}}</div><div class="sub">de ${{prods.length}} no filtro</div></div>
    <div class="kpi"><div class="label">Vendas CY (€)</div><div class="value">${{fmtEur(totCY)}}</div><div class="sub">LY: ${{fmtEur(totLY)}} <span class="${{growthClass(grow)}}">${{fmtPct(grow)}}</span></div></div>
    <div class="kpi"><div class="label">Produtos com Crescimento</div><div class="value">${{prodRows.filter(r=>r.grow>0).length}}</div><div class="sub">vs ${{prodRows.filter(r=>r.grow<0).length}} em queda</div></div>
  `;

  // Pagination
  const totalPages = Math.ceil(prodRows.length / PROD_PER_PAGE);
  if (prodPage > totalPages) prodPage = 1;
  const start = (prodPage-1)*PROD_PER_PAGE;
  const pageRows = prodRows.slice(start, start+PROD_PER_PAGE);

  document.getElementById('tblProdsBody').innerHTML = pageRows.map(r=>`
    <tr>
      <td>${{r.cnp}}</td>
      <td title="${{r.desc}}">${{r.desc.length>40?r.desc.slice(0,40)+'…':r.desc}}</td>
      <td><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:${{DATA.brand_colors[r.marca]||'#64748b'}};margin-right:4px"></span>${{r.marca}}</td>
      <td>${{r.cat}}</td>
      <td>${{r.gama}}</td>
      <td>${{fmtEur(r.cyV)}}</td>
      <td>${{fmtEur(r.lyV)}}</td>
      <td class="${{growthClass(r.grow)}}">${{fmtPct(r.grow)}}</td>
      <td>${{fmt(r.cyU)}}</td>
      <td>${{fmt(r.lyU)}}</td>
    </tr>`).join('');

  // Pagination controls
  const pag = document.getElementById('prodPagination');
  if (totalPages <= 1) {{ pag.innerHTML=''; return; }}
  const pages = [];
  pages.push(`<span class="page-info">Pág. ${{prodPage}} de ${{totalPages}} (${{prodRows.length}} produtos)</span>`);
  if (prodPage>1) pages.push(`<button onclick="goPage(${{prodPage-1}})">‹</button>`);
  const start2 = Math.max(1, prodPage-2), end2 = Math.min(totalPages, prodPage+2);
  for (let i=start2;i<=end2;i++) pages.push(`<button class="${{i===prodPage?'active':''}}" onclick="goPage(${{i}})">${{i}}</button>`);
  if (prodPage<totalPages) pages.push(`<button onclick="goPage(${{prodPage+1}})">›</button>`);
  pag.innerHTML = pages.join('');
}}

function goPage(p) {{ prodPage=p; renderProductsTable(); }}
function sortProdTable(col) {{
  if (prodSort.col===col) prodSort.dir *= -1; else {{ prodSort.col=col; prodSort.dir=-1; }}
  prodPage=1;
  renderProductsTable();
}}

// ─── TAB: GROWTH ─────────────────────────────────────────────────────────────
function renderGrowth() {{
  const pm = getPeriodMonths();
  const rows = getFilteredProducts().map(p => {{
    const cy=sumProd(p,pm.cy), ly=sumProd(p,pm.ly);
    const delta = cy.v - ly.v;
    const pct = ly.v>0 ? delta/ly.v*100 : null;
    return {{ ...p, cyV:cy.v, lyV:ly.v, delta, pct }};
  }}).filter(r => r.lyV > 0 || r.cyV > 0);

  const byDelta = [...rows].sort((a,b)=>b.delta-a.delta);
  const byPct   = [...rows].filter(r=>r.lyV>0).sort((a,b)=>b.pct-a.pct);

  const shortDesc = d => d.length>30?d.slice(0,30)+'…':d;

  function makeGrowChart(id, data, field, isNeg) {{
    destroyChart(id);
    CH[id] = new Chart(document.getElementById(id), {{
      plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
      type: 'bar',
      data: {{
        labels: data.map(r=>shortDesc(r.desc)),
        datasets: [{{
          label: field==='delta' ? 'Crescimento € vs LY' : 'Crescimento % vs LY',
          data: data.map(r=>r[field]),
          backgroundColor: data.map(r=>r[field]>=0?'#10b98180':'#ef444480'),
          borderColor:     data.map(r=>r[field]>=0?'#10b981':'#ef4444'),
          borderWidth:1.5
        }}]
      }},
      options: {{
        indexAxis:'y', responsive:true,
        plugins:{{
          legend:{{display:false}},
          tooltip:{{ callbacks:{{ label: c => {{
            const d = data[c.dataIndex];
            if (field==='delta') return [' €'+fmt(c.parsed.x), ' vs LY: '+fmtPct(d.pct)];
            return [' '+fmtPct(c.parsed.x), ' vs LY: €'+fmt(d.delta)];
          }} }} }},
          datalabels: {{
            display: true,
            anchor: 'center', align: 'center',
            formatter: v => v===null?'': field==='delta'?(Math.abs(v)>=10000?'€'+(v/1000).toFixed(0)+'k':'€'+fmt(v,0)):fmtPct(v),
            color: '#fff',
            font: {{ size: 9, weight: 'bold' }},
            textStrokeColor: '#00000060',
            textStrokeWidth: 2
          }}
        }},
        scales:{{ x:{{ ticks:{{ callback: v=>field==='delta'?'€'+fmt(v):v.toFixed(0)+'%' }} }} }}
      }}
    }});
  }}

  makeGrowChart('chartGrowthTop',    byDelta.slice(0,15).reverse(),     'delta', false);
  makeGrowChart('chartGrowthBot',    byDelta.slice(-15),                 'delta', true);
  makeGrowChart('chartGrowthTopPct', byPct.slice(0,15).reverse(),        'pct',   false);
  makeGrowChart('chartGrowthBotPct', byPct.slice(-15),                   'pct',   true);
}}

// ─── TAB: INNOV ──────────────────────────────────────────────────────────────
function renderInnov() {{
  const pm = getPeriodMonths();
  let innovProds = DATA.products.filter(p => p.innov);
  if (F.brand) innovProds = innovProds.filter(p=>p.marca===F.brand);
  if (F.cat)   innovProds = innovProds.filter(p=>p.cat===F.cat);

  const rows = innovProds.map(p => {{
    const cy2026 = sumProd(p, DATA.ym_2026);
    return {{ ...p, v2026:cy2026.v, u2026:cy2026.u }};
  }}).filter(r=>r.v2026>0).sort((a,b)=>b.v2026-a.v2026);

  // Total 2026 (all products)
  const total2026 = getFilteredTotal(DATA.ym_2026).v;
  const innovTotal = rows.reduce((s,r)=>s+r.v2026, 0);
  const innovPct = total2026 > 0 ? (innovTotal/total2026*100) : 0;

  document.getElementById('innovPct').textContent = innovPct.toFixed(1)+'%';

  document.getElementById('kpiInnov').innerHTML = `
    <div class="kpi"><div class="label">Produtos Inovação</div><div class="value">${{rows.length}}</div><div class="sub">sem histórico 2025</div></div>
    <div class="kpi"><div class="label">Vendas 2026 Inovação (€)</div><div class="value">${{fmtEur(innovTotal)}}</div><div class="sub">de ${{fmtEur(total2026)}} total</div></div>
    <div class="kpi"><div class="label">Unidades Inovação 2026</div><div class="value">${{fmt(rows.reduce((s,r)=>s+r.u2026,0))}}</div><div class="sub">unidades vendidas</div></div>
  `;

  destroyChart('chartInnov');
  const top20 = rows.slice(0,20);
  const innovSum = top20.reduce((s,r)=>s+r.v2026, 0);
  CH['chartInnov'] = new Chart(document.getElementById('chartInnov'), {{
    plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
    type: 'bar',
    data: {{
      labels: top20.map(r=>r.desc.length>35?r.desc.slice(0,35)+'…':r.desc),
      datasets: [{{
        label:'Vendas 2026 (€)',
        data: top20.map(r=>r.v2026),
        backgroundColor: top20.map(r=>DATA.brand_colors[r.marca]||'#64748b'),
        borderWidth:0
      }}]
    }},
    options: {{
      indexAxis:'y', responsive:true,
      plugins:{{
        legend:{{display:false}},
        tooltip:{{ callbacks:{{ label: c=>' €'+fmt(c.parsed.x)+' ('+(innovSum>0?(c.parsed.x/innovSum*100).toFixed(1):'0')+'%)' }} }},
        datalabels: {{
          display: true,
          anchor: 'center', align: 'center',
          formatter: v => {{
            if (!v || v < 50) return '';
            const eur = v>=10000?'€'+(v/1000).toFixed(0)+'k':'€'+fmt(v,0);
            const pct = innovSum>0?' ('+(v/innovSum*100).toFixed(1)+'%)':'';
            return eur+pct;
          }},
          color: '#fff',
          font: {{ size: 9, weight: 'bold' }},
          textStrokeColor: '#00000060',
          textStrokeWidth: 2
        }}
      }},
      scales:{{ x:{{ ticks:{{ callback: v=>'€'+fmt(v) }} }} }}
    }}
  }});

  document.getElementById('tblInnovBody').innerHTML = rows.map(r=>`
    <tr>
      <td>${{r.cnp}}</td>
      <td title="${{r.desc}}">${{r.desc.length>45?r.desc.slice(0,45)+'…':r.desc}} <span class="badge-innov">NOVO</span></td>
      <td>${{r.marca}}</td>
      <td>${{r.cat}}</td>
      <td>${{fmtEur(r.v2026)}}</td>
      <td>${{fmt(r.u2026)}}</td>
    </tr>`).join('');
}}

// ─── TAB: ALERTS ─────────────────────────────────────────────────────────────
function renderAlerts() {{
  const pm = getPeriodMonths();
  if (pm.ly.length === 0) {{
    ['alertProdUp','alertProdDn','alertStoreUp','alertStoreDn'].forEach(id => {{
      document.getElementById(id).innerHTML = '<div class="ac-row"><div class="ac-name" style="color:var(--muted)">Seleccione um período com comparativo (ex: YTD 2026)</div></div>';
    }});
    return;
  }}

  const prodRows = getFilteredProducts().map(p => {{
    const cy=sumProd(p,pm.cy), ly=sumProd(p,pm.ly);
    return {{ name:p.desc, delta:cy.v-ly.v, cy:cy.v, ly:ly.v }};
  }}).filter(r=>r.ly>0||r.cy>0).sort((a,b)=>b.delta-a.delta);

  function alertHtml(rows, cls) {{
    return rows.map((r,i)=>`
      <div class="ac-row">
        <div class="ac-rank ${{cls}}">${{i+1}}</div>
        <div class="ac-name" title="${{r.name}}">${{r.name.length>50?r.name.slice(0,50)+'…':r.name}}</div>
        <div class="ac-val ${{cls==='grow'?'badge-grow':'badge-drop'}}">${{r.delta>=0?'+':''}}${{fmtEur(r.delta)}}</div>
      </div>`).join('');
  }}

  const upProds = prodRows.filter(r=>r.delta>0);
  const dnProds = [...prodRows].filter(r=>r.delta<0).sort((a,b)=>a.delta-b.delta);
  document.getElementById('alertProdUp').innerHTML = alertHtml(upProds.slice(0,10),'grow');
  document.getElementById('alertProdDn').innerHTML = alertHtml(dnProds.slice(0,10),'drop');

  const storeRows = DATA.stores.map(s => {{
    let cy, ly;
    if (F.brand) {{ cy=sumSBM(s,F.brand,pm.cy); ly=sumSBM(s,F.brand,pm.ly); }}
    else if (F.cat) {{ cy=sumSCM(s,F.cat,pm.cy); ly=sumSCM(s,F.cat,pm.ly); }}
    else {{ cy=sumSM(s,pm.cy); ly=sumSM(s,pm.ly); }}
    return {{ name:s, delta:cy.v-ly.v, cy:cy.v, ly:ly.v }};
  }}).filter(r=>r.ly>0||r.cy>0).sort((a,b)=>b.delta-a.delta);

  const upStores = storeRows.filter(r=>r.delta>0);
  const dnStores = [...storeRows].filter(r=>r.delta<0).sort((a,b)=>a.delta-b.delta);
  document.getElementById('alertStoreUp').innerHTML = alertHtml(upStores.slice(0,10),'grow');
  document.getElementById('alertStoreDn').innerHTML = alertHtml(dnStores.slice(0,10),'drop');
}}

// ─── SORT TABLE ───────────────────────────────────────────────────────────────
const _sortState = {{}};
function sortTable(id, col) {{
  const st = _sortState[id] || {{ col:-1, dir:1 }};
  const dir = st.col===col ? -st.dir : -1;
  _sortState[id] = {{ col, dir }};
  const tbody = document.getElementById(id).querySelector('tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));
  rows.sort((a,b) => {{
    const av = a.cells[col]?.textContent.trim()||'';
    const bv = b.cells[col]?.textContent.trim()||'';
    const an = parseFloat(av.replace(/[€%+\s.,]/g,'').replace(',','.'));
    const bn = parseFloat(bv.replace(/[€%+\s.,]/g,'').replace(',','.'));
    if (!isNaN(an)&&!isNaN(bn)) return (an-bn)*dir;
    return av.localeCompare(bv)*dir;
  }});
  rows.forEach(r => tbody.appendChild(r));
}}

// ─── INIT ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {{
  initFilters();
  applyFilters();
}});
</script>
</body>
</html>"""

# ─── 6. WRITE OUTPUT ──────────────────────────────────────────────────────────
html_out = HTML.replace('{data_json}', data_json)
with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write(html_out)
print(f"\n✅ Dashboard gerado: {OUTPUT} ({Path(OUTPUT).stat().st_size/1024:.0f} KB)")
