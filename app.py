import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import base64, io, os

# ─── PAGE CONFIG ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Proyeksi Anggaran ATR/BPN",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── KONSTANTA ───────────────────────────────────────────────────────────────
BULAN = ['Februari','Maret','April','Mei','Juni','Juli',
         'Agustus','September','Oktober','November','Desember']
SEKSI_LIST = ['S1','S2','S3','S4','S5','S6']
SEKSI_NAMA = {
    'S1': 'Survei & Pemetaan',
    'S2': 'Penetapan Hak & Pendaftaran',
    'S3': 'Penataan & Pemberdayaan',
    'S4': 'Pengadaan Tanah & Pengembangan',
    'S5': 'Pengendalian & Sengketa',
    'S6': 'Sub Bag Tata Usaha'
}
COLORS = ['#1a3a6b','#00b894','#e17055','#c8a84b','#9b59b6','#2e86de']
SEKSI_COLORS = dict(zip(SEKSI_LIST, COLORS))

# Proyeksi 2026
PROYEKSI_2026 = {
    'S1': [9.54,9.14,8.76,8.38,8.00,7.64,7.28,6.94,6.60,6.27,5.95],
    'S2': [10.04,9.69,9.34,8.99,8.64,8.29,7.94,7.59,7.24,6.89,6.54],
    'S3': [7.66,7.35,7.04,6.74,6.45,6.17,5.89,5.61,5.35,5.09,4.83],
    'S4': [10.50,10.07,9.64,9.21,8.78,8.35,7.93,7.50,7.07,6.64,6.21],
    'S5': [9.99,9.66,9.34,9.02,8.70,8.37,8.05,7.73,7.40,7.08,6.76],
    'S6': [11.18,10.74,10.30,9.86,9.42,8.98,8.54,8.10,7.65,7.21,6.77],
}
# Proyeksi 2027
PROYEKSI_2027 = {
    'S1': [9.63,9.30,8.97,8.64,8.31,7.97,7.64,7.31,6.98,6.65,6.32],
    'S2': [10.03,9.68,9.33,8.98,8.63,8.28,7.93,7.58,7.23,6.88,6.53],
    'S3': [8.02,7.75,7.47,7.19,6.92,6.64,6.37,6.09,5.81,5.54,5.26],
    'S4': [10.50,10.07,9.64,9.21,8.78,8.35,7.93,7.50,7.07,6.64,6.21],
    'S5': [9.99,9.66,9.34,9.02,8.69,8.37,8.05,7.73,7.40,7.08,6.76],
    'S6': [11.18,10.74,10.30,9.86,9.42,8.98,8.54,8.10,7.65,7.21,6.77],
}
# Tren historis
TREN = {
    'S1': {2018:81.79,2019:81.07,2020:84.58,2021:96.08,2022:93.88,2023:93.88,2024:66.75,2025:98.25},
    'S2': {2018:85.00,2019:88.00,2020:86.41,2021:78.99,2022:88.88,2023:98.82,2024:99.94,2025:99.99},
    'S3': {2018:71.53,2019:71.09,2020:55.48,2021:24.34,2022:68.63,2023:95.49,2024:100.00,2025:98.74},
    'S4': {2018:98.00,2019:37.26,2020:100.00,2021:100.00,2022:99.95,2023:100.00,2024:100.00,2025:100.00},
    'S5': {2018:78.61,2019:75.71,2020:89.78,2021:99.81,2022:99.98,2023:99.94,2024:93.44,2025:99.25},
    'S6': {2018:96.60,2019:98.22,2020:98.29,2021:98.95,2022:98.90,2023:99.73,2024:99.39,2025:99.98},
}
# Regresi
REGRESI = {
    'S1': {'a':3.154,'b':-0.065,'transformasi':True,'r2':0.072,'normal':False,'p_value':0.009,'mae':2.53,'risiko':'Sedang'},
    'S2': {'a':10.39,'b':-0.35,'transformasi':False,'r2':0.083,'normal':True,'p_value':0.063,'mae':0.44,'risiko':'Rendah'},
    'S3': {'a':2.825,'b':-0.057,'transformasi':True,'r2':0.062,'normal':False,'p_value':0.011,'mae':4.99,'risiko':'Tinggi'},
    'S4': {'a':10.928,'b':-0.429,'transformasi':False,'r2':0.074,'normal':True,'p_value':0.058,'mae':0.01,'risiko':'Sedang'},
    'S5': {'a':10.31,'b':-0.323,'transformasi':False,'r2':0.073,'normal':True,'p_value':0.061,'mae':0.03,'risiko':'Rendah'},
    'S6': {'a':11.623,'b':-0.441,'transformasi':False,'r2':0.111,'normal':True,'p_value':0.072,'mae':0.01,'risiko':'Rendah'},
}

# ─── LOAD DATA EXCEL (di-embed) ──────────────────────────────────────────────
@st.cache_data
def load_data():
    b64_path = os.path.join(os.path.dirname(__file__), 'data_b64.txt')
    with open(b64_path, 'r') as f:
        b64 = f.read().strip()
    xls_bytes = base64.b64decode(b64)
    df = pd.read_excel(io.BytesIO(xls_bytes), sheet_name='DATA GABUNGAN LENGKAP', header=2)
    df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]
    df = df[df['Seksi'].isin(SEKSI_LIST)]
    df = df[df['Bulan'].isin(BULAN)]
    df['Tahun'] = df['Tahun'].astype(int)
    int_cols = ['Pagu PNBP (Rp)','Pagu RM PTSL (Rp)','Pagu RM Non-PTSL (Rp)',
                'Realisasi PNBP (Rp)','Realisasi RM PTSL (Rp)','Realisasi RM Non-PTSL (Rp)']
    for c in int_cols:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(int)
    df['TotalPagu'] = df['Pagu PNBP (Rp)'] + df['Pagu RM PTSL (Rp)'] + df['Pagu RM Non-PTSL (Rp)']
    df['TotalReal'] = df['Realisasi PNBP (Rp)'] + df['Realisasi RM PTSL (Rp)'] + df['Realisasi RM Non-PTSL (Rp)']
    return df

# ─── HELPER FUNCTIONS ────────────────────────────────────────────────────────
def get_pagu(df, seksi, tahun):
    sub = df[(df['Seksi']==seksi) & (df['Tahun']==tahun)]
    if sub.empty:
        return 0
    # pagu is same for all months, take first
    row = sub.iloc[0]
    return int(row['Pagu PNBP (Rp)'] + row['Pagu RM PTSL (Rp)'] + row['Pagu RM Non-PTSL (Rp)'])

def get_pct_bulanan(df, seksi, tahun):
    """% realisasi per bulan (non-kumulatif) — dihitung dari selisih kumulatif antar bulan."""
    pagu = get_pagu(df, seksi, tahun)
    if pagu == 0:
        return [0]*11
    kum_rp = []
    for bln in BULAN:
        sub = df[(df['Seksi']==seksi) & (df['Tahun']==tahun) & (df['Bulan']==bln)]
        kum_rp.append(int(sub.iloc[0]['TotalReal']) if not sub.empty else 0)
    # Non-kumulatif: selisih antar bulan
    hasil = []
    prev = 0
    for val in kum_rp:
        if val > 0:
            delta = val - prev
            hasil.append(round(delta / pagu * 100, 4))
            prev = val
        else:
            hasil.append(0)
    return hasil

def get_kumulatif(pct):
    kum, total = [], 0
    for p in pct:
        total += p
        kum.append(round(total, 4))
    return kum

def get_delta(pct):
    delta = [pct[0]]
    for i in range(1, len(pct)):
        delta.append(round(pct[i] - pct[i-1], 4))
    return delta

def get_sisa(df, seksi, tahun, sampai_idx):
    """TotalReal adalah kumulatif — ambil nilai bulan terakhir yang ada s/d sampai_idx."""
    pagu = get_pagu(df, seksi, tahun)
    total_real = 0
    # cari bulan terakhir yang ada data dari 0..sampai_idx
    for i in range(sampai_idx, -1, -1):
        bln = BULAN[i]
        sub = df[(df['Seksi']==seksi) & (df['Tahun']==tahun) & (df['Bulan']==bln)]
        if not sub.empty:
            total_real = int(sub.iloc[0]['TotalReal'])
            break
    sisa = pagu - total_real
    pct_real = round(total_real / pagu * 100, 2) if pagu > 0 else 0
    sisa_pct = round(sisa / pagu * 100, 2) if pagu > 0 else 0
    return {
        'total_pagu': pagu,
        'realisasi_rp': total_real,
        'sisa_rp': sisa,
        'realisasi_pct': pct_real,
        'sisa_pct': sisa_pct,
    }

def get_dashboard_summary(df, tahun):
    per_seksi = {}
    total_pagu = 0
    total_real = 0
    for s in SEKSI_LIST:
        pagu = get_pagu(df, s, tahun)
        # TotalReal adalah kumulatif — ambil bulan terakhir yang ada datanya
        sub = df[(df['Seksi']==s) & (df['Tahun']==tahun)].copy()
        sub['_bulan_idx'] = sub['Bulan'].apply(lambda b: BULAN.index(b) if b in BULAN else -1)
        sub = sub[sub['_bulan_idx'] >= 0].sort_values('_bulan_idx')
        real_rp = int(sub.iloc[-1]['TotalReal']) if not sub.empty else 0
        pct = round(real_rp / pagu * 100, 2) if pagu > 0 else 0
        per_seksi[s] = {
            'nama': SEKSI_NAMA[s],
            'pagu': pagu,
            'realisasi_rp': real_rp,
            'realisasi_pct': pct,
        }
        total_pagu += pagu
        total_real += real_rp
    pct_total = round(total_real / total_pagu * 100, 2) if total_pagu > 0 else 0
    return {
        'total_pagu': total_pagu,
        'total_realisasi': total_real,
        'pct_realisasi': pct_total,
        'sisa_pagu': total_pagu - total_real,
        'sisa_pct': round((total_pagu - total_real) / total_pagu * 100, 2) if total_pagu > 0 else 0,
        'per_seksi': per_seksi,
        'tahun': tahun,
    }

def badge_color(pct):
    if pct >= 80:
        return "🟢"
    elif pct >= 50:
        return "🟡"
    return "🔴"

def risiko_color(risiko):
    return {'Rendah':'#00b894','Sedang':'#fdcb6e','Tinggi':'#d63031'}.get(risiko,'#636e72')

# ─── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
:root {
  --primary: #1a3a6b;
  --secondary: #c8a84b;
  --accent: #2e86de;
  --bg: #f0f4f8;
  --success: #00b894;
  --danger: #d63031;
  --warning: #fdcb6e;
  --muted: #636e72;
}
[data-testid="stSidebar"] {
  background: var(--primary) !important;
}
[data-testid="stSidebar"] * { color: rgba(255,255,255,0.85) !important; }
[data-testid="stSidebar"] .stSelectbox label { color: rgba(255,255,255,0.6) !important; font-size:11px !important; }
[data-testid="stSidebarNav"] a { color: rgba(255,255,255,0.75) !important; }

.stat-card {
  border-radius: 16px; padding: 22px 24px;
  box-shadow: 0 4px 20px rgba(0,0,0,0.08);
  color: white; position: relative; overflow: hidden;
  transition: transform .2s;
}
.stat-card:hover { transform: translateY(-3px); }
.stat-icon { font-size: 28px; margin-bottom: 10px; }
.stat-value { font-size: 22px; font-weight: 800; }
.stat-label { font-size: 12px; opacity: 0.8; margin-top: 2px; }
.stat-sub { font-size: 11px; opacity: 0.7; margin-top: 6px; }

.chart-card {
  background: white; border-radius: 16px; padding: 24px;
  box-shadow: 0 4px 20px rgba(0,0,0,0.06); margin-bottom: 16px;
}
.chart-title {
  font-weight: 700; color: var(--primary); font-size: 14px;
  margin-bottom: 16px; display: flex; align-items: center; gap: 8px;
}
.section-header {
  font-weight: 800; font-size: 16px; color: var(--primary);
  margin-bottom: 16px; padding-bottom: 8px;
  border-bottom: 2px solid var(--secondary);
}
.badge-success { background:#d4edda; color:#155724; padding:3px 10px;
  border-radius:50px; font-size:11px; font-weight:700; }
.badge-warning { background:#fff3cd; color:#856404; padding:3px 10px;
  border-radius:50px; font-size:11px; font-weight:700; }
.badge-danger { background:#f8d7da; color:#721c24; padding:3px 10px;
  border-radius:50px; font-size:11px; font-weight:700; }
.progress-wrap { background:#f0f4f8; border-radius:50px; height:8px; overflow:hidden; }
.stMetric { background:white; border-radius:12px; padding:16px !important;
  box-shadow:0 2px 12px rgba(0,0,0,0.06); }
</style>
""", unsafe_allow_html=True)

# ─── LOAD ────────────────────────────────────────────────────────────────────
df = load_data()

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding:16px 0 12px; border-bottom:1px solid rgba(255,255,255,0.15); margin-bottom:8px'>
      <div style='font-size:42px'>🏛️</div>
      <div style='font-weight:700; font-size:13px; color:#c8a84b; margin-top:6px'>ATR/BPN SURABAYA I</div>
      <div style='font-size:11px; opacity:0.6'>Sistem Proyeksi Anggaran</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<div style='font-size:10px;text-transform:uppercase;letter-spacing:1.5px;opacity:0.45;padding:12px 0 4px;font-weight:600'>Menu Utama</div>", unsafe_allow_html=True)
    halaman = st.radio("", ["🏠 Dashboard", "📊 Monitoring Bulanan", "📈 Analisis & Proyeksi"],
                       label_visibility="collapsed")

    st.markdown("---")
    st.markdown("<div style='font-size:11px;opacity:0.5;text-align:center'>© 2026 ATR/BPN Surabaya I</div>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# HALAMAN 1 — DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════
if halaman == "🏠 Dashboard":
    st.markdown("<h4 style='color:#1a3a6b;font-weight:800;margin-bottom:24px'>🏠 Dashboard — Ringkasan Anggaran</h4>", unsafe_allow_html=True)

    tahun_ref = 2025
    summary = get_dashboard_summary(df, tahun_ref)

    # ── Stat Cards ──
    c1,c2,c3,c4 = st.columns(4)
    cards = [
        (c1, "💼", f"Rp {summary['total_pagu']:,.0f}", f"Total Pagu {tahun_ref}", "6 Seksi", "linear-gradient(135deg,#1a3a6b,#2e6da4)"),
        (c2, "✅", f"Rp {summary['total_realisasi']:,.0f}", "Total Realisasi", f"{summary['pct_realisasi']}% tercapai", "linear-gradient(135deg,#00b894,#00cec9)"),
        (c3, "⏳", f"Rp {summary['sisa_pagu']:,.0f}", "Sisa Pagu", f"{summary['sisa_pct']}% belum terserap", "linear-gradient(135deg,#e17055,#d63031)"),
        (c4, "📊", f"{summary['pct_realisasi']}%", "% Realisasi Keseluruhan", f"Tahun {tahun_ref}", "linear-gradient(135deg,#c8a84b,#f9ca24)"),
    ]
    for col, icon, val, lbl, sub, grad in cards:
        col.markdown(f"""
        <div class="stat-card" style="background:{grad}">
          <div class="stat-icon">{icon}</div>
          <div class="stat-value">{val}</div>
          <div class="stat-label">{lbl}</div>
          <div class="stat-sub">{sub}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Gauge + Bar ──
    col1, col2 = st.columns([4,6])
    with col1:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">🎯 Capaian Realisasi Keseluruhan</div>', unsafe_allow_html=True)
        pct = summary['pct_realisasi']
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=pct,
            delta={'reference':85, 'valueformat':'.1f'},
            number={'suffix':'%', 'font':{'size':40,'color':'#1a3a6b'}},
            gauge={
                'axis':{'range':[0,100],'tickcolor':'#636e72'},
                'bar':{'color':'#1a3a6b','thickness':0.25},
                'steps':[
                    {'range':[0,50],'color':'#ffeaea'},
                    {'range':[50,75],'color':'#fff3cd'},
                    {'range':[75,100],'color':'#d4edda'},
                ],
                'threshold':{'line':{'color':'#c8a84b','width':4},'thickness':0.75,'value':85}
            }
        ))
        fig_gauge.update_layout(margin=dict(t=20,b=20,l=20,r=20),
                                 paper_bgcolor='rgba(0,0,0,0)',
                                 plot_bgcolor='rgba(0,0,0,0)',
                                 height=280)
        st.plotly_chart(fig_gauge, use_container_width=True, config={'displayModeBar':False})
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">📊 Realisasi per Seksi</div>', unsafe_allow_html=True)
        seksi_names = [summary['per_seksi'][s]['nama'] for s in SEKSI_LIST]
        seksi_pcts  = [summary['per_seksi'][s]['realisasi_pct'] for s in SEKSI_LIST]
        bar_colors  = ['#00b894' if v>=80 else '#fdcb6e' if v>=50 else '#d63031' for v in seksi_pcts]
        fig_bar = go.Figure(go.Bar(
            x=seksi_names, y=seksi_pcts,
            marker=dict(color=bar_colors),
            text=[f"{v}%" for v in seksi_pcts], textposition='outside',
            hovertemplate='<b>%{x}</b><br>Realisasi: %{y}%<extra></extra>'
        ))
        fig_bar.add_shape(type='line', x0=-0.5, x1=5.5, y0=85, y1=85,
                          line=dict(color='#c8a84b', width=2, dash='dot'))
        fig_bar.add_annotation(x=5.5, y=85, text='Target 85%', showarrow=False,
                                font=dict(color='#c8a84b', size=11), xanchor='right')
        fig_bar.update_layout(margin=dict(t=20,b=60,l=40,r=20),
                               paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                               yaxis=dict(range=[0,115], title='%', gridcolor='#f0f0f0'),
                               xaxis=dict(tickangle=-15, tickfont=dict(size=10)),
                               height=280, font=dict(family='Segoe UI', color='#2d3436'))
        st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar':False})
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Tabel per Seksi ──
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    st.markdown(f'<div class="chart-title">📋 Ringkasan Anggaran per Seksi — {tahun_ref}</div>', unsafe_allow_html=True)
    rows = []
    for s in SEKSI_LIST:
        d = summary['per_seksi'][s]
        pct_v = d['realisasi_pct']
        badge = "🟢 " if pct_v>=80 else "🟡 " if pct_v>=50 else "🔴 "
        rows.append({
            'Seksi': f"{s} — {d['nama']}",
            'Total Pagu (Rp)': f"Rp {d['pagu']:,.0f}",
            'Realisasi (Rp)': f"Rp {d['realisasi_rp']:,.0f}",
            '% Realisasi': f"{badge}{pct_v}%",
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Akses Cepat ──
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    st.markdown('<div class="chart-title">🚀 Akses Cepat</div>', unsafe_allow_html=True)
    qa1, qa2 = st.columns(2)
    with qa1:
        st.info("📊 **Monitoring Bulanan** — Pantau sisa pagu & pertumbuhan per seksi. Pilih menu di sidebar.")
    with qa2:
        st.info("📈 **Analisis & Proyeksi** — Lihat proyeksi 2026–2027 & analisis risiko. Pilih menu di sidebar.")
    st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# HALAMAN 2 — MONITORING BULANAN
# ═══════════════════════════════════════════════════════════════════════════════
elif halaman == "📊 Monitoring Bulanan":
    st.markdown("<h4 style='color:#1a3a6b;font-weight:800;margin-bottom:8px'>📊 Monitoring Realisasi Bulanan</h4>", unsafe_allow_html=True)

    # ── Filter ──
    fc1, fc2, fc3 = st.columns([1,1,4])
    with fc1:
        tahun_sel = st.selectbox("Tahun", list(range(2018,2026)), index=7)
    with fc2:
        bulan_sel = st.selectbox("s/d Bulan", BULAN, index=10)
    bulan_idx = BULAN.index(bulan_sel)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Summary Cards per Seksi ──
    cols = st.columns(3)
    card_colors = ['#2196F3','#4CAF50','#FF9800','#E91E63','#9C27B0','#00BCD4']
    hasil_monitoring = {}
    for s in SEKSI_LIST:
        sisa = get_sisa(df, s, tahun_sel, bulan_idx)
        pct  = get_pct_bulanan(df, s, tahun_sel)
        hasil_monitoring[s] = {
            'nama': SEKSI_NAMA[s],
            'sisa': sisa,
            'pct_bulanan': pct,
            'kumulatif': get_kumulatif(pct),
            'delta': get_delta(pct),
        }

    for i, s in enumerate(SEKSI_LIST):
        d = hasil_monitoring[s]
        c = card_colors[i]
        sisa = d['sisa']
        pct_v = sisa['realisasi_pct']
        badge_cls = "badge-success" if pct_v>=80 else "badge-warning" if pct_v>=50 else "badge-danger"
        pct_bar = min(pct_v, 100)
        with cols[i % 3]:
            st.markdown(f"""
            <div style='background:white;border-radius:14px;padding:18px 20px;
                        box-shadow:0 4px 14px rgba(0,0,0,0.06);
                        border-top:4px solid {c};margin-bottom:16px'>
              <div style='display:flex;justify-content:space-between;align-items:flex-start'>
                <div>
                  <div style='font-weight:800;font-size:14px;color:#1a3a6b'>{s} — {d['nama']}</div>
                  <div style='font-size:11px;color:#636e72;margin-top:2px'>
                    Pagu: <strong>Rp {sisa['total_pagu']:,.0f}</strong>
                  </div>
                </div>
                <span class="{badge_cls}">{pct_v}% terserap</span>
              </div>
              <div style='background:#f0f4f8;border-radius:50px;height:8px;overflow:hidden;margin:12px 0 8px'>
                <div style='width:{pct_bar}%;height:100%;background:{c};border-radius:50px'></div>
              </div>
              <div style='display:flex;justify-content:space-between'>
                <div>
                  <div style='font-size:11px;color:#636e72'>Realisasi s/d {bulan_sel}</div>
                  <div style='font-size:13px;font-weight:700;color:#00b894'>Rp {sisa['realisasi_rp']:,.0f}</div>
                </div>
                <div style='text-align:right'>
                  <div style='font-size:11px;color:#636e72'>Sisa Pagu</div>
                  <div style='font-size:13px;font-weight:700;color:#d63031'>Rp {sisa['sisa_rp']:,.0f}</div>
                  <div style='font-size:11px;color:#d63031'>{sisa['sisa_pct']}% sisa</div>
                </div>
              </div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Charts Kumulatif & Delta ──
    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">📈 Kumulatif % Realisasi per Bulan</div>', unsafe_allow_html=True)
        sel_kum = st.selectbox("Pilih Seksi", SEKSI_LIST, key='kum_seksi',
                               format_func=lambda x: f"{x} — {SEKSI_NAMA[x]}")
        d = hasil_monitoring[sel_kum]
        fig_kum = go.Figure()
        fig_kum.add_trace(go.Scatter(
            x=BULAN, y=d['kumulatif'], mode='lines+markers',
            fill='tozeroy', fillcolor='rgba(46,134,222,0.1)',
            line=dict(color='#2e86de', width=2.5),
            marker=dict(size=7, color='#2e86de'),
            hovertemplate='<b>%{x}</b><br>Kumulatif: %{y:.2f}%<extra></extra>'
        ))
        fig_kum.update_layout(
            margin=dict(t=10,b=60,l=45,r=10),
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='#fafafa',
            xaxis=dict(tickangle=-30, gridcolor='#f0f0f0'),
            yaxis=dict(title='%', gridcolor='#f0f0f0'),
            height=300, font=dict(family='Segoe UI', size=12)
        )
        st.plotly_chart(fig_kum, use_container_width=True, config={'displayModeBar':False})
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">🌊 Pertumbuhan Delta % per Bulan</div>', unsafe_allow_html=True)
        sel_delta = st.selectbox("Pilih Seksi", SEKSI_LIST, key='delta_seksi',
                                 format_func=lambda x: f"{x} — {SEKSI_NAMA[x]}")
        d2 = hasil_monitoring[sel_delta]
        delta_colors = ['#00b894' if v >= 0 else '#d63031' for v in d2['delta']]
        fig_delta = go.Figure()
        fig_delta.add_trace(go.Bar(
            x=BULAN, y=d2['delta'],
            marker=dict(color=delta_colors),
            hovertemplate='<b>%{x}</b><br>Delta: %{y:+.2f}%<extra></extra>'
        ))
        fig_delta.update_layout(
            margin=dict(t=10,b=60,l=45,r=10),
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='#fafafa',
            xaxis=dict(tickangle=-30, gridcolor='#f0f0f0'),
            yaxis=dict(title='%', gridcolor='#f0f0f0', zeroline=True, zerolinecolor='#ccc'),
            height=300, font=dict(family='Segoe UI', size=12)
        )
        st.plotly_chart(fig_delta, use_container_width=True, config={'displayModeBar':False})
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Tabel Detail ──
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    st.markdown('<div class="chart-title">📋 Tabel % Realisasi Bulanan per Seksi</div>', unsafe_allow_html=True)
    rows = []
    for s in SEKSI_LIST:
        d = hasil_monitoring[s]
        row = {'Seksi': s}
        for j, b in enumerate(BULAN):
            v = d['pct_bulanan'][j]
            row[b[:3]] = f"{v:.2f}%" if v > 0 else "—"
        row['Kumulatif'] = f"{d['kumulatif'][-1]}%"
        sp = d['sisa']['sisa_pct']
        row['Sisa (%)'] = f"{'🟢' if sp<=20 else '🟡' if sp<=50 else '🔴'} {sp}%"
        rows.append(row)
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# HALAMAN 3 — ANALISIS & PROYEKSI
# ═══════════════════════════════════════════════════════════════════════════════
elif halaman == "📈 Analisis & Proyeksi":
    st.markdown("<h4 style='color:#1a3a6b;font-weight:800;margin-bottom:8px'>📈 Analisis & Proyeksi Realisasi Anggaran</h4>", unsafe_allow_html=True)

    tab_proj, tab_tren, tab_risiko = st.tabs(["📅 Proyeksi 2026–2027", "📉 Tren 2018–2027", "⚠️ Risiko & Toleransi"])

    # ── TAB 1: PROYEKSI ──
    with tab_proj:
        st.markdown("### 📅 Proyeksi Realisasi Bulanan per Seksi")

        tahun_proj = st.radio("Pilih Tahun Proyeksi", ["2026","2027"], horizontal=True)
        proj_data  = PROYEKSI_2026 if tahun_proj == "2026" else PROYEKSI_2027

        fig_proj = go.Figure()
        for i, s in enumerate(SEKSI_LIST):
            fig_proj.add_trace(go.Bar(
                name=s, x=BULAN, y=proj_data[s],
                marker=dict(color=COLORS[i]),
                hovertemplate=f'<b>{s}</b><br>%{{x}}: %{{y:.2f}}%<extra></extra>'
            ))
        fig_proj.update_layout(
            barmode='group',
            margin=dict(t=20,b=80,l=50,r=20),
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='#fafafa',
            yaxis=dict(title='% Realisasi', gridcolor='#f0f0f0'),
            xaxis=dict(tickangle=-20),
            legend=dict(orientation='h', y=-0.3),
            height=420, font=dict(family='Segoe UI')
        )
        st.plotly_chart(fig_proj, use_container_width=True, config={'displayModeBar':False})

        # Tabel 2026 & 2027
        c1, c2 = st.columns(2)
        for col, tahun_t, pdata in [(c1,"2026",PROYEKSI_2026),(c2,"2027",PROYEKSI_2027)]:
            with col:
                st.markdown(f"**📅 Proyeksi {tahun_t} per Seksi (%)**")
                rows = []
                for s in SEKSI_LIST:
                    row = {'Seksi': s}
                    for j, b in enumerate(BULAN):
                        row[b[:3]] = pdata[s][j]
                    row['Total'] = round(sum(pdata[s]),2)
                    rows.append(row)
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # ── TAB 2: TREN ──
    with tab_tren:
        st.markdown("### 📉 Tren Realisasi Anggaran Tahunan per Seksi (2018–2027)")

        tahun_all = list(range(2018, 2028))
        fig_tren = go.Figure()
        for i, s in enumerate(SEKSI_LIST):
            vals = []
            for t in tahun_all:
                if t < 2026:
                    vals.append(TREN[s].get(t, None))
                elif t == 2026:
                    vals.append(round(sum(PROYEKSI_2026[s]),2))
                else:
                    vals.append(round(sum(PROYEKSI_2027[s]),2))
            fig_tren.add_trace(go.Scatter(
                name=f"{s} — {SEKSI_NAMA[s][:18]}",
                x=tahun_all, y=vals, mode='lines+markers',
                line=dict(color=COLORS[i], width=2.5),
                marker=dict(size=8),
                hovertemplate=f'<b>{s}</b> %{{x}}<br>%{{y:.2f}}%<extra></extra>'
            ))
        fig_tren.add_vline(x=2025.5, line_dash="dash", line_color="#636e72", line_width=1.5)
        fig_tren.add_annotation(x=2025.5, y=108, text="← Aktual | Proyeksi →",
                                 showarrow=False, font=dict(color='#636e72', size=11))
        fig_tren.update_layout(
            margin=dict(t=20,b=60,l=50,r=20),
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='#fafafa',
            xaxis=dict(title='Tahun', gridcolor='#f0f0f0'),
            yaxis=dict(title='Total % Realisasi', gridcolor='#f0f0f0', range=[0,115]),
            legend=dict(orientation='h', y=-0.25),
            height=420, font=dict(family='Segoe UI')
        )
        st.plotly_chart(fig_tren, use_container_width=True, config={'displayModeBar':False})

        # Heatmap
        st.markdown("### 🔥 Heatmap Pola Realisasi Bulanan (2018–2025)")
        tahun_hist = list(range(2018,2026))
        z_data = []
        for s in SEKSI_LIST:
            row_vals = []
            for t in tahun_hist:
                pct = get_pct_bulanan(df, s, t)
                row_vals.append(round(sum(pct)/len([p for p in pct if p>0]) if any(p>0 for p in pct) else 0, 1))
            z_data.append(row_vals)

        fig_heat = go.Figure(go.Heatmap(
            z=z_data, x=tahun_hist, y=SEKSI_LIST,
            colorscale=[[0,'#fff0f0'],[0.5,'#ffd580'],[1,'#1a3a6b']],
            text=[[f"{v:.1f}%" for v in row] for row in z_data],
            texttemplate='%{text}',
            hovertemplate='Seksi %{y}<br>Tahun %{x}<br>Rata-rata: %{z:.1f}%<extra></extra>'
        ))
        fig_heat.update_layout(
            margin=dict(t=20,b=40,l=60,r=80),
            paper_bgcolor='rgba(0,0,0,0)',
            height=300, font=dict(family='Segoe UI')
        )
        st.plotly_chart(fig_heat, use_container_width=True, config={'displayModeBar':False})

        # Tabel tren
        st.markdown("**📋 Data Tren Realisasi Tahunan (%)**")
        rows = []
        for s in SEKSI_LIST:
            row = {'Seksi': f"{s} — {SEKSI_NAMA[s]}"}
            for t in range(2018,2026):
                row[str(t)] = TREN[s].get(t,0)
            row['2026 (P)'] = round(sum(PROYEKSI_2026[s]),2)
            row['2027 (P)'] = round(sum(PROYEKSI_2027[s]),2)
            rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # ── TAB 3: RISIKO ──
    with tab_risiko:
        st.markdown("### ⚠️ Klasifikasi Risiko & Toleransi Deviasi")

        c1, c2 = st.columns(2)
        risiko_clr = {'Rendah':'#00b894','Sedang':'#fdcb6e','Tinggi':'#d63031'}

        with c1:
            st.markdown("**🔴 Klasifikasi Risiko Deviasi per Seksi**")
            mae_vals = [REGRESI[s]['mae'] for s in SEKSI_LIST]
            risiko_vals = [REGRESI[s]['risiko'] for s in SEKSI_LIST]
            bubble_colors = [risiko_clr[REGRESI[s]['risiko']] for s in SEKSI_LIST]
            fig_risk = go.Figure(go.Scatter(
                x=SEKSI_LIST, y=mae_vals,
                mode='markers+text', text=risiko_vals, textposition='top center',
                marker=dict(
                    size=[v*12+20 for v in mae_vals],
                    color=bubble_colors, line=dict(width=2, color='#fff')
                ),
                hovertemplate='<b>%{x}</b><br>MAE: %{y:.2f}%<br>Risiko: %{text}<extra></extra>'
            ))
            fig_risk.add_hline(y=1, line_dash='dot', line_color='#fdcb6e', line_width=2)
            fig_risk.add_hline(y=3, line_dash='dot', line_color='#d63031', line_width=2)
            fig_risk.update_layout(
                margin=dict(t=20,b=40,l=50,r=20),
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='#fafafa',
                yaxis=dict(title='MAE (%)', gridcolor='#f0f0f0'),
                height=360, font=dict(family='Segoe UI')
            )
            st.plotly_chart(fig_risk, use_container_width=True, config={'displayModeBar':False})

        with c2:
            st.markdown("**⚖️ MAE per Seksi**")
            fig_mae = go.Figure(go.Bar(
                orientation='h',
                x=mae_vals,
                y=[f"{s} — {SEKSI_NAMA[s][:15]}" for s in SEKSI_LIST],
                marker=dict(color=bubble_colors),
                text=[f"{v}%" for v in mae_vals], textposition='outside',
                hovertemplate='%{y}<br>MAE: %{x:.2f}%<extra></extra>'
            ))
            fig_mae.update_layout(
                margin=dict(t=20,b=40,l=200,r=60),
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='#fafafa',
                xaxis=dict(title='MAE (%)', gridcolor='#f0f0f0'),
                height=360, font=dict(family='Segoe UI')
            )
            st.plotly_chart(fig_mae, use_container_width=True, config={'displayModeBar':False})

        # Tabel toleransi
        st.markdown("**📏 Tabel Toleransi Deviasi ±5% — Proyeksi 2026**")
        rows = []
        for s in SEKSI_LIST:
            for j, b in enumerate(BULAN[:6]):  # 6 bulan pertama
                proj = PROYEKSI_2026[s][j]
                rows.append({
                    'Seksi': f"{s} — {SEKSI_NAMA[s]}",
                    'Bulan': b,
                    'Proyeksi (%)': proj,
                    'Batas Bawah (-5%)': round(proj-5, 2),
                    'Batas Atas (+5%)': round(proj+5, 2),
                    'Keterangan': f"Jika < {round(proj-5,2)}% → perlu tindak lanjut"
                })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        # Tabel regresi
        st.markdown("**📐 Ringkasan Hasil Regresi Linear per Seksi**")
        rows_reg = []
        for s in SEKSI_LIST:
            r = REGRESI[s]
            rows_reg.append({
                'Seksi': s, 'Nama': SEKSI_NAMA[s],
                'a (Intersep)': r['a'], 'b (Koefisien)': r['b'],
                'R²': r['r2'],
                'Normalitas': '✅ Normal' if r['normal'] else '⚠️ Tidak Normal',
                'p-value': r['p_value'],
                'Transformasi': '√ (Sqrt)' if r.get('transformasi') else '—',
                'MAE (%)': r['mae'],
                'Risiko': r['risiko']
            })
        st.dataframe(pd.DataFrame(rows_reg), use_container_width=True, hide_index=True)
