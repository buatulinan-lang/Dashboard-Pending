"""
Dashboard Service Cabang (Offline) - Streamlit App
====================================================
Jalankan lokal dengan:
    pip install -r requirements.txt
    streamlit run app.py

Lalu upload file Excel (format sama seperti Gabungan_Semua_Cabang.xlsx:
satu sheet per cabang, dengan kolom baku seperti di data awal).
Setiap kali ada data baru, cukup upload ulang file -> dashboard otomatis
hitung ulang. Tidak perlu internet, semua proses jalan di komputer sendiri.
"""

import calendar
import io
from datetime import date

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
# Konfigurasi halaman & style
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Dashboard Service Cabang", layout="wide", page_icon="📊")

BULAN_NAMES = ['', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
               'Agustus', 'September', 'Oktober', 'November', 'Desember']

STATUS_COLOR = {
    'DONE': '#16a34a',
    'PENDING': '#f59e0b',
    'CANCEL': '#dc2626',
    'LAINNYA': '#94a3b8',
}

PALETTE = ['#6d3fbf', '#c93fa8', '#17a3a3', '#e0921f', '#3f8ac9',
           '#e0475a', '#3fbf7f', '#a855f7', '#eab308', '#0ea5e9']
PALETTE_URGENT = ['#c9392f', '#e0475a', '#e0721f', '#e0921f', '#c9682f',
                  '#b5473f', '#d1585a', '#e2a63a', '#c95050', '#d98a3f']

REQUIRED_COLUMNS = [
    'NOMOR PENGIRIMAN PESANAN', 'TGL PENGIRIMAN', 'STATUS PENGERJAAN',
    'NAMA TEKNISI', 'KERUSAKAN UTAMA',
]

st.markdown("""
<style>
  .kpi-wrap{display:grid;grid-template-columns:repeat(6,1fr);gap:14px;margin-bottom:6px;}
  .kpi{border-radius:16px;padding:16px 16px 18px;color:#fff;position:relative;overflow:hidden;
       min-height:112px;box-shadow:0 8px 20px rgba(30,20,60,.14);}
  .kpi .label{font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;opacity:.92;}
  .kpi .value{font-size:22px;font-weight:800;margin-top:10px;line-height:1.15;}
  .kpi .foot{font-size:11px;margin-top:6px;opacity:.9;}
  .badge-pill{display:inline-block;background:linear-gradient(90deg,#17a3a3,#0f8a82);color:#fff;
       font-size:12px;font-weight:700;padding:6px 14px;border-radius:999px;}
  .warn-banner{background:linear-gradient(90deg,#fff6f5,#fff);border:1px solid #f4c9c4;border-radius:14px;
       padding:14px 18px;margin-bottom:10px;}
  .warn-banner b{color:#7a2a24;}
  .warn-banner span{font-size:12.5px;color:#7a2a24;line-height:1.6;}
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Logika data (dedup + klasifikasi status + agregasi)
# ---------------------------------------------------------------------------

def classify_status(status) -> str:
    if pd.isna(status):
        return 'LAINNYA'
    s = str(status).strip().upper()
    if s in ('', 'N/A'):
        return 'LAINNYA'
    if s.startswith('CANCEL'):
        return 'CANCEL'
    if 'DONE' in s:
        return 'DONE'
    if 'PENDING' in s:
        return 'PENDING'
    if s == 'COMPLAIN':
        return 'PENDING'
    return 'LAINNYA'


def jenis_pending(status) -> str:
    s = str(status).upper() if pd.notna(status) else ''
    if 'TEKNISI' in s:
        return 'Teknisi'
    if 'CUSTOMER' in s:
        return 'Customer'
    if 'SPAREPART' in s:
        return 'Sparepart'
    if 'KLAIM' in s or 'GARANSI' in s:
        return 'Klaim Garansi'
    if 'KOMPLAIN' in s or 'COMPLAIN' in s:
        return 'Komplain'
    return 'Umum'


@st.cache_data(show_spinner="Membaca & memproses file Excel (bisa beberapa puluh detik untuk file besar)...")
def load_data(file_bytes: bytes) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
    frames = []
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        if df.empty:
            continue
        df['CABANG'] = sheet
        frames.append(df)
    if not frames:
        return pd.DataFrame()

    full = pd.concat(frames, ignore_index=True, sort=False)

    missing = [c for c in REQUIRED_COLUMNS if c not in full.columns]
    if missing:
        raise ValueError(
            "Kolom berikut tidak ditemukan di file: " + ", ".join(missing) +
            ". Pastikan format file sama seperti Gabungan_Semua_Cabang.xlsx."
        )

    # Baris yang seluruh kolom aslinya identik dianggap 1 transaksi.
    original_cols = [c for c in full.columns if c != 'CABANG']
    total_raw = len(full)
    full = full.drop_duplicates(subset=original_cols + ['CABANG'], keep='first').reset_index(drop=True)
    total_unique = len(full)

    full['TGL PENGIRIMAN'] = pd.to_datetime(full['TGL PENGIRIMAN'], errors='coerce')
    full['TAHUN'] = full['TGL PENGIRIMAN'].dt.year
    full['BULAN'] = full['TGL PENGIRIMAN'].dt.month
    full['STATUS_BUCKET'] = full['STATUS PENGERJAAN'].apply(classify_status)
    full['TEKNISI'] = full['NAMA TEKNISI'].apply(
        lambda v: str(v).strip() if pd.notna(v) and str(v).strip() else 'TIDAK ADA TEKNISI'
    )
    full['KERUSAKAN'] = full['KERUSAKAN UTAMA'].apply(
        lambda v: str(v).strip().upper() if pd.notna(v) and str(v).strip() else 'TIDAK ADA DATA'
    )

    full.attrs['total_raw_rows'] = total_raw
    full.attrs['total_unique'] = total_unique
    return full


def apply_filters(df: pd.DataFrame, tahun, bulan, cabang) -> pd.DataFrame:
    out = df
    if tahun != 'Semua Tahun':
        out = out[out['TAHUN'] == tahun]
    if bulan != 'Semua Bulan':
        out = out[out['BULAN'] == bulan]
    if cabang != 'Semua Cabang':
        out = out[out['CABANG'] == cabang]
    return out


def compute_period_days(df: pd.DataFrame) -> int:
    sub = df.dropna(subset=['TAHUN', 'BULAN'])
    if sub.empty:
        return 0
    pairs = sub[['TAHUN', 'BULAN']].drop_duplicates()
    today = date.today()
    total_days = 0
    for _, row in pairs.iterrows():
        y, m = int(row['TAHUN']), int(row['BULAN'])
        days_in_month = calendar.monthrange(y, m)[1]
        if y == today.year and m == today.month:
            days_in_month = min(days_in_month, today.day)
        elif (y > today.year) or (y == today.year and m > today.month):
            days_in_month = 0
        total_days += days_in_month
    return total_days


def kpi_html(cards) -> str:
    cells = []
    for c in cards:
        cells.append(f"""
        <div class="kpi" style="background:{c['grad']}">
          <div class="label">{c['label']}</div>
          <div class="value">{c['value']}</div>
          <div class="foot">{c.get('sub', '&nbsp;')}</div>
        </div>""")
    return f'<div class="kpi-wrap">{"".join(cells)}</div>'


# ---------------------------------------------------------------------------
# Sidebar: upload file & filter
# ---------------------------------------------------------------------------
st.sidebar.title("📁 Sumber Data")
uploaded = st.sidebar.file_uploader(
    "Upload file Excel (satu sheet per cabang)",
    type=['xlsx'],
    help="Format sama seperti Gabungan_Semua_Cabang.xlsx. Upload ulang kapan saja ada data baru."
)

if uploaded is None:
    st.title("📊 Dashboard Service Cabang")
    st.info(
        "Silakan upload file Excel data (format: satu sheet per cabang, kolom baku seperti "
        "NOMOR PENGIRIMAN PESANAN, TGL PENGIRIMAN, STATUS PENGERJAAN, NAMA TEKNISI, KERUSAKAN UTAMA, dst) "
        "lewat panel di sebelah kiri untuk mulai."
    )
    st.stop()

try:
    data = load_data(uploaded.getvalue())
except ValueError as e:
    st.error(str(e))
    st.stop()

if data.empty:
    st.warning("File tidak berisi data yang bisa dibaca.")
    st.stop()

total_raw_rows = data.attrs.get('total_raw_rows', len(data))
total_unique = data.attrs.get('total_unique', len(data))

tahun_opts = ['Semua Tahun'] + sorted([int(t) for t in data['TAHUN'].dropna().unique()])
bulan_opts = ['Semua Bulan'] + sorted([int(b) for b in data['BULAN'].dropna().unique()])
cabang_opts = ['Semua Cabang'] + sorted(data['CABANG'].dropna().unique().tolist())

st.sidebar.title("🔎 Filter")
f_tahun = st.sidebar.selectbox("Tahun", tahun_opts, format_func=lambda x: x)
f_bulan = st.sidebar.selectbox("Bulan", bulan_opts, format_func=lambda x: x if isinstance(x, str) else BULAN_NAMES[x])
f_cabang = st.sidebar.selectbox("Cabang", cabang_opts, format_func=lambda x: x)

st.sidebar.markdown("---")
st.sidebar.caption(
    f"📦 {total_raw_rows:,} baris mentah → {total_unique:,} transaksi unik "
    f"({total_raw_rows - total_unique:,} duplikat dihapus)."
)

filtered = apply_filters(data, f_tahun, f_bulan, f_cabang)

# ---------------------------------------------------------------------------
# TABS
# ---------------------------------------------------------------------------
tab_main, tab_pending = st.tabs(["📊 Dashboard Utama", "⚠️ Dashboard Pending"])

# =============================================================================
# TAB 1: DASHBOARD UTAMA
# =============================================================================
with tab_main:
    st.markdown("## Dashboard Status Pengerjaan — Semua Cabang")
    st.caption(
        f"{data['CABANG'].nunique()} cabang · {total_unique:,} transaksi unik "
        f"(dari {total_raw_rows:,} baris mentah)"
    )

    total = len(filtered)
    counts = filtered['STATUS_BUCKET'].value_counts()
    done = int(counts.get('DONE', 0))
    pending = int(counts.get('PENDING', 0))
    cancel = int(counts.get('CANCEL', 0))
    lainnya = int(counts.get('LAINNYA', 0))

    def pct(n):
        return f"{(n / total * 100):.1f}%" if total else "0%"

    period_days = compute_period_days(filtered)
    avg_day = (total / period_days) if period_days else 0

    cards = [
        {'label': 'Total Transaksi Unik', 'value': f"{total:,}", 'sub': 'sesuai filter aktif',
         'grad': 'linear-gradient(135deg,#3b5bfd,#5a72ff)'},
        {'label': 'Done', 'value': f"{done:,}", 'sub': f"{pct(done)} dari total",
         'grad': 'linear-gradient(135deg,#16a34a,#22c55e)'},
        {'label': 'Pending', 'value': f"{pending:,}", 'sub': f"{pct(pending)} dari total",
         'grad': 'linear-gradient(135deg,#f59e0b,#fbbf24)'},
        {'label': 'Cancel', 'value': f"{cancel:,}", 'sub': f"{pct(cancel)} dari total",
         'grad': 'linear-gradient(135deg,#dc2626,#ef4444)'},
        {'label': 'Lainnya', 'value': f"{lainnya:,}", 'sub': f"{pct(lainnya)} dari total",
         'grad': 'linear-gradient(135deg,#64748b,#94a3b8)'},
        {'label': 'Rata-rata / Hari', 'value': f"{avg_day:,.1f}",
         'sub': f"{period_days} hari periode" if period_days else '&nbsp;',
         'grad': 'linear-gradient(135deg,#7c3aed,#a855f7)'},
    ]
    st.markdown(kpi_html(cards), unsafe_allow_html=True)
    st.write("")

    col1, col2 = st.columns([1.2, 1])

    with col1:
        st.markdown("#### Tren Transaksi Sepanjang Tahun")
        chart_tahun_opts = sorted([int(t) for t in data['TAHUN'].dropna().unique()])
        default_year = f_tahun if f_tahun != 'Semua Tahun' else (chart_tahun_opts[-1] if chart_tahun_opts else None)
        chart_tahun = st.selectbox("Tahun grafik", chart_tahun_opts,
                                    index=chart_tahun_opts.index(default_year) if default_year in chart_tahun_opts else 0,
                                    key="main_chart_tahun")
        cabang_scope = data if f_cabang == 'Semua Cabang' else data[data['CABANG'] == f_cabang]
        year_df = cabang_scope[cabang_scope['TAHUN'] == chart_tahun]
        monthly = (year_df.groupby(['BULAN', 'STATUS_BUCKET']).size()
                   .reset_index(name='count'))
        fig = go.Figure()
        for status in ['DONE', 'PENDING', 'CANCEL', 'LAINNYA']:
            sub = monthly[monthly['STATUS_BUCKET'] == status]
            y = [int(sub[sub['BULAN'] == m]['count'].sum()) for m in range(1, 13)]
            fig.add_bar(x=[BULAN_NAMES[m][:3] for m in range(1, 13)], y=y, name=status,
                        marker_color=STATUS_COLOR[status])
        fig.update_layout(barmode='stack', height=320, margin=dict(l=10, r=10, t=10, b=10),
                           legend=dict(orientation='h', y=1.12))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("#### Komposisi Status")
        comp_df = pd.DataFrame({
            'Status': ['Done', 'Pending', 'Cancel', 'Lainnya'],
            'Jumlah': [done, pending, cancel, lainnya]
        })
        fig2 = px.pie(comp_df, names='Status', values='Jumlah', hole=0.55,
                      color='Status',
                      color_discrete_map={'Done': STATUS_COLOR['DONE'], 'Pending': STATUS_COLOR['PENDING'],
                                          'Cancel': STATUS_COLOR['CANCEL'], 'Lainnya': STATUS_COLOR['LAINNYA']})
        fig2.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### Rekap per Cabang")
    rekap_cabang = (filtered.groupby('CABANG')['STATUS_BUCKET']
                     .value_counts().unstack(fill_value=0))
    for col in ['DONE', 'PENDING', 'CANCEL', 'LAINNYA']:
        if col not in rekap_cabang.columns:
            rekap_cabang[col] = 0
    rekap_cabang['Total'] = rekap_cabang[['DONE', 'PENDING', 'CANCEL', 'LAINNYA']].sum(axis=1)
    rekap_cabang['% Done'] = (rekap_cabang['DONE'] / rekap_cabang['Total'] * 100).round(1)
    rekap_cabang = rekap_cabang.sort_values('Total', ascending=False)
    st.dataframe(
        rekap_cabang[['Total', 'DONE', 'PENDING', 'CANCEL', 'LAINNYA', '% Done']],
        use_container_width=True
    )

    with st.expander("ℹ️ Catatan metodologi"):
        st.write(
            "Baris dengan seluruh kolom identik dianggap 1 transaksi. Pengelompokan status: **Done** mencakup "
            "semua status berisi kata DONE; **Pending** mencakup semua status berisi kata PENDING serta status "
            "COMPLAIN; **Cancel** mencakup semua status berisi kata CANCEL; sisanya masuk **Lainnya**. Filter "
            "Tahun/Bulan mengacu kolom TGL PENGIRIMAN. Rata-rata/hari dihitung dari total transaksi sesuai filter "
            "dibagi jumlah hari kalender bulan-bulan yang tercakup (bulan berjalan dipotong sampai hari ini)."
        )

# =============================================================================
# TAB 2: DASHBOARD PENDING
# =============================================================================
with tab_pending:
    st.markdown("## Dashboard Pending — Breakdown Teknisi & Kerusakan")

    st.markdown("""
    <div class="warn-banner">
      ⚠️ <b>Pending bukan pencapaian — ini beban kerja yang harus segera dituntaskan.</b>
      <span>Semakin lama status tertahan, semakin besar risiko komplain dari customer. Gunakan ranking di bawah
      untuk menentukan prioritas tindak lanjut (teknisi, cabang, dan jenis kerusakan mana yang paling butuh perhatian).</span>
    </div>
    """, unsafe_allow_html=True)

    pending_all = filtered[filtered['STATUS_BUCKET'] == 'PENDING'].copy()
    total_p = len(pending_all)

    pending_all['JENIS'] = pending_all['STATUS PENGERJAAN'].apply(jenis_pending)

    period_days_p = compute_period_days(pending_all)
    avg_day_p = (total_p / period_days_p) if period_days_p else 0

    if total_p:
        top_tek = (pending_all[~pending_all['TEKNISI'].isin(['TIDAK ADA TEKNISI', 'N/A'])]
                   ['TEKNISI'].value_counts())
        top_tek_name = top_tek.index[0] if len(top_tek) else '-'
        top_tek_count = int(top_tek.iloc[0]) if len(top_tek) else 0
        top_cabang = pending_all['CABANG'].value_counts()
        top_cabang_name = top_cabang.index[0]
        top_cabang_count = int(top_cabang.iloc[0])
        top_jenis = pending_all['JENIS'].value_counts()
        top_jenis_name = top_jenis.index[0]
        top_jenis_count = int(top_jenis.iloc[0])
        top_ker = pending_all['KERUSAKAN'].value_counts()
        top_ker_name = top_ker.index[0]
        top_ker_count = int(top_ker.iloc[0])
    else:
        top_tek_name, top_tek_count = '-', 0
        top_cabang_name, top_cabang_count = '-', 0
        top_jenis_name, top_jenis_count = '-', 0
        top_ker_name, top_ker_count = '-', 0

    pct_all = (total_p / total_unique * 100) if total_unique else 0

    cards_p = [
        {'label': 'Total Pending', 'value': f"{total_p:,}", 'sub': 'transaksi pending (unik)',
         'grad': 'linear-gradient(135deg,#6d3fbf,#8e3fc0)'},
        {'label': 'Teknisi Terbanyak', 'value': top_tek_name,
         'sub': f"{top_tek_count} pending ({(top_tek_count/total_p*100 if total_p else 0):.1f}%)",
         'grad': 'linear-gradient(135deg,#9c3fc0,#c93fa8)'},
        {'label': 'Cabang Terbanyak', 'value': top_cabang_name,
         'sub': f"{top_cabang_count} pending ({(top_cabang_count/total_p*100 if total_p else 0):.1f}%)",
         'grad': 'linear-gradient(135deg,#b23f9e,#d1478d)'},
        {'label': 'Rata-rata / Hari', 'value': f"{avg_day_p:,.2f}",
         'sub': f"{period_days_p} hari periode" if period_days_p else '&nbsp;',
         'grad': 'linear-gradient(135deg,#17a3a3,#159a8d)'},
        {'label': 'Jenis Dominan', 'value': top_jenis_name,
         'sub': f"{top_jenis_count} kasus ({(top_jenis_count/total_p*100 if total_p else 0):.1f}%)",
         'grad': 'linear-gradient(135deg,#0f8a82,#0c7a6e)'},
        {'label': 'Kerusakan Terbanyak', 'value': top_ker_name,
         'sub': f"{top_ker_count} kasus ({(top_ker_count/total_p*100 if total_p else 0):.1f}%)",
         'grad': 'linear-gradient(135deg,#c9392f,#e0475a)'},
    ]
    st.markdown(kpi_html(cards_p), unsafe_allow_html=True)
    st.write("")

    if total_p == 0:
        st.info("Tidak ada data pending untuk filter yang dipilih.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### Teknisi dengan Pending Terbanyak — Perlu Ditindaklanjuti")
            tek_rank = (pending_all['TEKNISI']
                        .replace('TIDAK ADA TEKNISI', 'Belum ada teknisi')
                        .value_counts().head(10).sort_values(ascending=True))
            fig3 = go.Figure(go.Bar(
                x=tek_rank.values, y=tek_rank.index, orientation='h',
                marker_color=PALETTE_URGENT[:len(tek_rank)][::-1] if len(tek_rank) <= len(PALETTE_URGENT) else PALETTE_URGENT,
                text=[f"{v} ({v/total_p*100:.1f}%)" for v in tek_rank.values],
                textposition='outside'
            ))
            fig3.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig3, use_container_width=True)

        with col2:
            st.markdown("#### Kerusakan Paling Sering Menumpuk Jadi Pending")
            ker_rank = pending_all['KERUSAKAN'].value_counts().head(10).sort_values(ascending=True)
            fig4 = go.Figure(go.Bar(
                x=ker_rank.values, y=ker_rank.index, orientation='h',
                marker_color=PALETTE_URGENT[:len(ker_rank)][::-1] if len(ker_rank) <= len(PALETTE_URGENT) else PALETTE_URGENT,
                text=[f"{v} ({v/total_p*100:.1f}%)" for v in ker_rank.values],
                textposition='outside'
            ))
            fig4.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig4, use_container_width=True)

        col3, col4, col5 = st.columns(3)
        with col3:
            st.markdown("###### Jenis Pending")
            jd = pending_all['JENIS'].value_counts().reset_index()
            jd.columns = ['Jenis', 'Jumlah']
            figd1 = px.pie(jd, names='Jenis', values='Jumlah', hole=0.55,
                            color_discrete_sequence=PALETTE)
            figd1.update_layout(height=280, margin=dict(l=5, r=5, t=5, b=5), showlegend=True,
                                 legend=dict(font=dict(size=9)))
            st.plotly_chart(figd1, use_container_width=True)
        with col4:
            st.markdown("###### Per Cabang")
            cd = pending_all['CABANG'].value_counts().reset_index()
            cd.columns = ['Cabang', 'Jumlah']
            figd2 = px.pie(cd, names='Cabang', values='Jumlah', hole=0.55,
                            color_discrete_sequence=PALETTE)
            figd2.update_layout(height=280, margin=dict(l=5, r=5, t=5, b=5), showlegend=True,
                                 legend=dict(font=dict(size=9)))
            st.plotly_chart(figd2, use_container_width=True)
        with col5:
            st.markdown("###### Per Teknisi (Top 5 + Lainnya)")
            td = pending_all['TEKNISI'].replace('TIDAK ADA TEKNISI', 'Belum ada teknisi').value_counts()
            top5 = td.head(5)
            rest = td.iloc[5:].sum()
            if rest > 0:
                top5 = pd.concat([top5, pd.Series({'Lainnya': rest})])
            figd3 = px.pie(names=top5.index, values=top5.values, hole=0.55,
                            color_discrete_sequence=PALETTE)
            figd3.update_layout(height=280, margin=dict(l=5, r=5, t=5, b=5), showlegend=True,
                                 legend=dict(font=dict(size=9)))
            st.plotly_chart(figd3, use_container_width=True)

        st.markdown("#### Ranking Lengkap Teknisi")
        tek_full = (pending_all['TEKNISI'].replace('TIDAK ADA TEKNISI', 'Belum ada teknisi')
                    .value_counts().reset_index())
        tek_full.columns = ['Teknisi', 'Jumlah Pending']
        tek_full['Porsi (%)'] = (tek_full['Jumlah Pending'] / total_p * 100).round(1)
        st.dataframe(tek_full, use_container_width=True, height=300)

        st.markdown("#### Ranking Lengkap Jenis Kerusakan")
        ker_full = pending_all['KERUSAKAN'].value_counts().reset_index()
        ker_full.columns = ['Kerusakan Utama', 'Jumlah Pending']
        ker_full['Porsi (%)'] = (ker_full['Jumlah Pending'] / total_p * 100).round(1)
        st.dataframe(ker_full, use_container_width=True, height=300)

        st.markdown("#### Detail Transaksi Pending")
        search = st.text_input("Cari teknisi / nomor / customer / kerusakan", key="pending_search")
        detail_cols = ['TGL PENGIRIMAN', 'CABANG', 'TEKNISI', 'STATUS PENGERJAAN', 'NAMA CUSTOMER', 'KERUSAKAN UTAMA']
        detail_cols = [c for c in detail_cols if c in pending_all.columns]
        detail_df = pending_all[detail_cols].copy()
        if search:
            mask = detail_df.apply(lambda row: search.upper() in ' '.join(str(v) for v in row.values).upper(), axis=1)
            detail_df = detail_df[mask]
        st.dataframe(detail_df.sort_values('TGL PENGIRIMAN', ascending=False), use_container_width=True, height=350)

    with st.expander("ℹ️ Catatan metodologi"):
        st.write(
            "Yang termasuk **Pending** adalah status berisi kata PENDING serta status COMPLAIN yang belum selesai. "
            "Teknisi berlabel **Belum ada teknisi** berarti kolom Nama Teknisi kosong pada baris transaksi tsb. "
            "Rata-rata/hari dihitung dari total pending sesuai filter dibagi jumlah hari kalender bulan-bulan yang "
            "tercakup filter (bulan berjalan dipotong sampai tanggal hari ini)."
        )
