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
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
# Konfigurasi halaman & style
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Dashboard Service Cabang", layout="wide", page_icon="📊")

# Data bawaan yang tersimpan permanen di dalam repo (folder data/). Kalau file
# ini ada, dashboard otomatis memuatnya tanpa perlu upload manual setiap kali
# app "bangun" lagi setelah tidur (di Streamlit Community Cloud). Untuk update
# data secara permanen, ganti/timpa file ini di GitHub lalu reboot app.
DEFAULT_DATA_PATH = Path(__file__).parent / "data" / "latest_data.csv.gz"

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
PALETTE_SUCCESS = ['#16a34a', '#0f8a82', '#22c55e', '#17a3a3', '#3fbf7f',
                    '#0c7a6e', '#4ade80', '#0ea5e9', '#65b83f', '#159a8d']
PALETTE_CANCEL = ['#8b1e1e', '#a33131', '#b33f3f', '#7a1f1f', '#c94f4f',
                   '#992e2e', '#6e1a1a', '#d16060', '#851f35', '#a83250']

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
  .cancel-banner{background:linear-gradient(90deg,#fdf3f2,#fff);border:1px solid #e8b9b3;border-radius:14px;
       padding:14px 18px;margin-bottom:10px;}
  .cancel-banner b{color:#7a1f1f;}
  .cancel-banner span{font-size:12.5px;color:#7a1f1f;line-height:1.6;}
  .success-banner{background:linear-gradient(90deg,#f2fbf5,#fff);border:1px solid #b9e8c9;border-radius:14px;
       padding:14px 18px;margin-bottom:10px;}
  .success-banner b{color:#0f5132;}
  .success-banner span{font-size:12.5px;color:#0f5132;line-height:1.6;}
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


def jenis_done(status) -> str:
    s = str(status).upper() if pd.notna(status) else ''
    if 'DIAMBIL' in s:
        return 'Sudah Diambil'
    if 'KLAIM' in s or 'GARANSI' in s:
        return 'Klaim Garansi'
    if 'KOMPLAIN' in s or 'COMPLAIN' in s:
        return 'Komplain'
    return 'Standar'


def jenis_cancel(status) -> str:
    s = str(status).upper() if pd.notna(status) else ''
    if 'DIAMBIL' in s:
        return 'Diambil'
    if 'TEKNISI' in s:
        return 'Teknisi'
    if 'CUSTOMER' in s or 'USER' in s:
        return 'Customer'
    if 'SPAREPART' in s:
        return 'Sparepart'
    if 'ADMIN' in s:
        return 'Admin'
    return 'Lainnya'


def _read_xlsx_raw(file_bytes: bytes) -> pd.DataFrame:
    """Baca file Excel (satu sheet per cabang) jadi satu DataFrame + kolom CABANG."""
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
    return pd.concat(frames, ignore_index=True, sort=False)


def _read_csv_gz_raw(file_bytes: bytes) -> pd.DataFrame:
    """Baca file data bawaan (csv terkompresi gzip) yang kolom CABANG-nya sudah ada."""
    return pd.read_csv(io.BytesIO(file_bytes), compression='gzip')


@st.cache_data(show_spinner="Membaca & memproses data (bisa beberapa puluh detik untuk file besar)...")
def load_data(file_bytes: bytes, source_kind: str) -> pd.DataFrame:
    if source_kind == 'csv_gz':
        full = _read_csv_gz_raw(file_bytes)
    else:
        full = _read_xlsx_raw(file_bytes)

    if full.empty:
        return pd.DataFrame()

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


def render_detail_dashboard(filtered_df, total_unique_all, *, status_bucket, jenis_func,
                             jenis_col_label, palette, banner_html, card_grads, rank_label,
                             note_text, key_prefix):
    """Dashboard breakdown detail (teknisi & kerusakan) untuk 1 status_bucket
    tertentu (PENDING / DONE / CANCEL). Dipakai bersama supaya ketiga
    dashboard status punya level detail yang sama persis."""
    sub = filtered_df[filtered_df['STATUS_BUCKET'] == status_bucket].copy()
    total_s = len(sub)

    if banner_html:
        st.markdown(banner_html, unsafe_allow_html=True)

    sub['JENIS'] = sub['STATUS PENGERJAAN'].apply(jenis_func)

    period_days_s = compute_period_days(sub)
    avg_day_s = (total_s / period_days_s) if period_days_s else 0

    if total_s:
        top_tek = (sub[~sub['TEKNISI'].isin(['TIDAK ADA TEKNISI', 'N/A'])]['TEKNISI'].value_counts())
        top_tek_name = top_tek.index[0] if len(top_tek) else '-'
        top_tek_count = int(top_tek.iloc[0]) if len(top_tek) else 0
        top_cabang = sub['CABANG'].value_counts()
        top_cabang_name = top_cabang.index[0]
        top_cabang_count = int(top_cabang.iloc[0])
        top_jenis = sub['JENIS'].value_counts()
        top_jenis_name = top_jenis.index[0]
        top_jenis_count = int(top_jenis.iloc[0])
        top_ker = sub['KERUSAKAN'].value_counts()
        top_ker_name = top_ker.index[0]
        top_ker_count = int(top_ker.iloc[0])
    else:
        top_tek_name, top_tek_count = '-', 0
        top_cabang_name, top_cabang_count = '-', 0
        top_jenis_name, top_jenis_count = '-', 0
        top_ker_name, top_ker_count = '-', 0

    def _pct(n):
        return f"{(n / total_s * 100):.1f}%" if total_s else "0%"

    cards = [
        {'label': f'Total {rank_label}', 'value': f"{total_s:,}", 'sub': f'transaksi {rank_label.lower()} (unik)',
         'grad': card_grads[0]},
        {'label': 'Teknisi Terbanyak', 'value': top_tek_name,
         'sub': f"{top_tek_count} ({_pct(top_tek_count)})", 'grad': card_grads[1]},
        {'label': 'Cabang Terbanyak', 'value': top_cabang_name,
         'sub': f"{top_cabang_count} ({_pct(top_cabang_count)})", 'grad': card_grads[2]},
        {'label': 'Rata-rata / Hari', 'value': f"{avg_day_s:,.2f}",
         'sub': f"{period_days_s} hari periode" if period_days_s else '&nbsp;', 'grad': card_grads[3]},
        {'label': f'Jenis {jenis_col_label} Dominan', 'value': top_jenis_name,
         'sub': f"{top_jenis_count} ({_pct(top_jenis_count)})", 'grad': card_grads[4]},
        {'label': 'Kerusakan Terbanyak', 'value': top_ker_name,
         'sub': f"{top_ker_count} ({_pct(top_ker_count)})", 'grad': card_grads[5]},
    ]
    st.markdown(kpi_html(cards), unsafe_allow_html=True)
    st.write("")

    if total_s == 0:
        st.info(f"Tidak ada data {rank_label.lower()} untuk filter yang dipilih.")
        with st.expander("ℹ️ Catatan metodologi"):
            st.write(note_text)
        return

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"#### Teknisi dengan {rank_label} Terbanyak")
        tek_rank = (sub['TEKNISI'].replace('TIDAK ADA TEKNISI', 'Belum ada teknisi')
                    .value_counts().head(10).sort_values(ascending=True))
        fig3 = go.Figure(go.Bar(
            x=tek_rank.values, y=tek_rank.index, orientation='h',
            marker_color=palette[:len(tek_rank)][::-1] if len(tek_rank) <= len(palette) else palette,
            text=[f"{v} ({v/total_s*100:.1f}%)" for v in tek_rank.values],
            textposition='outside'
        ))
        fig3.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig3, use_container_width=True, key=f"{key_prefix}_fig3")

    with col2:
        st.markdown(f"#### Kerusakan Terbanyak pada {rank_label}")
        ker_rank = sub['KERUSAKAN'].value_counts().head(10).sort_values(ascending=True)
        fig4 = go.Figure(go.Bar(
            x=ker_rank.values, y=ker_rank.index, orientation='h',
            marker_color=palette[:len(ker_rank)][::-1] if len(ker_rank) <= len(palette) else palette,
            text=[f"{v} ({v/total_s*100:.1f}%)" for v in ker_rank.values],
            textposition='outside'
        ))
        fig4.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig4, use_container_width=True, key=f"{key_prefix}_fig4")

    col3, col4, col5 = st.columns(3)
    with col3:
        st.markdown(f"###### Jenis {jenis_col_label}")
        jd = sub['JENIS'].value_counts().reset_index()
        jd.columns = ['Jenis', 'Jumlah']
        figd1 = px.pie(jd, names='Jenis', values='Jumlah', hole=0.55, color_discrete_sequence=palette)
        figd1.update_layout(height=280, margin=dict(l=5, r=5, t=5, b=5), showlegend=True,
                             legend=dict(font=dict(size=9)))
        st.plotly_chart(figd1, use_container_width=True, key=f"{key_prefix}_figd1")
    with col4:
        st.markdown("###### Per Cabang")
        cd = sub['CABANG'].value_counts().reset_index()
        cd.columns = ['Cabang', 'Jumlah']
        figd2 = px.pie(cd, names='Cabang', values='Jumlah', hole=0.55, color_discrete_sequence=palette)
        figd2.update_layout(height=280, margin=dict(l=5, r=5, t=5, b=5), showlegend=True,
                             legend=dict(font=dict(size=9)))
        st.plotly_chart(figd2, use_container_width=True, key=f"{key_prefix}_figd2")
    with col5:
        st.markdown("###### Per Teknisi (Top 5 + Lainnya)")
        td = sub['TEKNISI'].replace('TIDAK ADA TEKNISI', 'Belum ada teknisi').value_counts()
        top5 = td.head(5)
        rest = td.iloc[5:].sum()
        if rest > 0:
            top5 = pd.concat([top5, pd.Series({'Lainnya': rest})])
        figd3 = px.pie(names=top5.index, values=top5.values, hole=0.55, color_discrete_sequence=palette)
        figd3.update_layout(height=280, margin=dict(l=5, r=5, t=5, b=5), showlegend=True,
                             legend=dict(font=dict(size=9)))
        st.plotly_chart(figd3, use_container_width=True, key=f"{key_prefix}_figd3")

    st.markdown("#### Ranking Lengkap Teknisi")
    tek_full = (sub['TEKNISI'].replace('TIDAK ADA TEKNISI', 'Belum ada teknisi').value_counts().reset_index())
    tek_full.columns = ['Teknisi', f'Jumlah {rank_label}']
    tek_full['Porsi (%)'] = (tek_full[f'Jumlah {rank_label}'] / total_s * 100).round(1)
    st.dataframe(tek_full, use_container_width=True, height=300, key=f"{key_prefix}_tek_full")

    st.markdown("#### Ranking Lengkap Jenis Kerusakan")
    ker_full = sub['KERUSAKAN'].value_counts().reset_index()
    ker_full.columns = ['Kerusakan Utama', f'Jumlah {rank_label}']
    ker_full['Porsi (%)'] = (ker_full[f'Jumlah {rank_label}'] / total_s * 100).round(1)
    st.dataframe(ker_full, use_container_width=True, height=300, key=f"{key_prefix}_ker_full")

    st.markdown(f"#### Detail Transaksi {rank_label}")
    search = st.text_input("Cari teknisi / nomor / customer / kerusakan", key=f"{key_prefix}_search")
    detail_cols = ['TGL PENGIRIMAN', 'CABANG', 'TEKNISI', 'STATUS PENGERJAAN', 'NAMA CUSTOMER', 'KERUSAKAN UTAMA']
    detail_cols = [c for c in detail_cols if c in sub.columns]
    detail_df = sub[detail_cols].copy()
    if search:
        mask = detail_df.apply(lambda row: search.upper() in ' '.join(str(v) for v in row.values).upper(), axis=1)
        detail_df = detail_df[mask]
    st.dataframe(detail_df.sort_values('TGL PENGIRIMAN', ascending=False), use_container_width=True,
                 height=350, key=f"{key_prefix}_detail")

    with st.expander("ℹ️ Catatan metodologi"):
        st.write(note_text)


# ---------------------------------------------------------------------------
# Sidebar: upload file & filter
# ---------------------------------------------------------------------------
st.sidebar.title("📁 Sumber Data")
uploaded = st.sidebar.file_uploader(
    "Upload file Excel (satu sheet per cabang)",
    type=['xlsx'],
    help="Opsional. Kalau tidak upload apa pun, dashboard memakai data bawaan yang tersimpan di repo (kalau ada)."
)

if uploaded is not None:
    source_bytes = uploaded.getvalue()
    source_kind = 'xlsx'
    st.sidebar.success("📤 Memakai file yang baru diupload (berlaku untuk sesi ini saja).")
elif DEFAULT_DATA_PATH.exists():
    source_bytes = DEFAULT_DATA_PATH.read_bytes()
    source_kind = 'csv_gz'
    st.sidebar.info("📦 Memakai data bawaan yang tersimpan di repo — tidak perlu upload.")
else:
    st.title("📊 Dashboard Service Cabang")
    st.info(
        "Silakan upload file Excel data (format: satu sheet per cabang, kolom baku seperti "
        "NOMOR PENGIRIMAN PESANAN, TGL PENGIRIMAN, STATUS PENGERJAAN, NAMA TEKNISI, KERUSAKAN UTAMA, dst) "
        "lewat panel di sebelah kiri untuk mulai. Atau, tambahkan file `data/latest_data.csv.gz` ke repo "
        "supaya dashboard punya data bawaan dan tidak perlu upload setiap kali app dibuka ulang."
    )
    st.stop()

try:
    data = load_data(source_bytes, source_kind)
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
f_tahun = st.sidebar.selectbox("Tahun", tahun_opts, format_func=lambda x: str(x))
f_bulan = st.sidebar.selectbox("Bulan", bulan_opts, format_func=lambda x: x if isinstance(x, str) else BULAN_NAMES[x])
f_cabang = st.sidebar.selectbox("Cabang", cabang_opts, format_func=lambda x: str(x))

st.sidebar.markdown("---")
st.sidebar.caption(
    f"📦 {total_raw_rows:,} baris mentah → {total_unique:,} transaksi unik "
    f"({total_raw_rows - total_unique:,} duplikat dihapus)."
)

filtered = apply_filters(data, f_tahun, f_bulan, f_cabang)

# ---------------------------------------------------------------------------
# TABS
# ---------------------------------------------------------------------------
tab_main, tab_pending, tab_done, tab_cancel = st.tabs(
    ["📊 Dashboard Utama", "⚠️ Dashboard Pending", "✅ Dashboard Done", "🚫 Dashboard Cancel"]
)

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
    render_detail_dashboard(
        filtered, total_unique,
        status_bucket='PENDING',
        jenis_func=jenis_pending,
        jenis_col_label='Pending',
        palette=PALETTE_URGENT,
        banner_html="""
        <div class="warn-banner">
          ⚠️ <b>Pending bukan pencapaian — ini beban kerja yang harus segera dituntaskan dan di selesaikan.</b>
          <span>Semakin lama status tertahan, semakin besar risiko komplain dari customer. Gunakan ranking di bawah
          untuk menentukan prioritas tindak lanjut (teknisi, cabang, dan jenis kerusakan mana yang paling butuh perhatian).</span>
        </div>
        """,
        card_grads=[
            'linear-gradient(135deg,#6d3fbf,#8e3fc0)', 'linear-gradient(135deg,#9c3fc0,#c93fa8)',
            'linear-gradient(135deg,#b23f9e,#d1478d)', 'linear-gradient(135deg,#17a3a3,#159a8d)',
            'linear-gradient(135deg,#0f8a82,#0c7a6e)', 'linear-gradient(135deg,#c9392f,#e0475a)',
        ],
        rank_label='Pending',
        note_text=(
            "Yang termasuk **Pending** adalah status berisi kata PENDING serta status COMPLAIN yang belum selesai. "
            "Teknisi berlabel **Belum ada teknisi** berarti kolom Nama Teknisi kosong pada baris transaksi tsb. "
            "Rata-rata/hari dihitung dari total pending sesuai filter dibagi jumlah hari kalender bulan-bulan yang "
            "tercakup filter (bulan berjalan dipotong sampai tanggal hari ini)."
        ),
        key_prefix='pending',
    )

# =============================================================================
# TAB 3: DASHBOARD DONE
# =============================================================================
with tab_done:
    st.markdown("## Dashboard Done — Breakdown Teknisi & Kerusakan")
    render_detail_dashboard(
        filtered, total_unique,
        status_bucket='DONE',
        jenis_func=jenis_done,
        jenis_col_label='Done',
        palette=PALETTE_SUCCESS,
        banner_html="""
        <div class="success-banner">
          ✅ <b>Done adalah unit yang sudah selesai dikerjakan.</b>
          <span>Gunakan breakdown ini untuk melihat teknisi paling produktif, cabang dengan penyelesaian terbanyak,
          dan jenis kerusakan apa yang paling sering berhasil dituntaskan — bahan evaluasi kinerja & kapasitas tim.</span>
        </div>
        """,
        card_grads=[
            'linear-gradient(135deg,#16a34a,#22c55e)', 'linear-gradient(135deg,#0f8a82,#159a8d)',
            'linear-gradient(135deg,#17a3a3,#0ea5e9)', 'linear-gradient(135deg,#22c55e,#4ade80)',
            'linear-gradient(135deg,#0c7a6e,#0f8a82)', 'linear-gradient(135deg,#3fbf7f,#16a34a)',
        ],
        rank_label='Done',
        note_text=(
            "Yang termasuk **Done** adalah status berisi kata DONE (Done, Done Diambil, Done Klaim Garansi, "
            "Done Komplain, Complain Done, dst). Teknisi berlabel **Belum ada teknisi** berarti kolom Nama Teknisi "
            "kosong pada baris transaksi tsb. Rata-rata/hari dihitung dari total done sesuai filter dibagi jumlah "
            "hari kalender bulan-bulan yang tercakup filter (bulan berjalan dipotong sampai tanggal hari ini)."
        ),
        key_prefix='done',
    )

# =============================================================================
# TAB 4: DASHBOARD CANCEL
# =============================================================================
with tab_cancel:
    st.markdown("## Dashboard Cancel — Breakdown Teknisi & Kerusakan")
    render_detail_dashboard(
        filtered, total_unique,
        status_bucket='CANCEL',
        jenis_func=jenis_cancel,
        jenis_col_label='Cancel',
        palette=PALETTE_CANCEL,
        banner_html="""
        <div class="cancel-banner">
          🚫 <b>Cancel berarti kehilangan pekerjaan — dan bisa jadi sinyal ada yang perlu dibenahi.</b>
          <span>Setiap pembatalan berpotensi menandakan masalah di proses (harga, waktu tunggu, ketersediaan
          sparepart, komunikasi, dll). Gunakan ranking di bawah untuk melihat pola pembatalan yang perlu diselidiki.</span>
        </div>
        """,
        card_grads=[
            'linear-gradient(135deg,#8b1e1e,#a33131)', 'linear-gradient(135deg,#a33131,#c94f4f)',
            'linear-gradient(135deg,#992e2e,#b33f3f)', 'linear-gradient(135deg,#7a1f1f,#8b1e1e)',
            'linear-gradient(135deg,#851f35,#a83250)', 'linear-gradient(135deg,#6e1a1a,#992e2e)',
        ],
        rank_label='Cancel',
        note_text=(
            "Yang termasuk **Cancel** adalah status berisi kata CANCEL (Cancel Diambil, Cancel By Customer, "
            "Cancel By Teknisi, Cancel By Sparepart, Cancel By Admin, dst). Teknisi berlabel **Belum ada teknisi** "
            "berarti kolom Nama Teknisi kosong pada baris transaksi tsb. Rata-rata/hari dihitung dari total cancel "
            "sesuai filter dibagi jumlah hari kalender bulan-bulan yang tercakup filter (bulan berjalan dipotong "
            "sampai tanggal hari ini)."
        ),
        key_prefix='cancel',
    )
