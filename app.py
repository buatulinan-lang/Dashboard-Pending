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
import re
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

# Data faktur penjualan (omzet, modal, laba) — dipakai tab Penjualan & Voucher MLF.
DEFAULT_SALES_PATH = Path(__file__).parent / "data" / "penjualan.csv.gz"

# Nama barang untuk voucher tiket MLF (dicocokkan tanpa membedakan huruf besar/kecil)
MLF_ITEM = "VOUCHER TICKET MLF 2026"

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


SALES_REQUIRED = ['TGL FAKTUR', 'NO FAKTUR', 'KATEGORI BARANG', 'NAMA BARANG',
                  'HARGA BELI', 'QTY', '@HARGA', 'TOTAL HARGA', 'CABANG']


@st.cache_data(show_spinner="Membaca data penjualan...")
def load_sales(file_bytes: bytes, source_kind: str) -> pd.DataFrame:
    """Baca data faktur penjualan.

    Catatan penting soal kolom HARGA BELI: di sumber data ini nilainya sudah
    berupa TOTAL modal untuk baris tersebut (sudah dikali QTY), bukan harga
    satuan. Ini diverifikasi dari baris ber-QTY>1, mis. voucher MLF dengan
    HARGA BELI 34.000 untuk QTY 2 — sama dengan 17.000/unit pada baris QTY 1.
    Karena itu MODAL = HARGA BELI (tidak dikalikan QTY lagi).
    """
    if source_kind == 'csv_gz':
        df = pd.read_csv(io.BytesIO(file_bytes), compression='gzip')
    else:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        frames = []
        for sheet in xls.sheet_names:
            d = xls.parse(sheet)
            if d.empty:
                continue
            d['CABANG'] = sheet
            frames.append(d)
        if not frames:
            return pd.DataFrame()
        df = pd.concat(frames, ignore_index=True, sort=False)

    if df.empty:
        return df

    missing = [c for c in SALES_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(
            "Kolom berikut tidak ditemukan di data penjualan: " + ", ".join(missing)
        )

    for c in ['HARGA BELI', 'QTY', '@HARGA', 'TOTAL HARGA']:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    df['TGL'] = pd.to_datetime(df['TGL FAKTUR'], errors='coerce')
    df['TAHUN'] = df['TGL'].dt.year
    df['BULAN'] = df['TGL'].dt.month
    df['MODAL'] = df['HARGA BELI']
    df['LABA'] = df['TOTAL HARGA'] - df['MODAL']
    df['KATEGORI'] = (df['KATEGORI BARANG'].astype(str).str.strip().str.upper()
                      .replace({'NAN': 'TIDAK ADA DATA', '': 'TIDAK ADA DATA'}))
    df['BARANG'] = df['NAMA BARANG'].astype(str).str.strip()
    df['BARANG_U'] = df['BARANG'].str.upper()
    if 'YANG MENYERAHKAN/MENJUAL' in df.columns:
        df['PENJUAL'] = (df['YANG MENYERAHKAN/MENJUAL'].astype(str).str.strip()
                         .replace({'nan': 'TIDAK ADA DATA', '': 'TIDAK ADA DATA'}))
    else:
        df['PENJUAL'] = 'TIDAK ADA DATA'

    # Nama teknisi: pakai kolom (FINAL) sebagai acuan karena sudah dikoreksi,
    # jatuh ke kolom NAMA TEKNISI bila kosong.
    fin = df['NAMA TEKNISI (FINAL)'] if 'NAMA TEKNISI (FINAL)' in df.columns else pd.Series(index=df.index, dtype=object)
    asli = df['NAMA TEKNISI'] if 'NAMA TEKNISI' in df.columns else pd.Series(index=df.index, dtype=object)
    tek = fin.fillna(asli)
    df['TEKNISI'] = (tek.astype(str).str.strip().str.upper()
                     .replace({'NAN': '', 'NONE': ''}))
    df.loc[df['TEKNISI'] == '', 'TEKNISI'] = 'TIDAK ADA TEKNISI'

    # Simpan kata kunci yang cocok saja. Tarifnya sengaja TIDAK dihitung di sini
    # supaya bisa diubah pengguna dari dashboard tanpa memuat ulang data.
    is_jasa = df['KATEGORI'] == 'JASA'
    df['KW_MATCH'] = ''
    if is_jasa.any():
        df.loc[is_jasa, 'KW_MATCH'] = (df.loc[is_jasa, 'BARANG']
                                       .map(lambda s: '|'.join(cocok_kata_kunci(s))))
    return df


# --- aturan bagi hasil jasa teknisi -----------------------------------------
# Tarif ditentukan dari kata kunci pada NAMA BARANG. Semua angka di bawah hanya
# NILAI AWAL — pengguna bisa mengubahnya langsung dari tab dashboard.
KATA_KUNCI_TARIF = ['INTERFACE', 'NORMAL', 'MATI TOTAL', 'PROMO']
TARIF_AWAL = {
    'Interface': 20.0,
    'Normal': 30.0,
    'Mati Total': 32.0,
    'Promo': 60.0,
}
TARIF_DEFAULT_AWAL = 30.0     # item jasa tanpa kata kunci (mis. "JASA REPAIR")
TARIF_PEMBANDING_AWAL = 30.0  # skema pembanding: seluruh omzet jasa x tarif ini
LABEL_LAINNYA = 'Lainnya'


def cocok_kata_kunci(nama_barang):
    """Daftar kata kunci yang terkandung di satu nama barang (urut tetap)."""
    s = str(nama_barang).upper()
    return [k for k in KATA_KUNCI_TARIF if k in s]


def pilih_label_tarif(kw_str, urutan):
    """Tentukan label tarif dari kata kunci yang cocok, mengikuti urutan prioritas."""
    if not kw_str:
        return LABEL_LAINNYA
    cocok = kw_str.split('|')
    for k in urutan:
        if k in cocok:
            return k.title()
    return cocok[0].title()


def periode_gaji(bulan_gaji: int, tahun_gaji: int):
    """Rentang tanggal cutoff untuk satu bulan penggajian.

    Aturan: gaji bulan M dihitung dari 24 bulan (M-2) sampai 23 bulan (M-1).
    Contoh: gaji Mei 2026 -> 24 Maret 2026 s/d 23 April 2026.
    """
    m_akhir = bulan_gaji - 1
    th_akhir = tahun_gaji
    if m_akhir < 1:
        m_akhir += 12
        th_akhir -= 1
    m_awal = m_akhir - 1
    th_awal = th_akhir
    if m_awal < 1:
        m_awal += 12
        th_awal -= 1
    return (pd.Timestamp(th_awal, m_awal, 24),
            pd.Timestamp(th_akhir, m_akhir, 23))


def label_periode(bulan_gaji: int, tahun_gaji: int):
    a, b = periode_gaji(bulan_gaji, tahun_gaji)
    return (f"Gaji {BULAN_NAMES[bulan_gaji]} {tahun_gaji}  "
            f"({a.day} {BULAN_NAMES[a.month]} – {b.day} {BULAN_NAMES[b.month]} {b.year})")


def daftar_periode_gaji(tgl_min, tgl_max):
    """Semua bulan penggajian yang periodenya beririsan dengan rentang data."""
    hasil = []
    if pd.isna(tgl_min) or pd.isna(tgl_max):
        return hasil
    y, m = tgl_min.year, tgl_min.month
    for _ in range(60):
        # majukan bulan gaji sampai periodenya melewati data
        a, b = periode_gaji(m, y)
        if a > tgl_max:
            break
        if b >= tgl_min:
            hasil.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return hasil


# Warna kartu KPI pada laporan PDF. Sengaja berupa kode hex biasa supaya
# dashboard tetap jalan walau pustaka pembuat PDF belum terpasang.
_PN = "#1F3864"   # navy
_PG = "#16A34A"   # hijau
_PR = "#C0392B"   # merah
_PA = "#D97706"   # oranye
_PM = "#6B7280"   # abu-abu


def nfid(v, desimal=0):
    """Angka gaya Indonesia: 68.838 · 1.234,5"""
    try:
        s = f"{float(v):,.{desimal}f}"
    except (TypeError, ValueError):
        return str(v)
    return s.replace(",", "#").replace(".", ",").replace("#", ".")


def pctid(v, desimal=1):
    """Persen gaya Indonesia: 84,0%"""
    try:
        return f"{float(v):.{desimal}f}%".replace(".", ",")
    except (TypeError, ValueError):
        return str(v)


def rp(v, singkat=True):
    """Format rupiah ringkas: 1,2 M / 340,5 jt / 12.500."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "-"
    neg = v < 0
    v = abs(v)
    if singkat:
        if v >= 1_000_000_000:
            s = f"Rp {v/1_000_000_000:,.2f} M"
        elif v >= 1_000_000:
            s = f"Rp {v/1_000_000:,.1f} jt"
        elif v >= 1_000:
            s = f"Rp {v/1_000:,.0f} rb"
        else:
            s = f"Rp {v:,.0f}"
    else:
        s = f"Rp {v:,.0f}"
    s = s.replace(",", "#").replace(".", ",").replace("#", ".")
    return ("-" + s) if neg else s


# ---------------------------------------------------------------------------
# Pembanding periode (bulan ini vs bulan lalu, tahun ini vs tahun lalu)
# ---------------------------------------------------------------------------
def potong_periode(df: pd.DataFrame, kol_tgl: str):
    """Siapkan potongan data yang setara untuk perbandingan MoM & YoY.

    Bulan berjalan biasanya belum penuh, jadi bulan pembanding ikut dipotong
    pada tanggal yang sama. Hari terakhir dibuang bila datanya jauh di bawah
    kebiasaan (tanda data belum lengkap sehari penuh).
    """
    if df is None or df.empty or kol_tgl not in df.columns:
        return None
    t = df[kol_tgl].dropna()
    if t.empty:
        return None

    tmax = t.max()
    th, bl = int(tmax.year), int(tmax.month)
    hb = t[(t.dt.year == th) & (t.dt.month == bl)].dt.day.value_counts().sort_index()
    hari_n = int(tmax.day)
    dibuang = None
    if len(hb) >= 3 and hb.get(hari_n, 0) < 0.4 * hb.iloc[:-1].median():
        dibuang = hari_n
        hari_n -= 1

    bl_p, th_p = (bl - 1, th) if bl > 1 else (12, th - 1)
    tg = df[kol_tgl]

    def rng(bulan, tahun):
        return df[(tg.dt.year == tahun) & (tg.dt.month == bulan) & (tg.dt.day <= hari_n)]

    # YTD: 1 Januari s/d tanggal & bulan yang sama
    def ytd(tahun):
        return df[(tg.dt.year == tahun) &
                  ((tg.dt.month < bl) |
                   ((tg.dt.month == bl) & (tg.dt.day <= hari_n)))]

    return {
        'bulan_ini': rng(bl, th), 'bulan_lalu': rng(bl_p, th_p),
        'tahun_ini': ytd(th), 'tahun_lalu': ytd(th - 1),
        'nama_bulan_ini': f"{BULAN_NAMES[bl]} {th}",
        'nama_bulan_lalu': f"{BULAN_NAMES[bl_p]} {th_p}",
        'nama_tahun_ini': str(th), 'nama_tahun_lalu': str(th - 1),
        'hari_n': hari_n, 'hari_dibuang': dibuang, 'bulan': bl, 'tahun': th,
    }


def _delta_html(baru, lama, fmt, higher_better=True, satuan=""):
    """Satu baris perbandingan: nilai baru, nilai lama, dan selisihnya."""
    if lama in (None, 0) or pd.isna(lama):
        badge = "<span style='color:#6b7280'>—</span>"
    else:
        p = (baru - lama) / abs(lama) * 100
        naik = p > 0
        baik = naik if higher_better else (not naik)
        warna = '#16a34a' if (baik and abs(p) > 0.05) else ('#dc2626' if abs(p) > 0.05 else '#6b7280')
        panah = '▲' if naik else ('▼' if p < 0 else '■')
        badge = (f"<span style='color:{warna};font-weight:700'>{panah} "
                 f"{abs(p):.1f}%".replace('.', ',') + "</span>")
    return (f"<td style='padding:6px 10px;font-weight:700'>{fmt(baru)}{satuan}</td>"
            f"<td style='padding:6px 10px;color:#6b7280'>{fmt(lama)}{satuan}</td>"
            f"<td style='padding:6px 10px;text-align:right'>{badge}</td>")


def hitung_banding(pot, metrik):
    """Hitung isi tabel perbandingan (dipakai layar maupun PDF).

    Kembalikan list dict siap pakai untuk pdf_export.build_pdf.
    """
    PG, PR, PM = _PG, _PR, _PM
    hasil = []
    if not pot:
        return hasil

    def susun(judul, lama_df, baru_df, nama_baru, nama_lama, catatan):
        baris = []
        for nama, fn, fmt, hb in metrik:
            try:
                v_baru, v_lama = fn(baru_df), fn(lama_df)
            except Exception:
                continue
            if v_lama in (None, 0) or pd.isna(v_lama):
                teks, warna = "—", PM
            else:
                p = (v_baru - v_lama) / abs(v_lama) * 100
                naik = p > 0
                baik = naik if hb else (not naik)
                warna = PG if (baik and abs(p) > 0.05) else (PR if abs(p) > 0.05 else PM)
                teks = f"{'+' if naik else ''}{p:.1f}%".replace('.', ',')
            baris.append((nama, fmt(v_baru), fmt(v_lama), teks, warna))
        return {'judul': judul, 'baris': baris, 'nama_baru': nama_baru,
                'nama_lama': nama_lama, 'catatan': catatan}

    cat_mom = (f"Dibandingkan setara: tanggal 1-{pot['hari_n']} pada kedua bulan."
               + (f" Tanggal {pot['hari_dibuang']} dikecualikan karena datanya belum "
                  "lengkap sehari penuh." if pot['hari_dibuang'] else ""))
    hasil.append(susun("Bulan Ini vs Bulan Lalu", pot['bulan_lalu'], pot['bulan_ini'],
                       pot['nama_bulan_ini'], pot['nama_bulan_lalu'], cat_mom))

    if not pot['tahun_lalu'].empty:
        cat_yoy = (f"Dibandingkan setara: 1 Januari - {pot['hari_n']} "
                   f"{BULAN_NAMES[pot['bulan']]} pada kedua tahun.")
        hasil.append(susun("Tahun Ini vs Tahun Lalu", pot['tahun_lalu'], pot['tahun_ini'],
                           pot['nama_tahun_ini'], pot['nama_tahun_lalu'], cat_yoy))
    else:
        hasil.append({'judul': 'Tahun Ini vs Tahun Lalu', 'baris': [],
                      'nama_baru': pot['nama_tahun_ini'],
                      'nama_lama': pot['nama_tahun_lalu'],
                      'catatan': f"Tidak bisa dibandingkan - data "
                                 f"{pot['nama_tahun_lalu']} tidak tersedia."})
    return hasil


def tombol_pdf(nama_dashboard, pot, metrik, temuan, kpis=None, metodologi="",
               ringkasan="", key=""):
    """Tampilkan tombol unduh laporan analisa dalam bentuk PDF."""
    try:
        from pdf_export import build_pdf
    except ModuleNotFoundError as e:
        hilang = getattr(e, "name", "") or str(e)
        if hilang == "pdf_export":
            st.info("📄 Unduhan PDF belum aktif — file **pdf_export.py** belum ada "
                    "di folder yang sama dengan app.py.")
        else:
            st.info(f"📄 Unduhan PDF belum aktif — pustaka **{hilang}** belum terpasang. "
                    f"Tambahkan `reportlab>=4.0` ke **requirements.txt**, lalu reboot "
                    f"aplikasi. (Dashboard lain tetap berjalan normal.)")
        return
    except Exception as e:  # noqa: BLE001
        st.caption(f"Unduhan PDF tidak tersedia: {e}")
        return

    periode = ("Semua Periode" if f_tahun == 'Semua Tahun' and f_bulan == 'Semua Bulan'
               else (f"Tahun {f_tahun}" if f_bulan == 'Semua Bulan'
                     else (f"{BULAN_NAMES[int(f_bulan)]} (semua tahun)"
                           if f_tahun == 'Semua Tahun'
                           else f"{BULAN_NAMES[int(f_bulan)]} {f_tahun}")))
    cab_txt = "Seluruh Cabang" if f_cabang == 'Semua Cabang' else f"Cabang {f_cabang}"

    try:
        data = build_pdf(
            judul=nama_dashboard, periode=periode, cabang=cab_txt,
            kpis=kpis or [], banding=hitung_banding(pot, metrik),
            temuan=temuan, metodologi=metodologi, ringkasan=ringkasan,
            penyusun=st.session_state.get('ppt_penyusun', ''))
    except Exception as e:  # noqa: BLE001
        st.caption(f"Gagal menyusun PDF: {e}")
        return

    nm = re.sub(r'[^A-Za-z0-9]+', '_', nama_dashboard).strip('_')
    tag = f"{f_tahun}" if f_tahun != 'Semua Tahun' else 'semua'
    st.download_button(
        "📄 Unduh Analisa (PDF)", data=data,
        file_name=f"Analisa_{nm}_{tag}.pdf", mime="application/pdf",
        key=f"pdf_{key}", use_container_width=False)


def render_banding(pot, metrik, key_prefix=""):
    """Tampilkan tabel perbandingan MoM & YoY.

    metrik: list of (nama, fungsi(df)->angka, formatter, higher_better)
    """
    if not pot:
        st.caption("Data tanggal tidak tersedia untuk perbandingan periode.")
        return

    def tabel(judul, a_df, b_df, nama_a, nama_b, catatan):
        rows = []
        for nama, fn, fmt, hb in metrik:
            try:
                va, vb = fn(b_df), fn(a_df)   # b_df = periode lama
            except Exception:
                continue
            rows.append(
                f"<tr><td style='padding:6px 10px;color:#20242e'>{nama}</td>"
                + _delta_html(va, vb, fmt, hb) + "</tr>")
        if not rows:
            return ""
        return f"""
        <div style="background:#fff;border:1px solid #e3e7f0;border-radius:12px;
                    padding:12px 14px;height:100%">
          <div style="font-weight:800;color:#1f3864;font-size:13px;margin-bottom:8px">{judul}</div>
          <table style="width:100%;border-collapse:collapse;font-size:12.5px">
            <tr style="color:#6b7280;font-size:10.5px;text-transform:uppercase;
                       letter-spacing:.03em">
              <th style="text-align:left;padding:4px 10px">Metrik</th>
              <th style="text-align:left;padding:4px 10px">{nama_a}</th>
              <th style="text-align:left;padding:4px 10px">{nama_b}</th>
              <th style="text-align:right;padding:4px 10px">Selisih</th>
            </tr>
            {''.join(rows)}
          </table>
          <div style="font-size:10.5px;color:#6b7280;margin-top:8px">{catatan}</div>
        </div>"""

    cat_mom = (f"Dibandingkan setara: tanggal 1–{pot['hari_n']} pada kedua bulan."
               + (f" Tanggal {pot['hari_dibuang']} dikecualikan karena datanya belum "
                  "lengkap sehari penuh." if pot['hari_dibuang'] else ""))
    cat_yoy = (f"Dibandingkan setara: 1 Januari – {pot['hari_n']} "
               f"{BULAN_NAMES[pot['bulan']]} pada kedua tahun.")

    kosong_yoy = pot['tahun_lalu'].empty
    c1, c2 = st.columns(2)
    with c1:
        html = tabel("📅 Bulan Ini vs Bulan Lalu", pot['bulan_lalu'], pot['bulan_ini'],
                     pot['nama_bulan_ini'], pot['nama_bulan_lalu'], cat_mom)
        st.markdown(html, unsafe_allow_html=True)
    with c2:
        if kosong_yoy:
            st.markdown(
                f"""<div style="background:#fffaf0;border:1px dashed #f0d9a8;
                     border-radius:12px;padding:14px;height:100%">
                  <div style="font-weight:800;color:#7a5b18;font-size:13px">
                    📆 Tahun Ini vs Tahun Lalu</div>
                  <div style="font-size:12px;color:#7a5b18;margin-top:8px;line-height:1.6">
                    Tidak bisa dibandingkan — data {pot['nama_tahun_lalu']} tidak tersedia
                    pada sumber ini. Perbandingan antar tahun baru bisa dilakukan setelah
                    ada data periode sebelumnya.</div>
                </div>""", unsafe_allow_html=True)
        else:
            html = tabel("📆 Tahun Ini vs Tahun Lalu", pot['tahun_lalu'], pot['tahun_ini'],
                         pot['nama_tahun_ini'], pot['nama_tahun_lalu'], cat_yoy)
            st.markdown(html, unsafe_allow_html=True)


def panel_analisa(items):
    """Kotak 'Analisa & Tindak Lanjut'. items: list of (jenis, judul, isi).

    jenis: 'baik' | 'perhatian' | 'aksi' | 'info'
    """
    warna = {
        'baik':      ('#f2fbf5', '#b9e8c9', '#0f5132', '✅'),
        'perhatian': ('#fff8ec', '#f0d9a8', '#7a5b18', '⚠️'),
        'aksi':      ('#fdf3f2', '#ebcfcb', '#7a2a24', '🎯'),
        'info':      ('#f7f9fd', '#e3e7f0', '#1f3864', 'ℹ️'),
    }
    blok = []
    for jenis, judul, isi in items:
        bg, br, tx, ik = warna.get(jenis, warna['info'])
        blok.append(f"""
        <div style="background:{bg};border:1px solid {br};border-radius:12px;
                    padding:12px 14px;margin-bottom:10px">
          <div style="font-weight:800;color:{tx};font-size:13px;margin-bottom:4px">
            {ik} {judul}</div>
          <div style="font-size:12.5px;color:{tx};line-height:1.65">{isi}</div>
        </div>""")
    st.markdown("".join(blok), unsafe_allow_html=True)


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


KELOMPOK_UMUR = [
    (0, 7, "0–7 hari", "#16a34a"),
    (8, 14, "8–14 hari", "#65b83f"),
    (15, 30, "15–30 hari", "#e0b31f"),
    (31, 60, "31–60 hari", "#e0921f"),
    (61, 90, "61–90 hari", "#e0651f"),
    (91, 10 ** 6, "> 90 hari", "#c0392b"),
]


def kelompok_umur(hari):
    for a, b, label, _ in KELOMPOK_UMUR:
        if a <= hari <= b:
            return label
    return KELOMPOK_UMUR[-1][2]


def render_umur_pending(sub, semua_df, key_prefix):
    """Tampilkan sebaran umur unit yang tertahan.

    Umur dihitung dari TGL PENGIRIMAN sampai tanggal data terakhir (bukan
    tanggal hari ini), supaya angkanya tidak ikut membengkak bila dashboard
    dibuka lama setelah data terakhir ditarik.
    """
    if 'TGL PENGIRIMAN' not in sub.columns:
        return None
    s = sub.dropna(subset=['TGL PENGIRIMAN']).copy()
    if s.empty:
        return None

    acuan = semua_df['TGL PENGIRIMAN'].max()
    if pd.isna(acuan):
        return None
    s['UMUR'] = (acuan - s['TGL PENGIRIMAN']).dt.days.clip(lower=0)
    s['KEL_UMUR'] = s['UMUR'].map(kelompok_umur)

    n = len(s)
    med = s['UMUR'].median()
    n30 = int((s['UMUR'] > 30).sum())
    n90 = int((s['UMUR'] > 90).sum())
    tertua = s.loc[s['UMUR'].idxmax()]

    st.markdown("#### ⏳ Umur Pendingan")
    st.caption(
        f"Dihitung dari tanggal masuk sampai tanggal data terakhir "
        f"(**{acuan:%d %B %Y}**). Jumlah pending saja belum cukup — yang "
        f"menentukan risiko komplain adalah berapa lama unit tertahan."
    )
    st.markdown(kpi_html([
        {'label': 'Median Umur', 'value': f"{nfid(med)} hari",
         'sub': f"separuh pendingan di bawah ini",
         'grad': 'linear-gradient(135deg,#1f3864,#2e5394)'},
        {'label': 'Lebih dari 30 Hari', 'value': nfid(n30),
         'sub': f"{pctid(n30/n*100 if n else 0)} dari pendingan",
         'grad': 'linear-gradient(135deg,#e0921f,#e2b21a)'},
        {'label': 'Lebih dari 90 Hari', 'value': nfid(n90),
         'sub': f"{pctid(n90/n*100 if n else 0)} — sudah kritis",
         'grad': 'linear-gradient(135deg,#c0392b,#e0475a)'},
        {'label': 'Unit Tertua', 'value': f"{nfid(tertua['UMUR'])} hari",
         'sub': f"{tertua['CABANG']} · {tertua['TGL PENGIRIMAN']:%d/%m/%Y}",
         'grad': 'linear-gradient(135deg,#7c3aed,#a855f7)'},
        {'label': 'Rata-rata Umur', 'value': f"{nfid(s['UMUR'].mean())} hari",
         'sub': "lebih tinggi dari median = ada ekor panjang",
         'grad': 'linear-gradient(135deg,#0f8a82,#17a3a3)'},
    ]), unsafe_allow_html=True)
    st.write("")

    u1, u2 = st.columns([1, 1.15])
    with u1:
        st.markdown("###### Sebaran Kelompok Umur")
        urut = [k[2] for k in KELOMPOK_UMUR]
        warna = {k[2]: k[3] for k in KELOMPOK_UMUR}
        vc = s['KEL_UMUR'].value_counts().reindex(urut, fill_value=0)
        figu = go.Figure(go.Bar(
            x=vc.index, y=vc.values,
            marker_color=[warna[i] for i in vc.index],
            text=[f"{v}" for v in vc.values], textposition='outside'))
        figu.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10),
                           yaxis_title='Jumlah unit')
        st.plotly_chart(figu, use_container_width=True, key=f'{key_prefix}_umur_bar')
    with u2:
        st.markdown("###### Umur Pendingan per Cabang")
        tab = pd.crosstab(s['CABANG'], s['KEL_UMUR'])
        for k in urut:
            if k not in tab.columns:
                tab[k] = 0
        tab = tab[urut]
        tab['Total'] = tab.sum(axis=1)
        tab['> 30 hari'] = s[s['UMUR'] > 30].groupby('CABANG').size().reindex(
            tab.index, fill_value=0)
        tab['% > 30 hari'] = (tab['> 30 hari'] / tab['Total'] * 100).round(1)
        tab = tab.sort_values('> 30 hari', ascending=False)
        st.dataframe(tab, use_container_width=True, height=320,
                     key=f'{key_prefix}_umur_cab')

    st.markdown("###### 20 Unit Paling Lama Tertahan")
    kol = ['TGL PENGIRIMAN', 'UMUR', 'CABANG', 'TEKNISI', 'KERUSAKAN',
           'STATUS PENGERJAAN', 'NAMA CUSTOMER']
    kol = [c for c in kol if c in s.columns]
    lama = s.nlargest(20, 'UMUR')[kol].rename(columns={
        'TGL PENGIRIMAN': 'Tgl Masuk', 'UMUR': 'Umur (hari)', 'CABANG': 'Cabang',
        'TEKNISI': 'Teknisi', 'KERUSAKAN': 'Kerusakan',
        'STATUS PENGERJAAN': 'Status', 'NAMA CUSTOMER': 'Customer'})
    st.dataframe(lama, use_container_width=True, height=340, hide_index=True,
                 key=f'{key_prefix}_umur_lama')

    csv_umur = s[kol + ['KEL_UMUR']].sort_values('UMUR', ascending=False)
    st.download_button(
        "⬇️ Unduh daftar pendingan beserta umurnya (CSV)",
        data=csv_umur.to_csv(index=False).encode('utf-8-sig'),
        file_name="pendingan_dengan_umur.csv", mime="text/csv",
        key=f'{key_prefix}_umur_unduh')

    return {'n': n, 'median': med, 'rata': s['UMUR'].mean(), 'n30': n30,
            'n90': n90, 'tertua': tertua, 'acuan': acuan,
            'per_cabang': s[s['UMUR'] > 30]['CABANG'].value_counts(),
            'per_kerusakan': s[s['UMUR'] > 30]['KERUSAKAN'].value_counts()
            if 'KERUSAKAN' in s.columns else pd.Series(dtype=int)}


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

    # ---------- umur pendingan ----------
    # Jumlah pending saja belum cukup: unit yang tertahan 3 hari itu operasi
    # normal, sedangkan yang tertahan 200 hari sudah jadi aset mati sekaligus
    # komplain yang tinggal menunggu waktu.
    umur_info = None
    if status_bucket == 'PENDING':
        umur_info = render_umur_pending(sub, filtered_df, key_prefix)

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

    # ---------- perbandingan periode ----------
    st.markdown("### 📊 Perbandingan Periode")
    _pot = potong_periode(filtered_df, 'TGL PENGIRIMAN')
    _sb = status_bucket
    _metrik_d = [
        (f"Jumlah {rank_label}", lambda x: int((x['STATUS_BUCKET'] == _sb).sum()),
         lambda v: f"{v:,.0f}", (_sb == 'DONE')),
        (f"% {rank_label} dari total",
         lambda x: (x['STATUS_BUCKET'] == _sb).mean() * 100 if len(x) else 0,
         lambda v: f"{v:.1f}%".replace('.', ','), (_sb == 'DONE')),
        ("Total transaksi", lambda x: len(x), lambda v: nfid(v), True),
    ]
    render_banding(_pot, _metrik_d, key_prefix=key_prefix)

    # ---------- analisa otomatis ----------
    st.markdown("### 🧭 Analisa & Tindak Lanjut")
    _an = []
    if _pot and not _pot['bulan_ini'].empty and not _pot['bulan_lalu'].empty:
        n_i = int((_pot['bulan_ini']['STATUS_BUCKET'] == _sb).sum())
        n_l = int((_pot['bulan_lalu']['STATUS_BUCKET'] == _sb).sum())
        p_i = (_pot['bulan_ini']['STATUS_BUCKET'] == _sb).mean() * 100
        p_l = (_pot['bulan_lalu']['STATUS_BUCKET'] == _sb).mean() * 100
        if _sb == 'PENDING':
            _an.append(('info', 'Cara membaca angka pending',
                        "Pending adalah <b>kondisi terkini</b>, bukan kejadian pada bulan "
                        "tersebut. Pendingan bulan lalu sebagian sudah tuntas sehingga "
                        "selalu tampak lebih kecil — jadi kenaikan angka di sini "
                        "<b>bukan berarti kinerja memburuk</b>. Yang perlu dilihat adalah "
                        "besarnya tumpukan saat ini dan di mana konsentrasinya."))
        elif n_l:
            selisih = (n_i - n_l) / n_l * 100
            baik = (selisih > 0) if _sb == 'DONE' else (selisih < 0)
            _an.append(('baik' if baik else 'perhatian',
                        f'{rank_label} {"naik" if selisih > 0 else "turun"} '
                        f'{abs(selisih):.1f}% dibanding bulan lalu',
                        f"Dari {n_l:,} menjadi {n_i:,} unit (porsi {p_l:.1f}% → {p_i:.1f}%). "
                        + ("<b>Tindakan:</b> pertahankan dan bakukan praktik yang berjalan."
                           if baik else
                           "<b>Tindakan:</b> telusuri penyebabnya pada cabang dengan "
                           "perubahan terbesar.")))

    if total_s:
        tek_r = sub[~sub['TEKNISI'].isin(['TIDAK ADA TEKNISI', 'N/A'])]['TEKNISI'].value_counts()
        cab_r = sub['CABANG'].value_counts()
        ker_r = sub['KERUSAKAN'].value_counts()
        if _sb == 'PENDING' and umur_info:
            ui = umur_info
            if ui['n30']:
                cab30 = ui['per_cabang']
                ker30 = ui['per_kerusakan']
                _an.append((
                    'aksi', f"{nfid(ui['n30'])} unit tertahan lebih dari 30 hari",
                    f"Itu {pctid(ui['n30']/ui['n']*100)} dari seluruh pendingan, dan "
                    f"{nfid(ui['n90'])} di antaranya sudah lewat 90 hari. "
                    + (f"Terbanyak di <b>{cab30.index[0]}</b> ({nfid(cab30.iloc[0])} unit)"
                       if len(cab30) else "")
                    + (f", kerusakan <b>{ker30.index[0]}</b>" if len(ker30) else "")
                    + ". <b>Tindakan:</b> unit setua ini biasanya menunggu keputusan "
                      "pelanggan atau sparepart yang tak kunjung ada — hubungi "
                      "pelanggannya untuk keputusan lanjut/ambil, jangan dibiarkan "
                      "menumpuk di rak."))
            if ui['rata'] > ui['median'] * 2:
                _an.append((
                    'perhatian', 'Sebagian kecil unit menyeret rata-rata jauh ke atas',
                    f"Median umur {nfid(ui['median'])} hari tapi rata-ratanya "
                    f"{nfid(ui['rata'])} hari — artinya mayoritas pendingan sebenarnya "
                    f"masih wajar, hanya ada segelintir unit sangat tua "
                    f"(tertua {nfid(ui['tertua']['UMUR'])} hari di "
                    f"{ui['tertua']['CABANG']}). <b>Tindakan:</b> selesaikan ekor "
                    f"panjang ini lebih dulu; jumlahnya sedikit tapi dampaknya pada "
                    f"citra layanan paling besar."))

        if _sb == 'PENDING':
            dua = cab_r.head(2)
            porsi = dua.sum() / total_s * 100
            _an.append(('aksi', 'Prioritaskan dua cabang ini',
                        f"<b>{dua.index[0]}</b> ({dua.iloc[0]} unit) dan "
                        f"<b>{dua.index[1]}</b> ({dua.iloc[1]} unit) menyumbang "
                        f"{porsi:.1f}% dari seluruh pendingan. <b>Tindakan:</b> "
                        f"penelusuran di dua titik ini memberi dampak terbesar."))
            _an.append(('aksi', f'Kerusakan {ker_r.index[0]} paling menumpuk',
                        f"{ker_r.iloc[0]} unit ({ker_r.iloc[0]/total_s*100:.1f}% pendingan). "
                        f"<b>Tindakan:</b> cek ketersediaan sparepart dan kejelasan "
                        f"estimasi biaya untuk jenis kerusakan ini."))
        elif _sb == 'CANCEL':
            _an.append(('aksi', f'Kerusakan {ker_r.index[0]} paling sering dibatalkan',
                        f"{ker_r.iloc[0]} unit ({ker_r.iloc[0]/total_s*100:.1f}% pembatalan). "
                        f"<b>Tindakan:</b> sampaikan estimasi biaya lebih awal untuk "
                        f"jenis ini agar pelanggan tidak menunggu lama sebelum membatalkan."))
            _an.append(('perhatian', f'Pembatalan terbanyak di {cab_r.index[0]}',
                        f"{cab_r.iloc[0]} unit ({cab_r.iloc[0]/total_s*100:.1f}%). "
                        f"<b>Tindakan:</b> bandingkan harga dan waktu tunggu cabang ini "
                        f"dengan cabang yang pembatalannya rendah."))
        else:  # DONE
            if len(tek_r):
                _an.append(('baik', 'Teknisi paling produktif',
                            f"<b>{tek_r.index[0]}</b> menyelesaikan {tek_r.iloc[0]:,} unit. "
                            f"<b>Tindakan:</b> pelajari cara kerjanya untuk dijadikan "
                            f"acuan pelatihan teknisi lain."))
            _an.append(('info', 'Beban kerja terbesar',
                        f"Kerusakan <b>{ker_r.index[0]}</b> menyumbang "
                        f"{ker_r.iloc[0]/total_s*100:.1f}% penyelesaian. "
                        f"<b>Tindakan:</b> pastikan stok sparepart jenis ini aman."))

        tanpa = (sub['TEKNISI'] == 'TIDAK ADA TEKNISI').sum()
        if tanpa > total_s * 0.05:
            _an.append(('perhatian', 'Banyak baris tanpa nama teknisi',
                        f"{tanpa:,} baris ({tanpa/total_s*100:.1f}%) tidak tercatat "
                        f"teknisinya. <b>Tindakan:</b> rapikan pengisian di sumber data "
                        f"agar penilaian kinerja per teknisi akurat."))

    if not _an:
        _an.append(('info', 'Belum ada temuan menonjol',
                    "Tidak ada pola mencolok pada filter ini."))
    panel_analisa(_an)

    _wrn = {'DONE': _PG, 'CANCEL': _PR, 'PENDING': _PA}.get(_sb, _PN)
    _porsi = (total_s / len(filtered_df) * 100) if len(filtered_df) else 0
    _kp = [{'label': f'Jumlah {rank_label}', 'value': nfid(total_s),
            'sub': f"{pctid(_porsi)} dari total", 'warna': _wrn}]
    if total_s:
        _kp += [{'label': 'Teknisi Terbanyak', 'value': str(top_tek_name)[:20],
                 'sub': (f"{nfid(top_tek_count)} unit" if top_tek_count else '-'),
                 'warna': _PN},
                {'label': 'Cabang Terbanyak', 'value': str(top_cabang_name)[:16],
                 'sub': f"{nfid(top_cabang_count)} unit", 'warna': _PN},
                {'label': 'Kerusakan Terbanyak', 'value': str(top_ker_name)[:18],
                 'sub': f"{nfid(top_ker_count)} unit", 'warna': _PR}]
    if umur_info:
        _kp += [{'label': 'Median Umur', 'value': f"{nfid(umur_info['median'])} hari",
                 'sub': 'separuh pendingan di bawah ini', 'warna': _PN},
                {'label': 'Tertahan > 30 Hari', 'value': nfid(umur_info['n30']),
                 'sub': f"{pctid(umur_info['n30']/umur_info['n']*100)} dari pendingan",
                 'warna': _PA},
                {'label': 'Tertahan > 90 Hari', 'value': nfid(umur_info['n90']),
                 'sub': 'sudah kritis', 'warna': _PR}]
    tombol_pdf(f"Dashboard {rank_label}", _pot, _metrik_d, _an, kpis=_kp,
               ringkasan=(f"{nfid(total_s)} transaksi berstatus {rank_label.lower()} "
                          f"({pctid(_porsi)} dari seluruh transaksi pada filter ini)."),
               metodologi=note_text, key=key_prefix)

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

# --- data penjualan (opsional; kalau tidak ada, tab Penjualan & MLF nonaktif) ---
sales_uploaded = st.sidebar.file_uploader(
    "Upload data penjualan (opsional)",
    type=['xlsx', 'gz', 'csv'],
    help="Format faktur penjualan: satu sheet per cabang, atau file penjualan.csv.gz.",
    key="sales_upload",
)
sales = pd.DataFrame()
sales_err = ""
try:
    if sales_uploaded is not None:
        _kind = 'csv_gz' if sales_uploaded.name.endswith(('.gz', '.csv')) else 'xlsx'
        sales = load_sales(sales_uploaded.getvalue(), _kind)
        st.sidebar.success("💰 Data penjualan dari file yang diupload.")
    elif DEFAULT_SALES_PATH.exists():
        sales = load_sales(DEFAULT_SALES_PATH.read_bytes(), 'csv_gz')
        st.sidebar.info("💰 Data penjualan bawaan repo dimuat.")
except Exception as e:  # noqa: BLE001
    sales_err = str(e)
    st.sidebar.warning(f"Data penjualan tidak terbaca: {e}")

tahun_src = set(int(t) for t in data['TAHUN'].dropna().unique())
bulan_src = set(int(b) for b in data['BULAN'].dropna().unique())
cabang_src = set(data['CABANG'].dropna().unique().tolist())
if not sales.empty:
    tahun_src |= set(int(t) for t in sales['TAHUN'].dropna().unique())
    bulan_src |= set(int(b) for b in sales['BULAN'].dropna().unique())
    cabang_src |= set(sales['CABANG'].dropna().unique().tolist())

tahun_opts = ['Semua Tahun'] + sorted(tahun_src)
bulan_opts = ['Semua Bulan'] + sorted(bulan_src)
cabang_opts = ['Semua Cabang'] + sorted(cabang_src)

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
sales_f = apply_filters(sales, f_tahun, f_bulan, f_cabang) if not sales.empty else sales

# ---------------------------------------------------------------------------
# Sidebar: buat PPT dari dashboard
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.title("📤 Buat Presentasi")
st.sidebar.caption(
    "Membuat file PowerPoint mengikuti template perusahaan, berisi data sesuai "
    "filter yang sedang aktif."
)

_penyusun = st.sidebar.text_input(
    "Nama penyusun (opsional)",
    placeholder="cth: ARM MFLASH – BUDIARJA IBRAHIM",
    key="ppt_penyusun",
)
_SLIDE_LABELS = {
    "ringkasan": "Ringkasan Kinerja",
    "status": "Komposisi Status",
    "progres_bulan": "Progres Bulanan",
    "banding_bulan": "Bulan Ini vs Bulan Lalu",
    "harian": "Rekap Transaksi Harian",
    "hari_tertinggi": "Hari Transaksi Tertinggi",
    "cabang": "Kinerja per Cabang",
    "pending": "Detail Pending",
    "done": "Detail Done",
    "cancel": "Detail Cancel",
    "penjualan": "Penjualan (Modal & Laba)",
    "mlf": "Voucher Tiket MLF",
    "bagihasil": "Bagi Hasil Teknisi",
    "penutup": "Kesimpulan",
}
_slide_pilihan = st.sidebar.multiselect(
    "Slide yang disertakan",
    options=list(_SLIDE_LABELS.keys()),
    default=list(_SLIDE_LABELS.keys()),
    format_func=lambda x: _SLIDE_LABELS[x],
    key="ppt_slides",
)

# periode penggajian khusus slide Bagi Hasil Teknisi
if not sales.empty and 'bagihasil' in _slide_pilihan:
    _jasa_src = sales[sales['KATEGORI'] == 'JASA']
    _per_list = (daftar_periode_gaji(_jasa_src['TGL'].min(), _jasa_src['TGL'].max())
                 if not _jasa_src.empty else [])
    if _per_list:
        _opsi_bh = ['Ikuti filter di atas'] + _per_list
        _pilih_bh = st.sidebar.selectbox(
            "Periode gaji (slide Bagi Hasil)", _opsi_bh,
            index=len(_opsi_bh) - 1,
            format_func=lambda x: (x if isinstance(x, str)
                                   else label_periode(x[1], x[0])),
            key='ppt_periode_pilih',
            help="Cutoff tanggal 24 s/d 23. Pilih 'Ikuti filter di atas' agar "
                 "memakai filter Tahun/Bulan biasa.")
        st.session_state['ppt_periode_bh'] = (
            None if isinstance(_pilih_bh, str) else _pilih_bh)

if st.sidebar.button("🎬 Buat PPT dari Dashboard", use_container_width=True, type="primary"):
    if filtered.empty:
        st.sidebar.error("Tidak ada data pada filter aktif — ganti filter dulu.")
    else:
        try:
            # pastikan folder app ikut dicari saat import (beberapa host tidak
            # otomatis menaruh folder skrip di sys.path)
            import sys
            _app_dir = str(Path(__file__).parent.resolve())
            if _app_dir not in sys.path:
                sys.path.insert(0, _app_dir)

            from pptx_export import build_deck

            # tarif bagi hasil mengikuti isian di tab Bagi Hasil Teknisi;
            # kalau tab itu belum pernah dibuka, dipakai nilai awal.
            _tarif_map = {
                'Interface': st.session_state.get('tar_int', TARIF_AWAL['Interface']),
                'Normal': st.session_state.get('tar_nor', TARIF_AWAL['Normal']),
                'Mati Total': st.session_state.get('tar_mat', TARIF_AWAL['Mati Total']),
                'Promo': st.session_state.get('tar_pro', TARIF_AWAL['Promo']),
                'Lainnya': st.session_state.get('tar_lain', TARIF_DEFAULT_AWAL),
            }
            _tarif_flat = st.session_state.get('tar_flat', TARIF_PEMBANDING_AWAL)
            _prioritas = st.session_state.get('tar_prio', 'Normal')
            _periode_bh = st.session_state.get('ppt_periode_bh')

            with st.spinner("Menyusun presentasi..."):
                _ppt_bytes = build_deck(
                    filtered,
                    total_unique_all=total_unique,
                    total_raw_rows=total_raw_rows,
                    f_tahun=f_tahun, f_bulan=f_bulan, f_cabang=f_cabang,
                    penyusun=_penyusun,
                    sertakan=tuple(_slide_pilihan) if _slide_pilihan else ("ringkasan",),
                    sales_filtered=(sales_f if not sales_f.empty else None),
                    tarif_map=_tarif_map,
                    tarif_flat=_tarif_flat,
                    prioritas=_prioritas,
                    periode_bagihasil=_periode_bh,
                )
            st.session_state["ppt_bytes"] = _ppt_bytes
            _tag_t = "SemuaTahun" if f_tahun == "Semua Tahun" else str(f_tahun)
            _tag_b = "" if f_bulan == "Semua Bulan" else f"-{BULAN_NAMES[int(f_bulan)]}"
            _tag_c = "SemuaCabang" if f_cabang == "Semua Cabang" else str(f_cabang)
            st.session_state["ppt_name"] = f"Laporan_Service_{_tag_t}{_tag_b}_{_tag_c}.pptx"
            st.sidebar.success("Presentasi siap diunduh.")
        except ModuleNotFoundError as e:
            _missing = getattr(e, "name", "") or str(e)
            if _missing == "pptx_export":
                _here = sorted(p.name for p in Path(__file__).parent.iterdir())
                st.sidebar.error(
                    "File **pptx_export.py** tidak ditemukan di folder yang sama "
                    "dengan app.py.\n\nIsi folder saat ini: " + ", ".join(_here)
                )
            elif _missing in ("pptx", "python-pptx"):
                st.sidebar.error(
                    "Library **python-pptx** belum terpasang.\n\n"
                    "Tambahkan baris `python-pptx>=1.0` ke file **requirements.txt** "
                    "di repo, lalu reboot aplikasi. Kalau menjalankan di komputer "
                    "sendiri: `pip install python-pptx`."
                )
            else:
                st.sidebar.error(
                    f"Library **{_missing}** belum terpasang. Tambahkan ke "
                    "requirements.txt lalu reboot aplikasi."
                )
        except Exception as e:  # noqa: BLE001
            st.sidebar.error(f"Gagal membuat PPT: {type(e).__name__}: {e}")

if st.session_state.get("ppt_bytes"):
    st.sidebar.download_button(
        "⬇️ Unduh file PPT",
        data=st.session_state["ppt_bytes"],
        file_name=st.session_state.get("ppt_name", "Laporan_Service.pptx"),
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )

# ---------------------------------------------------------------------------
# TABS
# ---------------------------------------------------------------------------
(tab_main, tab_pending, tab_done, tab_cancel, tab_mati, tab_jual, tab_mlf,
 tab_tek, tab_bundling) = st.tabs(
    ["📊 Dashboard Utama", "⚠️ Dashboard Pending", "✅ Dashboard Done",
     "🚫 Dashboard Cancel", "🔌 Mati Total", "💰 Penjualan", "🎫 Voucher MLF",
     "🧰 Omzet & Bagi Hasil Teknisi", "🎁 Bundling Aksesoris"]
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

    # ---------- perbandingan periode ----------
    st.markdown("### 📊 Perbandingan Periode")
    _pot = potong_periode(filtered, 'TGL PENGIRIMAN')
    _metrik_utama = [
        ("Total transaksi", lambda x: len(x), lambda v: nfid(v), True),
        ("Selesai (Done)", lambda x: int((x['STATUS_BUCKET'] == 'DONE').sum()),
         lambda v: nfid(v), True),
        ("Batal (Cancel)", lambda x: int((x['STATUS_BUCKET'] == 'CANCEL').sum()),
         lambda v: nfid(v), False),
        ("% Penyelesaian", lambda x: (x['STATUS_BUCKET'] == 'DONE').mean() * 100 if len(x) else 0,
         lambda v: pctid(v), True),
        ("% Pembatalan", lambda x: (x['STATUS_BUCKET'] == 'CANCEL').mean() * 100 if len(x) else 0,
         lambda v: pctid(v), False),
    ]
    render_banding(_pot, _metrik_utama, key_prefix='main')

    # ---------- analisa otomatis ----------
    st.markdown("### 🧭 Analisa & Tindak Lanjut")
    _an = []
    if _pot and not _pot['bulan_ini'].empty and not _pot['bulan_lalu'].empty:
        bi, bl_ = _pot['bulan_ini'], _pot['bulan_lalu']
        g_vol = (len(bi) - len(bl_)) / len(bl_) * 100 if len(bl_) else 0
        dr_i = (bi['STATUS_BUCKET'] == 'DONE').mean() * 100
        dr_l = (bl_['STATUS_BUCKET'] == 'DONE').mean() * 100
        cr_i = (bi['STATUS_BUCKET'] == 'CANCEL').mean() * 100
        cr_l = (bl_['STATUS_BUCKET'] == 'CANCEL').mean() * 100

        # Transaksi bulan berjalan belum sempat selesai semua, sehingga %Done
        # bulan terbaru SELALU tampak lebih rendah. Bedakan efek ini dari
        # penurunan kinerja yang sesungguhnya.
        blm_i = 100 - dr_i - cr_i          # porsi belum tuntas (pending + lainnya)
        blm_l = 100 - dr_l - cr_l
        efek_matang = (blm_i - blm_l) > 3

        if efek_matang:
            _an.append(('info', 'Penurunan %Done bulan ini sebagian besar semu',
                        f"Porsi transaksi yang belum tuntas naik dari {blm_l:.1f}% ke "
                        f"{blm_i:.1f}%, karena pekerjaan {_pot['nama_bulan_ini']} memang "
                        f"belum sempat diselesaikan semua saat data ditarik. "
                        f"Jadi turunnya %Done dari {dr_l:.1f}% ke {dr_i:.1f}% "
                        f"<b>belum tentu berarti kinerja memburuk</b>. Yang lebih andal "
                        f"dibandingkan adalah <b>jumlah transaksi masuk</b> dan "
                        f"<b>jumlah unit selesai</b>, bukan persentasenya."))

        if g_vol > 5 and dr_i < dr_l - 1 and not efek_matang:
            _an.append(('aksi', 'Volume naik tapi penyelesaian tertinggal',
                        f"Transaksi naik {g_vol:.1f}% dibanding {_pot['nama_bulan_lalu']}, "
                        f"namun tingkat penyelesaian justru turun dari {dr_l:.1f}% ke "
                        f"{dr_i:.1f}%, dan itu <b>bukan</b> sekadar efek pekerjaan yang "
                        f"belum matang. <b>Tindakan:</b> cek antrean di cabang dengan "
                        f"kenaikan tertinggi dan pertimbangkan penambahan shift atau "
                        f"pengalihan teknisi sementara."))
        elif g_vol > 5:
            _an.append(('baik', 'Volume tumbuh dengan mutu terjaga',
                        f"Transaksi naik {g_vol:.1f}% dan penyelesaian bertahan di "
                        f"{dr_i:.1f}%. <b>Tindakan:</b> pertahankan pola kerja saat ini; "
                        f"jadikan cabang berkinerja terbaik sebagai acuan."))
        elif g_vol < -5:
            _an.append(('perhatian', 'Volume transaksi menurun',
                        f"Transaksi turun {abs(g_vol):.1f}% dibanding "
                        f"{_pot['nama_bulan_lalu']}. <b>Tindakan:</b> telusuri apakah "
                        f"karena musiman, berkurangnya promosi, atau persaingan di "
                        f"cabang tertentu."))

        if cr_i > cr_l + 1 and not efek_matang:
            _an.append(('aksi', 'Pembatalan meningkat',
                        f"Porsi pembatalan naik dari {cr_l:.1f}% ke {cr_i:.1f}%. "
                        f"<b>Tindakan:</b> periksa kecepatan pemberian estimasi biaya dan "
                        f"lama waktu tunggu, dua hal yang paling sering memicu batal."))
        elif cr_i < cr_l - 1 and not efek_matang:
            _an.append(('baik', 'Pembatalan menurun',
                        f"Porsi pembatalan turun dari {cr_l:.1f}% ke {cr_i:.1f}%. "
                        f"<b>Tindakan:</b> catat perubahan proses yang berjalan bulan ini "
                        f"agar bisa dibakukan ke cabang lain."))

        # pembanding yang lebih tahan efek pematangan: jumlah unit selesai
        n_done_i = int((bi['STATUS_BUCKET'] == 'DONE').sum())
        n_done_l = int((bl_['STATUS_BUCKET'] == 'DONE').sum())
        if n_done_l:
            gd = (n_done_i - n_done_l) / n_done_l * 100
            _an.append(('baik' if gd >= 0 else 'perhatian',
                        f"Unit selesai {'naik' if gd >= 0 else 'turun'} {abs(gd):.1f}% "
                        f"(pembanding yang lebih andal)",
                        f"{n_done_l:,} → {n_done_i:,} unit pada rentang tanggal yang sama. "
                        f"Angka mutlak seperti ini tidak terpengaruh pekerjaan yang belum "
                        f"matang, sehingga lebih layak dipakai menilai kinerja bulanan."))

    # cabang yang perlu perhatian
    if len(filtered):
        _rk = filtered.groupby('CABANG')['STATUS_BUCKET'].agg(
            n='size', done=lambda s: (s == 'DONE').mean() * 100,
            cancel=lambda s: (s == 'CANCEL').mean() * 100)
        _rk = _rk[_rk['n'] >= 100]
        if len(_rk) >= 3:
            _rata_done = (filtered['STATUS_BUCKET'] == 'DONE').mean() * 100
            _buruk = _rk.nsmallest(2, 'done')
            _nm = ", ".join(f"{i} ({r.done:.1f}%)" for i, r in _buruk.iterrows())
            _an.append(('perhatian', 'Cabang dengan penyelesaian terendah',
                        f"{_nm} — di bawah rata-rata {_rata_done:.1f}%. "
                        f"<b>Tindakan:</b> tinjau beban per teknisi dan ketersediaan "
                        f"sparepart di cabang tersebut."))
            _tinggi = _rk.nlargest(1, 'cancel')
            for i, r in _tinggi.iterrows():
                if r.cancel > filtered['STATUS_BUCKET'].eq('CANCEL').mean() * 100 + 3:
                    _an.append(('aksi', f'Pembatalan tertinggi di {i}',
                                f"Pembatalan {r.cancel:.1f}%, jauh di atas rata-rata "
                                f"{filtered['STATUS_BUCKET'].eq('CANCEL').mean()*100:.1f}%. "
                                f"<b>Tindakan:</b> wawancarai admin cabang soal alasan "
                                f"pelanggan membatalkan."))

    # pending menumpuk
    _pd = filtered[filtered['STATUS_BUCKET'] == 'PENDING']
    if len(_pd):
        _pc = _pd['CABANG'].value_counts()
        _pk = _pd['KERUSAKAN'].value_counts()
        _an.append(('aksi', f'{len(_pd):,} unit masih tertahan',
                    f"Terbanyak di <b>{_pc.index[0]}</b> ({_pc.iloc[0]} unit) dengan "
                    f"kerusakan dominan <b>{_pk.index[0]}</b> ({_pk.iloc[0]} unit). "
                    f"<b>Tindakan:</b> jadikan dua titik ini prioritas penyelesaian "
                    f"pekan ini sebelum berkembang jadi komplain."))

    if not _an:
        _an.append(('info', 'Belum ada temuan menonjol',
                    "Angka pada filter ini relatif stabil. Perluas rentang filter "
                    "atau bandingkan antar cabang untuk melihat pola yang lebih jelas."))
    panel_analisa(_an)

    tombol_pdf(
        "Dashboard Status Pengerjaan", _pot, _metrik_utama, _an,
        kpis=[{'label': 'Total Transaksi', 'value': nfid(total), 'sub': 'sesuai filter', 'warna': _PN},
              {'label': 'Selesai (Done)', 'value': nfid(done), 'sub': pct(done), 'warna': _PG},
              {'label': 'Batal (Cancel)', 'value': nfid(cancel), 'sub': pct(cancel), 'warna': _PR},
              {'label': 'Pending', 'value': nfid(pending), 'sub': pct(pending), 'warna': _PA},
              {'label': 'Rata-rata / Hari', 'value': nfid(avg_day, 1),
               'sub': f"{period_days} hari", 'warna': _PN}],
        ringkasan=(f"Dari {nfid(total)} transaksi pada periode ini, {nfid(done)} selesai "
                   f"({pct(done)}), {nfid(cancel)} batal ({pct(cancel)}), dan {nfid(pending)} "
                   f"masih tertahan."),
        metodologi=("Baris dengan seluruh kolom identik dihitung sebagai 1 transaksi. "
                    "Done = status mengandung DONE; Pending = mengandung PENDING atau "
                    "COMPLAIN; Cancel = mengandung CANCEL. Filter Tahun/Bulan mengacu "
                    "kolom TGL PENGIRIMAN."),
        key='utama')

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
          ⚠️ <b>Pending bukan pencapaian — ini beban kerja yang harus segera dituntaskan.</b>
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

# =============================================================================
# TAB 5: PENJUALAN (harga modal vs harga jual)
# =============================================================================
with tab_jual:
    st.markdown("## Dashboard Penjualan — Modal, Omzet & Laba Kotor")

    if sales.empty:
        st.info(
            "Data penjualan belum tersedia. Tambahkan file `data/penjualan.csv.gz` ke repo, "
            "atau upload lewat panel kiri (**Upload data penjualan**)."
            + (f"\n\nPesan: {sales_err}" if sales_err else "")
        )
    elif sales_f.empty:
        st.warning("Tidak ada data penjualan untuk filter yang dipilih.")
    else:
        sj = sales_f
        omzet = sj['TOTAL HARGA'].sum()
        modal = sj['MODAL'].sum()
        laba = sj['LABA'].sum()
        margin = (laba / omzet * 100) if omzet else 0
        # Nomor faktur berjalan sendiri-sendiri di tiap cabang (satu nomor bisa
        # dipakai beberapa cabang), jadi satu nota = kombinasi cabang + nomor.
        n_faktur = sj.groupby(['CABANG', 'NO FAKTUR']).ngroups
        qty = sj['QTY'].sum()

        st.markdown(kpi_html([
            {'label': 'Omzet (Harga Jual)', 'value': rp(omzet), 'sub': f"{len(sj):,} baris faktur",
             'grad': 'linear-gradient(135deg,#1f3864,#2e5394)'},
            {'label': 'Modal (Harga Beli)', 'value': rp(modal),
             'sub': f"{(modal/omzet*100 if omzet else 0):.1f}% dari omzet",
             'grad': 'linear-gradient(135deg,#c9392f,#e0475a)'},
            {'label': 'Laba Kotor', 'value': rp(laba), 'sub': f"margin {margin:.1f}%",
             'grad': 'linear-gradient(135deg,#16a34a,#22c55e)'},
            {'label': 'Jumlah Faktur', 'value': f"{n_faktur:,}",
             'sub': f"rata-rata {rp(omzet/n_faktur if n_faktur else 0)}/faktur",
             'grad': 'linear-gradient(135deg,#7c3aed,#a855f7)'},
            {'label': 'Unit Terjual', 'value': f"{qty:,.0f}",
             'sub': f"laba {rp(laba/qty if qty else 0)}/unit",
             'grad': 'linear-gradient(135deg,#0f8a82,#17a3a3)'},
            {'label': 'Baris Rugi', 'value': f"{(sj['LABA'] < 0).sum():,}",
             'sub': f"{(sj['LABA'] < 0).mean()*100:.1f}% dari baris",
             'grad': 'linear-gradient(135deg,#e0921f,#e2b21a)'},
        ]), unsafe_allow_html=True)
        st.write("")

        c1, c2 = st.columns([1.25, 1])
        with c1:
            st.markdown("#### Tren Omzet, Modal & Laba per Bulan")
            tren = (sj.groupby('BULAN')[['TOTAL HARGA', 'MODAL', 'LABA']].sum()
                    .sort_index())
            figt = go.Figure()
            figt.add_bar(x=[BULAN_NAMES[int(b)][:3] for b in tren.index],
                         y=tren['MODAL'], name='Modal', marker_color='#c9392f')
            figt.add_bar(x=[BULAN_NAMES[int(b)][:3] for b in tren.index],
                         y=tren['LABA'], name='Laba', marker_color='#16a34a')
            figt.update_layout(barmode='stack', height=350,
                               margin=dict(l=10, r=10, t=10, b=10),
                               legend=dict(orientation='h', y=1.12),
                               yaxis_title='Rupiah')
            st.plotly_chart(figt, use_container_width=True, key='jual_tren')
            st.caption("Tinggi batang = omzet (modal + laba).")
        with c2:
            st.markdown("#### Komposisi Omzet per Kategori")
            kd = (sj.groupby('KATEGORI')['TOTAL HARGA'].sum()
                  .sort_values(ascending=False).head(8).reset_index())
            kd.columns = ['Kategori', 'Omzet']
            figk = px.pie(kd, names='Kategori', values='Omzet', hole=0.55,
                          color_discrete_sequence=PALETTE)
            figk.update_layout(height=350, margin=dict(l=5, r=5, t=5, b=5),
                               legend=dict(font=dict(size=9)))
            st.plotly_chart(figk, use_container_width=True, key='jual_pie')

        st.markdown("#### Rekap per Kategori Barang")
        gk = sj.groupby('KATEGORI').agg(
            Qty=('QTY', 'sum'), Omzet=('TOTAL HARGA', 'sum'),
            Modal=('MODAL', 'sum'), Laba=('LABA', 'sum'), Baris=('QTY', 'size'))
        gk['Margin %'] = (gk['Laba'] / gk['Omzet'] * 100).round(1)
        gk = gk.sort_values('Omzet', ascending=False)
        st.dataframe(gk.style.format({'Qty': '{:,.0f}', 'Omzet': 'Rp {:,.0f}',
                                      'Modal': 'Rp {:,.0f}', 'Laba': 'Rp {:,.0f}',
                                      'Baris': '{:,.0f}'}),
                     use_container_width=True, key='jual_kat')

        st.markdown("#### Rekap per Cabang")
        gc = sj.groupby('CABANG').agg(
            Qty=('QTY', 'sum'), Omzet=('TOTAL HARGA', 'sum'),
            Modal=('MODAL', 'sum'), Laba=('LABA', 'sum'),
            Faktur=('NO FAKTUR', 'nunique'))
        gc['Margin %'] = (gc['Laba'] / gc['Omzet'] * 100).round(1)
        gc = gc.sort_values('Omzet', ascending=False)
        cc1, cc2 = st.columns([1, 1.1])
        with cc1:
            st.dataframe(gc.style.format({'Qty': '{:,.0f}', 'Omzet': 'Rp {:,.0f}',
                                          'Modal': 'Rp {:,.0f}', 'Laba': 'Rp {:,.0f}',
                                          'Faktur': '{:,.0f}'}),
                         use_container_width=True, height=420, key='jual_cab')
        with cc2:
            gcs = gc.sort_values('Omzet', ascending=True)
            figc = go.Figure()
            figc.add_bar(y=gcs.index, x=gcs['Modal'], orientation='h',
                         name='Modal', marker_color='#c9392f')
            figc.add_bar(y=gcs.index, x=gcs['Laba'], orientation='h',
                         name='Laba', marker_color='#16a34a')
            figc.update_layout(barmode='stack', height=420,
                               margin=dict(l=10, r=10, t=10, b=10),
                               legend=dict(orientation='h', y=1.06),
                               xaxis_title='Rupiah')
            st.plotly_chart(figc, use_container_width=True, key='jual_cab_fig')

        b1, b2 = st.columns(2)
        with b1:
            st.markdown("#### 15 Barang — Omzet Tertinggi")
            tb = (sj.groupby('BARANG').agg(Qty=('QTY', 'sum'),
                                           Omzet=('TOTAL HARGA', 'sum'),
                                           Laba=('LABA', 'sum'))
                  .sort_values('Omzet', ascending=False).head(15))
            tb['Margin %'] = (tb['Laba'] / tb['Omzet'] * 100).round(1)
            st.dataframe(tb.style.format({'Qty': '{:,.0f}', 'Omzet': 'Rp {:,.0f}',
                                          'Laba': 'Rp {:,.0f}'}),
                         use_container_width=True, height=380, key='jual_brg1')
        with b2:
            st.markdown("#### 15 Barang — Laba Tertinggi")
            tl = (sj.groupby('BARANG').agg(Qty=('QTY', 'sum'),
                                           Omzet=('TOTAL HARGA', 'sum'),
                                           Laba=('LABA', 'sum'))
                  .sort_values('Laba', ascending=False).head(15))
            tl['Margin %'] = (tl['Laba'] / tl['Omzet'] * 100).round(1)
            st.dataframe(tl.style.format({'Qty': '{:,.0f}', 'Omzet': 'Rp {:,.0f}',
                                          'Laba': 'Rp {:,.0f}'}),
                         use_container_width=True, height=380, key='jual_brg2')

        st.markdown("#### Detail Faktur")
        q = st.text_input("Cari barang / faktur / customer / kategori", key='jual_cari')
        cols = ['TGL FAKTUR', 'NO FAKTUR', 'CABANG', 'KATEGORI', 'NAMA BARANG',
                'HARGA BELI', 'QTY', '@HARGA', 'TOTAL HARGA', 'LABA']
        cols = [c for c in cols if c in sj.columns]
        dd = sj[cols].copy()
        if q:
            mask = dd.apply(lambda r: q.upper() in ' '.join(str(v) for v in r.values).upper(), axis=1)
            dd = dd[mask]
        st.caption(f"{len(dd):,} baris ditampilkan (maksimal 1.000).")
        st.dataframe(dd.head(1000), use_container_width=True, height=360, key='jual_detail')

        # ---------- perbandingan periode ----------
        st.markdown("### 📊 Perbandingan Periode")
        _potj = potong_periode(sj, 'TGL')
        _metrik_j = [
            ("Omzet", lambda x: x['TOTAL HARGA'].sum(), lambda v: rp(v), True),
            ("Modal", lambda x: x['MODAL'].sum(), lambda v: rp(v), False),
            ("Laba kotor", lambda x: x['LABA'].sum(), lambda v: rp(v), True),
            ("Margin", lambda x: (x['LABA'].sum() / x['TOTAL HARGA'].sum() * 100)
             if x['TOTAL HARGA'].sum() else 0,
             lambda v: pctid(v), True),
            ("Jumlah nota", lambda x: x.groupby(['CABANG', 'NO FAKTUR']).ngroups,
             lambda v: nfid(v), True),
        ]
        render_banding(_potj, _metrik_j, key_prefix='jual')

        st.markdown("### 🧭 Analisa & Tindak Lanjut")
        _anj = []
        if _potj and not _potj['bulan_ini'].empty and not _potj['bulan_lalu'].empty:
            bi, bl_ = _potj['bulan_ini'], _potj['bulan_lalu']
            o_i, o_l = bi['TOTAL HARGA'].sum(), bl_['TOTAL HARGA'].sum()
            m_i = bi['LABA'].sum() / o_i * 100 if o_i else 0
            m_l = bl_['LABA'].sum() / o_l * 100 if o_l else 0
            g = (o_i - o_l) / o_l * 100 if o_l else 0
            if g > 3 and m_i < m_l - 1:
                _anj.append(('aksi', 'Omzet naik tapi margin tergerus',
                             f"Omzet naik {g:.1f}% namun margin turun dari {m_l:.1f}% ke "
                             f"{m_i:.1f}%. Kemungkinan komposisi bergeser ke barang "
                             f"bermargin tipis. <b>Tindakan:</b> tinjau harga jual "
                             f"kategori bervolume besar dan dorong item bermargin tinggi."))
            elif g > 3:
                _anj.append(('baik', 'Omzet tumbuh sehat',
                             f"Omzet naik {g:.1f}% dengan margin {m_i:.1f}%. "
                             f"<b>Tindakan:</b> pertahankan bauran produk saat ini."))
            elif g < -3:
                _anj.append(('perhatian', 'Omzet menurun',
                             f"Omzet turun {abs(g):.1f}% dibanding "
                             f"{_potj['nama_bulan_lalu']}. <b>Tindakan:</b> periksa cabang "
                             f"penyumbang penurunan terbesar."))

        _gk2 = sj.groupby('KATEGORI').agg(om=('TOTAL HARGA', 'sum'), lb=('LABA', 'sum'))
        _gk2 = _gk2[_gk2['om'] > 0]
        _gk2['mg'] = _gk2['lb'] / _gk2['om'] * 100
        _brg = _gk2.drop(index=['JASA'], errors='ignore')
        if len(_brg):
            _low = _brg[_brg['om'] >= _brg['om'].max() * 0.1].nsmallest(1, 'mg')
            for i, r in _low.iterrows():
                _anj.append(('perhatian', f'Margin {i} paling tipis',
                             f"Hanya {r.mg:.1f}% dari omzet {rp(r.om)}. "
                             f"<b>Tindakan:</b> negosiasi ulang harga beli atau sesuaikan "
                             f"harga jual — kategori ini bervolume besar sehingga "
                             f"perbaikan kecil berdampak besar."))
        if 'JASA' in _gk2.index:
            _pj = _gk2.loc['JASA', 'om'] / sj['TOTAL HARGA'].sum() * 100
            _anj.append(('info', 'Hati-hati membaca margin JASA',
                         f"JASA menyumbang {_pj:.1f}% omzet dengan margin tampil "
                         f"~100% karena biaya tenaga kerja tidak dibebankan per faktur. "
                         f"<b>Jangan</b> dibandingkan langsung dengan margin barang."))

        _rug = (sj['LABA'] < 0).sum()
        if _rug:
            _anj.append(('aksi', f'{_rug:,} baris terjual rugi',
                         f"{_rug/len(sj)*100:.1f}% baris punya harga jual di bawah modal "
                         f"(total {rp(sj.loc[sj['LABA'] < 0, 'LABA'].sum())}). "
                         f"<b>Tindakan:</b> cek apakah salah input harga atau memang "
                         f"diskon berlebih."))
        if not _anj:
            _anj.append(('info', 'Belum ada temuan menonjol', "Angka relatif stabil."))
        panel_analisa(_anj)

        tombol_pdf("Dashboard Penjualan", _potj, _metrik_j, _anj,
                   kpis=[{'label': 'Omzet', 'value': rp(omzet), 'sub': f"{nfid(len(sj))} baris", 'warna': _PN},
                         {'label': 'Modal', 'value': rp(modal),
                          'sub': f"{pctid(modal/omzet*100 if omzet else 0)} omzet", 'warna': _PR},
                         {'label': 'Laba Kotor', 'value': rp(laba),
                          'sub': f"margin {pctid(margin)}", 'warna': _PG},
                         {'label': 'Jumlah Nota', 'value': nfid(n_faktur),
                          'sub': f"rata-rata {rp(omzet/n_faktur if n_faktur else 0)}", 'warna': _PN},
                         {'label': 'Unit Terjual', 'value': nfid(qty),
                          'sub': f"laba {rp(laba/qty if qty else 0)}/unit", 'warna': _PA}],
                   ringkasan=(f"Omzet {rp(omzet)} dengan modal {rp(modal)} menghasilkan "
                              f"laba kotor {rp(laba)} (margin {pctid(margin)})."),
                   metodologi=("Modal dari kolom HARGA BELI yang sudah berupa total per "
                               "baris. Laba kotor = TOTAL HARGA - MODAL, belum dikurangi "
                               "biaya operasional. Kategori JASA bermodal nol sehingga "
                               "marginnya tampil 100%. Satu nota = kombinasi Cabang + "
                               "Nomor Faktur."),
                   key='jual')

        with st.expander("ℹ️ Catatan metodologi"):
            st.write(
                "**Modal** diambil dari kolom `HARGA BELI`, yang di sumber data ini sudah berupa "
                "**total** untuk baris tersebut (sudah dikali QTY) — bukan harga satuan. Ini "
                "diverifikasi dari baris ber-QTY lebih dari 1; contohnya voucher MLF dengan "
                "HARGA BELI 34.000 untuk QTY 2, konsisten dengan 17.000/unit pada baris QTY 1. "
                "Kalau modal dianggap harga satuan, total modal jadi 3x omzet dan margin menjadi "
                "minus 197% — jelas keliru.\n\n"
                "**Laba kotor** = TOTAL HARGA − MODAL, jadi belum dikurangi biaya operasional, "
                "gaji, sewa, dan lainnya.\n\n"
                "Baris berkategori **JASA** hampir semuanya bermodal 0, sehingga marginnya tampil "
                "100%. Itu wajar karena biaya tenaga kerja tidak dibebankan per faktur — perlu "
                "diingat saat membandingkan margin antar kategori.\n\n"
                "Semua baris dihitung apa adanya tanpa pembuangan duplikat, sesuai kesepakatan "
                "(di data ini duplikat persis hanya sekitar 90 dari 172.791 baris)."
            )

# =============================================================================
# TAB 6: VOUCHER TIKET MLF
# =============================================================================
with tab_mlf:
    st.markdown("## Dashboard Voucher Tiket MLF — Rekap Semua Cabang")

    if sales.empty:
        st.info(
            "Data penjualan belum tersedia. Tambahkan file `data/penjualan.csv.gz` ke repo, "
            "atau upload lewat panel kiri."
        )
    else:
        mlf_all = sales[sales['BARANG_U'].str.contains('MLF', na=False)]
        mlf = sales_f[sales_f['BARANG_U'].str.contains('MLF', na=False)] if not sales_f.empty \
            else sales_f

        if mlf_all.empty:
            st.warning("Tidak ditemukan barang yang mengandung kata 'MLF' pada data penjualan.")
        elif mlf.empty:
            st.warning("Tidak ada penjualan voucher MLF untuk filter yang dipilih.")
        else:
            qty = mlf['QTY'].sum()
            omzet = mlf['TOTAL HARGA'].sum()
            modal = mlf['MODAL'].sum()
            laba = mlf['LABA'].sum()
            margin = (laba / omzet * 100) if omzet else 0
            n_cab = mlf['CABANG'].nunique()
            hari = mlf['TGL'].dt.normalize().nunique()

            st.markdown(kpi_html([
                {'label': 'Voucher Terjual', 'value': f"{qty:,.0f}", 'sub': f"{len(mlf):,} transaksi",
                 'grad': 'linear-gradient(135deg,#1f3864,#2e5394)'},
                {'label': 'Omzet', 'value': rp(omzet),
                 'sub': f"rata-rata {rp(omzet/qty if qty else 0)}/voucher",
                 'grad': 'linear-gradient(135deg,#2e9bd6,#17a3a3)'},
                {'label': 'Modal', 'value': rp(modal),
                 'sub': f"{(modal/omzet*100 if omzet else 0):.1f}% dari omzet",
                 'grad': 'linear-gradient(135deg,#c9392f,#e0475a)'},
                {'label': 'Laba Kotor', 'value': rp(laba), 'sub': f"margin {margin:.1f}%",
                 'grad': 'linear-gradient(135deg,#16a34a,#22c55e)'},
                {'label': 'Cabang Menjual', 'value': f"{n_cab}",
                 'sub': f"dari {sales['CABANG'].nunique()} cabang",
                 'grad': 'linear-gradient(135deg,#7c3aed,#a855f7)'},
                {'label': 'Rata-rata / Hari', 'value': f"{(qty/hari if hari else 0):,.1f}",
                 'sub': f"{hari} hari ada penjualan",
                 'grad': 'linear-gradient(135deg,#e0921f,#e2b21a)'},
            ]), unsafe_allow_html=True)
            st.write("")

            m1, m2 = st.columns([1.3, 1])
            with m1:
                st.markdown("#### Tren Penjualan Voucher per Hari")
                dly = mlf.groupby(mlf['TGL'].dt.normalize())['QTY'].sum().sort_index()
                figd = go.Figure(go.Scatter(
                    x=dly.index, y=dly.values, mode='lines+markers',
                    line=dict(color='#1f3864', width=2), marker=dict(size=4),
                    name='Voucher'))
                figd.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10),
                                   yaxis_title='Jumlah voucher')
                st.plotly_chart(figd, use_container_width=True, key='mlf_daily')
                if len(dly):
                    st.caption(
                        f"Tertinggi: {int(dly.max())} voucher pada "
                        f"{dly.idxmax().strftime('%d %B %Y')}."
                    )
            with m2:
                st.markdown("#### Porsi Penjualan per Cabang")
                pc = mlf.groupby('CABANG')['QTY'].sum().sort_values(ascending=False)
                top6 = pc.head(6)
                if len(pc) > 6:
                    top6 = pd.concat([top6, pd.Series({'Lainnya': pc.iloc[6:].sum()})])
                figp = px.pie(names=top6.index, values=top6.values, hole=0.55,
                              color_discrete_sequence=PALETTE)
                figp.update_layout(height=340, margin=dict(l=5, r=5, t=5, b=5),
                                   legend=dict(font=dict(size=9)))
                st.plotly_chart(figp, use_container_width=True, key='mlf_pie')

            st.markdown("#### Rekap per Cabang")
            gm = mlf.groupby('CABANG').agg(
                Voucher=('QTY', 'sum'), Transaksi=('QTY', 'size'),
                Omzet=('TOTAL HARGA', 'sum'), Modal=('MODAL', 'sum'), Laba=('LABA', 'sum'))
            gm['Margin %'] = (gm['Laba'] / gm['Omzet'] * 100).round(1)
            gm['Porsi %'] = (gm['Voucher'] / gm['Voucher'].sum() * 100).round(1)
            gm = gm.sort_values('Voucher', ascending=False)

            g1, g2 = st.columns([1, 1])
            with g1:
                st.dataframe(gm.style.format({'Voucher': '{:,.0f}', 'Transaksi': '{:,.0f}',
                                              'Omzet': 'Rp {:,.0f}', 'Modal': 'Rp {:,.0f}',
                                              'Laba': 'Rp {:,.0f}'}),
                             use_container_width=True, height=420, key='mlf_cab')
            with g2:
                gms = gm.sort_values('Voucher', ascending=True)
                figb = go.Figure(go.Bar(
                    y=gms.index, x=gms['Voucher'], orientation='h',
                    marker_color='#1f3864',
                    text=[f"{int(v):,}" for v in gms['Voucher']], textposition='outside'))
                figb.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=10),
                                   xaxis_title='Voucher terjual')
                st.plotly_chart(figb, use_container_width=True, key='mlf_cab_fig')

            p1, p2 = st.columns(2)
            with p1:
                st.markdown("#### 15 Penjual Teratas")
                gp = (mlf.groupby('PENJUAL').agg(Voucher=('QTY', 'sum'),
                                                 Omzet=('TOTAL HARGA', 'sum'))
                      .sort_values('Voucher', ascending=False).head(15))
                gp['Porsi %'] = (gp['Voucher'] / mlf['QTY'].sum() * 100).round(1)
                st.dataframe(gp.style.format({'Voucher': '{:,.0f}', 'Omzet': 'Rp {:,.0f}'}),
                             use_container_width=True, height=380, key='mlf_penjual')
            with p2:
                st.markdown("#### Rekap per Bulan")
                gb = mlf.groupby('BULAN').agg(Voucher=('QTY', 'sum'),
                                              Omzet=('TOTAL HARGA', 'sum'),
                                              Laba=('LABA', 'sum')).sort_index()
                gb.index = [BULAN_NAMES[int(b)] for b in gb.index]
                st.dataframe(gb.style.format({'Voucher': '{:,.0f}', 'Omzet': 'Rp {:,.0f}',
                                              'Laba': 'Rp {:,.0f}'}),
                             use_container_width=True, height=380, key='mlf_bulan')

            st.markdown("#### Detail Transaksi Voucher")
            qm = st.text_input("Cari cabang / penjual / faktur", key='mlf_cari')
            colm = ['TGL FAKTUR', 'NO FAKTUR', 'CABANG', 'PENJUAL', 'NAMA BARANG',
                    'HARGA BELI', 'QTY', '@HARGA', 'TOTAL HARGA', 'LABA']
            colm = [c for c in colm if c in mlf.columns]
            dm = mlf[colm].copy()
            if qm:
                mask = dm.apply(lambda r: qm.upper() in ' '.join(str(v) for v in r.values).upper(), axis=1)
                dm = dm[mask]
            st.caption(f"{len(dm):,} baris ditampilkan (maksimal 1.000).")
            st.dataframe(dm.head(1000), use_container_width=True, height=340, key='mlf_detail')

            # ---------- perbandingan periode ----------
            st.markdown("### 📊 Perbandingan Periode")
            _potm = potong_periode(mlf, 'TGL')
            _metrik_m = [
                ("Voucher terjual", lambda x: x['QTY'].sum(), lambda v: nfid(v), True),
                ("Transaksi", lambda x: len(x), lambda v: nfid(v), True),
                ("Omzet", lambda x: x['TOTAL HARGA'].sum(), lambda v: rp(v), True),
                ("Laba kotor", lambda x: x['LABA'].sum(), lambda v: rp(v), True),
                ("Cabang menjual", lambda x: x['CABANG'].nunique(),
                 lambda v: nfid(v), True),
            ]
            render_banding(_potm, _metrik_m, key_prefix='mlf')

            st.markdown("### 🧭 Analisa & Tindak Lanjut")
            _anm = []
            if _potm and not _potm['bulan_ini'].empty and not _potm['bulan_lalu'].empty:
                q_i = _potm['bulan_ini']['QTY'].sum()
                q_l = _potm['bulan_lalu']['QTY'].sum()
                if q_l:
                    g = (q_i - q_l) / q_l * 100
                    _anm.append(('baik' if g > 0 else 'perhatian',
                                 f"Penjualan voucher {'naik' if g > 0 else 'turun'} "
                                 f"{abs(g):.1f}% dibanding bulan lalu",
                                 f"Dari {q_l:,.0f} menjadi {q_i:,.0f} voucher. "
                                 + ("<b>Tindakan:</b> pertahankan dorongan penjualan dan "
                                    "tiru pola cabang terbaik." if g > 0 else
                                    "<b>Tindakan:</b> cek apakah stok voucher tersedia dan "
                                    "apakah penawaran masih rutin disampaikan ke pelanggan.")))

            _semua_cab = set(sales['CABANG'].unique())
            _jual_cab = set(mlf['CABANG'].unique())
            _belum = sorted(_semua_cab - _jual_cab)
            if _belum:
                _anm.append(('aksi', f'{len(_belum)} cabang belum menjual sama sekali',
                             f"{', '.join(_belum)}. <b>Tindakan:</b> pastikan stok voucher "
                             f"terdistribusi dan tim di sana sudah diberi pengarahan."))

            _pcab = mlf.groupby('CABANG')['QTY'].sum().sort_values(ascending=False)
            if len(_pcab) >= 3:
                _rata = _pcab.mean()
                _bawah = _pcab[_pcab < _rata * 0.6]
                if len(_bawah):
                    _anm.append(('perhatian', 'Cabang jauh di bawah rata-rata',
                                 f"{', '.join(f'{i} ({int(v)})' for i, v in _bawah.head(4).items())} "
                                 f"— rata-rata cabang {_rata:.0f} voucher. "
                                 f"<b>Tindakan:</b> samakan cara penawaran dengan "
                                 f"<b>{_pcab.index[0]}</b> yang mencapai {int(_pcab.iloc[0])}."))

            _mg = (laba / omzet * 100) if omzet else 0
            _anm.append(('info', f'Margin voucher {_mg:.1f}%',
                         f"Dari omzet {rp(omzet)} diperoleh laba {rp(laba)}. Margin voucher "
                         f"memang tipis karena harga belinya sudah tinggi — nilainya lebih "
                         f"pada menarik pelanggan datang, bukan pada labanya sendiri."))
            panel_analisa(_anm)

            tombol_pdf("Dashboard Voucher Tiket MLF", _potm, _metrik_m, _anm,
                       kpis=[{'label': 'Voucher Terjual', 'value': nfid(qty),
                              'sub': f"{nfid(len(mlf))} transaksi", 'warna': _PN},
                             {'label': 'Omzet', 'value': rp(omzet),
                              'sub': f"rata-rata {rp(omzet/qty if qty else 0)}/voucher", 'warna': _PN},
                             {'label': 'Laba Kotor', 'value': rp(laba),
                              'sub': f"margin {pctid(margin)}", 'warna': _PG},
                             {'label': 'Cabang Menjual', 'value': f"{n_cab}",
                              'sub': f"dari {sales['CABANG'].nunique()} cabang", 'warna': _PA}],
                       ringkasan=(f"{nfid(qty)} voucher terjual di {n_cab} cabang, "
                                  f"omzet {rp(omzet)} dengan laba {rp(laba)}."),
                       metodologi=("Dihitung dari baris yang nama barangnya mengandung "
                                   "kata MLF. Modal dari kolom HARGA BELI yang sudah "
                                   "berupa total per baris."),
                       key='mlf')

            with st.expander("ℹ️ Catatan metodologi"):
                nm = mlf['BARANG'].value_counts()
                st.write(
                    "Yang dihitung adalah semua baris yang **nama barangnya mengandung kata "
                    "'MLF'**. Item yang terdeteksi pada data saat ini:"
                )
                st.write("\n".join(f"- {k} ({v:,} baris)" for k, v in nm.items()))
                st.write(
                    "\n**Modal** dari kolom HARGA BELI yang sudah berupa total per baris "
                    "(bukan harga satuan). **Laba kotor** = TOTAL HARGA − MODAL, belum "
                    "dikurangi biaya operasional."
                )

# =============================================================================
# TAB 7: OMZET & BAGI HASIL TEKNISI
# =============================================================================
with tab_tek:
    st.markdown("## Omzet Jasa & Bagi Hasil per Teknisi")

    if sales.empty:
        st.info(
            "Data penjualan belum tersedia. Tambahkan file `data/penjualan.csv.gz` ke repo, "
            "atau upload lewat panel kiri."
        )
    else:
        jasa_all = sales[sales['KATEGORI'] == 'JASA'].copy()
        if jasa_all.empty:
            st.warning("Tidak ada baris berkategori JASA pada data penjualan.")
        else:
            # ---------- pengaturan tarif (bisa diubah manual) ----------
            with st.expander("⚙️ Pengaturan Tarif Bagi Hasil — klik untuk mengubah",
                             expanded=False):
                st.caption(
                    "Ubah angka di bawah sesuai kebijakan yang berlaku. Semua perhitungan, "
                    "tabel, dan grafik di tab ini langsung menyesuaikan."
                )
                ca, cb, cc, cd = st.columns(4)
                tarif_input = {}
                with ca:
                    tarif_input['Interface'] = st.number_input(
                        "Interface (%)", min_value=0.0, max_value=100.0,
                        value=TARIF_AWAL['Interface'], step=1.0, key='tar_int')
                with cb:
                    tarif_input['Normal'] = st.number_input(
                        "Normal (%)", min_value=0.0, max_value=100.0,
                        value=TARIF_AWAL['Normal'], step=1.0, key='tar_nor')
                with cc:
                    tarif_input['Mati Total'] = st.number_input(
                        "Mati Total (%)", min_value=0.0, max_value=100.0,
                        value=TARIF_AWAL['Mati Total'], step=1.0, key='tar_mat')
                with cd:
                    tarif_input['Promo'] = st.number_input(
                        "Promo (%)", min_value=0.0, max_value=100.0,
                        value=TARIF_AWAL['Promo'], step=1.0, key='tar_pro')

                ce, cf, cg = st.columns([1, 1, 1.6])
                with ce:
                    tarif_lain = st.number_input(
                        "Tanpa kata kunci (%)", min_value=0.0, max_value=100.0,
                        value=TARIF_DEFAULT_AWAL, step=1.0, key='tar_lain',
                        help="Dipakai untuk item seperti JASA REPAIR, JASA BATERAI, "
                             "JASA LCD 50% yang tidak mengandung kata kunci mana pun.")
                with cf:
                    tarif_flat = st.number_input(
                        "Tarif pembanding (%)", min_value=0.0, max_value=100.0,
                        value=TARIF_PEMBANDING_AWAL, step=1.0, key='tar_flat',
                        help="Skema pembanding: seluruh omzet jasa dikali tarif ini.")
                with cg:
                    prioritas = st.selectbox(
                        "Kalau satu nama mengandung 2 kata kunci, yang menang:",
                        ['Normal', 'Promo', 'Mati Total', 'Interface'],
                        index=0, key='tar_prio',
                        help="Contoh kasus: 'JS PROMO LCD 250K - NORMAL' mengandung "
                             "Promo dan Normal sekaligus.")

                if st.button("↩️ Kembalikan ke tarif awal", key='tar_reset'):
                    for k, v in [('tar_int', 'Interface'), ('tar_nor', 'Normal'),
                                 ('tar_mat', 'Mati Total'), ('tar_pro', 'Promo')]:
                        st.session_state[k] = TARIF_AWAL[v]
                    st.session_state['tar_lain'] = TARIF_DEFAULT_AWAL
                    st.session_state['tar_flat'] = TARIF_PEMBANDING_AWAL
                    st.session_state['tar_prio'] = 'Normal'
                    st.rerun()

            # urutan prioritas: pilihan pengguna didahulukan, sisanya menyusul
            urutan = [prioritas.upper()] + [k for k in KATA_KUNCI_TARIF
                                            if k != prioritas.upper()]
            peta_tarif = {k: v / 100.0 for k, v in tarif_input.items()}
            peta_tarif[LABEL_LAINNYA] = tarif_lain / 100.0

            # hitung ulang label & nilai bagi hasil berdasarkan input di atas
            jasa_all['TARIF_LABEL'] = jasa_all['KW_MATCH'].map(
                lambda s: pilih_label_tarif(s, urutan))
            jasa_all['TARIF'] = jasa_all['TARIF_LABEL'].map(peta_tarif).fillna(0.0)
            jasa_all['BAGI_HASIL'] = jasa_all['TOTAL HARGA'] * jasa_all['TARIF']
            jasa_all['BAGI_HASIL_FLAT'] = jasa_all['TOTAL HARGA'] * (tarif_flat / 100.0)

            ringkas_tarif = " · ".join(
                f"{k} {v:.0f}%" for k, v in list(tarif_input.items()) +
                [('Lainnya', tarif_lain)])
            st.caption(f"**Tarif aktif:** {ringkas_tarif}  ·  "
                       f"pembanding flat {tarif_flat:.0f}%  ·  prioritas bentrok: {prioritas}")

            # ---------- pemilih periode penggajian (cutoff 24 -> 23) ----------
            periode_list = daftar_periode_gaji(jasa_all['TGL'].min(), jasa_all['TGL'].max())
            opsi = ['Semua Periode'] + periode_list

            cpa, cpb = st.columns([2, 1])
            with cpa:
                pilih = st.selectbox(
                    "Periode penggajian (cutoff tanggal 24 s/d 23)",
                    opsi,
                    index=len(opsi) - 1 if len(opsi) > 1 else 0,
                    format_func=lambda x: ("Semua Periode (tanpa cutoff)"
                                           if isinstance(x, str)
                                           else label_periode(x[1], x[0])),
                    key='tek_periode',
                )
            with cpb:
                pakai_cabang = st.checkbox(
                    "Ikuti filter Cabang di panel kiri", value=True, key='tek_pakai_cab')

            jasa = jasa_all
            if pakai_cabang and f_cabang != 'Semua Cabang':
                jasa = jasa[jasa['CABANG'] == f_cabang]

            if isinstance(pilih, str):
                periode_txt = "Seluruh periode data (tanpa cutoff penggajian)"
                a = b = None
            else:
                a, b = periode_gaji(pilih[1], pilih[0])
                jasa = jasa[(jasa['TGL'] >= a) & (jasa['TGL'] <= b)]
                periode_txt = (f"{a.day} {BULAN_NAMES[a.month]} {a.year} – "
                               f"{b.day} {BULAN_NAMES[b.month]} {b.year}")

            st.caption(f"Rentang data dihitung: **{periode_txt}**"
                       + ("" if not (pakai_cabang and f_cabang != 'Semua Cabang')
                          else f" · cabang **{f_cabang}**"))

            if jasa.empty:
                st.warning("Tidak ada transaksi jasa pada periode/cabang tersebut.")
            else:
                omzet_j = jasa['TOTAL HARGA'].sum()
                bh = jasa['BAGI_HASIL'].sum()
                bh_flat = jasa['BAGI_HASIL_FLAT'].sum()
                selisih = bh - bh_flat
                n_tek = jasa.loc[jasa['TEKNISI'] != 'TIDAK ADA TEKNISI', 'TEKNISI'].nunique()

                # peringatan bila periode ini belum memakai penamaan berkata kunci
                n_kw = (jasa['TARIF_LABEL'] != 'Lainnya').sum()
                if n_kw == 0:
                    sama = abs(tarif_lain - tarif_flat) < 1e-9
                    st.warning(
                        "Pada periode ini **tidak ada satu pun item jasa yang mengandung kata "
                        "kunci** (Interface / Normal / Mati Total / Promo). Semua item memakai "
                        "penamaan lama seperti `JASA REPAIR`, sehingga seluruhnya kena tarif "
                        f"tanpa-kata-kunci **{tarif_lain:.0f}%**"
                        + (f" — dan karena tarif pembanding juga {tarif_flat:.0f}%, kedua skema "
                           "menghasilkan angka **sama persis**." if sama else ".")
                        + " Penamaan berkata kunci baru mulai dipakai sekitar Juli 2026."
                    )
                elif n_kw < len(jasa) * 0.5:
                    st.info(
                        f"Baru **{n_kw:,} dari {len(jasa):,} baris** ({n_kw/len(jasa)*100:.0f}%) "
                        "yang memakai penamaan berkata kunci; sisanya kena tarif default 30%. "
                        "Selisih terhadap skema flat masih kecil selama penamaan belum seragam."
                    )

                st.markdown(kpi_html([
                    {'label': 'Omzet Jasa', 'value': rp(omzet_j),
                     'sub': f"{len(jasa):,} baris jasa",
                     'grad': 'linear-gradient(135deg,#1f3864,#2e5394)'},
                    {'label': 'Bagi Hasil (Aturan)', 'value': rp(bh),
                     'sub': f"{(bh/omzet_j*100 if omzet_j else 0):.1f}% dari omzet jasa",
                     'grad': 'linear-gradient(135deg,#16a34a,#22c55e)'},
                    {'label': f'Pembanding (Flat {tarif_flat:.0f}%)', 'value': rp(bh_flat),
                     'sub': f'seluruh omzet jasa × {tarif_flat:.0f}%',
                     'grad': 'linear-gradient(135deg,#7c3aed,#a855f7)'},
                    {'label': 'Selisih', 'value': rp(selisih),
                     'sub': ('aturan lebih besar' if selisih > 0 else
                             ('flat lebih besar' if selisih < 0 else 'sama')),
                     'grad': ('linear-gradient(135deg,#e0921f,#e2b21a)' if selisih >= 0
                              else 'linear-gradient(135deg,#c9392f,#e0475a)')},
                    {'label': 'Jumlah Teknisi', 'value': nfid(n_tek),
                     'sub': f"rata-rata {rp(bh/n_tek if n_tek else 0)}/teknisi",
                     'grad': 'linear-gradient(135deg,#0f8a82,#17a3a3)'},
                    {'label': 'Tanpa Nama Teknisi', 'value': rp(
                        jasa.loc[jasa['TEKNISI'] == 'TIDAK ADA TEKNISI', 'TOTAL HARGA'].sum()),
                     'sub': f"{(jasa['TEKNISI'] == 'TIDAK ADA TEKNISI').sum():,} baris",
                     'grad': 'linear-gradient(135deg,#64748b,#94a3b8)'},
                ]), unsafe_allow_html=True)
                st.write("")

                # ---------- rekap per teknisi ----------
                gt = jasa.groupby('TEKNISI').agg(
                    Baris=('TOTAL HARGA', 'size'),
                    Omzet_Jasa=('TOTAL HARGA', 'sum'),
                    Bagi_Hasil=('BAGI_HASIL', 'sum'),
                    Flat_30=('BAGI_HASIL_FLAT', 'sum'))
                gt['Selisih'] = gt['Bagi_Hasil'] - gt['Flat_30']
                gt['Efektif %'] = (gt['Bagi_Hasil'] / gt['Omzet_Jasa'] * 100).round(1)
                gt = gt.sort_values('Bagi_Hasil', ascending=False)
                lbl_flat = f'Pembanding {tarif_flat:.0f}%'
                gt_show = gt.rename(columns={'Omzet_Jasa': 'Omzet Jasa',
                                             'Bagi_Hasil': 'Bagi Hasil (Aturan)',
                                             'Flat_30': lbl_flat})

                t1, t2 = st.columns([1.05, 1])
                with t1:
                    st.markdown("#### 15 Teknisi — Bagi Hasil Tertinggi")
                    top = gt[gt.index != 'TIDAK ADA TEKNISI'].head(15).sort_values('Bagi_Hasil')
                    figt = go.Figure()
                    figt.add_bar(y=top.index, x=top['Bagi_Hasil'], orientation='h',
                                 name='Aturan', marker_color='#16a34a')
                    figt.add_bar(y=top.index, x=top['Flat_30'], orientation='h',
                                 name=f'Flat {tarif_flat:.0f}%', marker_color='#a855f7')
                    figt.update_layout(barmode='group', height=520,
                                       margin=dict(l=10, r=10, t=10, b=10),
                                       legend=dict(orientation='h', y=1.04),
                                       xaxis_title='Rupiah')
                    st.plotly_chart(figt, use_container_width=True, key='tek_fig')
                with t2:
                    st.markdown("#### Komposisi Omzet Jasa per Tarif")
                    gtar = jasa.groupby('TARIF_LABEL').agg(
                        Baris=('TOTAL HARGA', 'size'),
                        Omzet=('TOTAL HARGA', 'sum'),
                        Bagi_Hasil=('BAGI_HASIL', 'sum'))
                    gtar['Tarif'] = gtar.index.map(
                        lambda k: peta_tarif.get(k, 0.0) * 100)
                    gtar['Tarif'] = gtar['Tarif'].round(1).astype(str) + '%'
                    gtar = gtar.sort_values('Omzet', ascending=False)
                    st.dataframe(
                        gtar[['Tarif', 'Baris', 'Omzet', 'Bagi_Hasil']].rename(
                            columns={'Bagi_Hasil': 'Bagi Hasil'}).style.format(
                            {'Baris': '{:,.0f}', 'Omzet': 'Rp {:,.0f}',
                             'Bagi Hasil': 'Rp {:,.0f}'}),
                        use_container_width=True, key='tek_tarif')
                    figtar = px.pie(names=gtar.index, values=gtar['Omzet'], hole=0.55,
                                    color_discrete_sequence=PALETTE)
                    figtar.update_layout(height=280, margin=dict(l=5, r=5, t=5, b=5),
                                         legend=dict(font=dict(size=9)))
                    st.plotly_chart(figtar, use_container_width=True, key='tek_tarif_fig')

                st.markdown("#### Rekap Lengkap per Teknisi")
                st.dataframe(
                    gt_show.style.format({
                        'Baris': '{:,.0f}', 'Omzet Jasa': 'Rp {:,.0f}',
                        'Bagi Hasil (Aturan)': 'Rp {:,.0f}',
                        lbl_flat: 'Rp {:,.0f}', 'Selisih': 'Rp {:,.0f}'}),
                    use_container_width=True, height=460, key='tek_tabel')

                csv_tek = gt_show.reset_index().to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    "⬇️ Unduh rekap teknisi (CSV)", data=csv_tek,
                    file_name=(f"bagi_hasil_teknisi_"
                               f"{'semua' if isinstance(pilih, str) else f'{pilih[0]}-{pilih[1]:02d}'}.csv"),
                    mime="text/csv", key='tek_unduh')

                st.markdown("#### Rekap per Cabang")
                gcb = jasa.groupby('CABANG').agg(
                    Teknisi=('TEKNISI', 'nunique'),
                    Omzet_Jasa=('TOTAL HARGA', 'sum'),
                    Bagi_Hasil=('BAGI_HASIL', 'sum'),
                    Flat_30=('BAGI_HASIL_FLAT', 'sum'))
                gcb['Selisih'] = gcb['Bagi_Hasil'] - gcb['Flat_30']
                gcb = gcb.sort_values('Omzet_Jasa', ascending=False)
                st.dataframe(
                    gcb.rename(columns={'Omzet_Jasa': 'Omzet Jasa',
                                        'Bagi_Hasil': 'Bagi Hasil (Aturan)',
                                        'Flat_30': lbl_flat}).style.format({
                        'Teknisi': '{:,.0f}', 'Omzet Jasa': 'Rp {:,.0f}',
                        'Bagi Hasil (Aturan)': 'Rp {:,.0f}',
                        lbl_flat: 'Rp {:,.0f}', 'Selisih': 'Rp {:,.0f}'}),
                    use_container_width=True, height=380, key='tek_cabang')

                st.markdown("#### Detail Transaksi Jasa")
                qt = st.text_input("Cari teknisi / barang / faktur", key='tek_cari')
                colt = ['TGL FAKTUR', 'NO FAKTUR', 'CABANG', 'TEKNISI', 'NAMA BARANG',
                        'TARIF_LABEL', 'TARIF', 'TOTAL HARGA', 'BAGI_HASIL', 'BAGI_HASIL_FLAT']
                colt = [c for c in colt if c in jasa.columns]
                dt = jasa[colt].copy()
                dt = dt.rename(columns={'TARIF_LABEL': 'Kategori Tarif', 'TARIF': 'Tarif',
                                        'BAGI_HASIL': 'Bagi Hasil',
                                        'BAGI_HASIL_FLAT': f'Flat {tarif_flat:.0f}%'})
                if qt:
                    mask = dt.apply(lambda r: qt.upper() in ' '.join(str(v) for v in r.values).upper(), axis=1)
                    dt = dt[mask]
                st.caption(f"{len(dt):,} baris ditampilkan (maksimal 1.000).")
                st.dataframe(dt.head(1000), use_container_width=True, height=360, key='tek_detail')

                # ---------- perbandingan periode ----------
                st.markdown("### 📊 Perbandingan Periode")
                _pott = potong_periode(jasa_all, 'TGL')
                _metrik_t = [
                    ("Omzet jasa", lambda x: x['TOTAL HARGA'].sum(), lambda v: rp(v), True),
                    ("Bagi hasil (aturan)", lambda x: x['BAGI_HASIL'].sum(),
                     lambda v: rp(v), True),
                    ("Pembanding flat", lambda x: x['BAGI_HASIL_FLAT'].sum(),
                     lambda v: rp(v), True),
                    ("Tarif efektif", lambda x: (x['BAGI_HASIL'].sum() /
                                                 x['TOTAL HARGA'].sum() * 100)
                     if x['TOTAL HARGA'].sum() else 0,
                     lambda v: pctid(v), False),
                    ("Teknisi aktif",
                     lambda x: x.loc[x['TEKNISI'] != 'TIDAK ADA TEKNISI', 'TEKNISI'].nunique(),
                     lambda v: nfid(v), True),
                ]
                render_banding(_pott, _metrik_t, key_prefix='tek')

                st.markdown("### 🧭 Analisa & Tindak Lanjut")
                _ant = []
                if selisih < 0:
                    _ant.append(('baik', 'Skema aturan lebih hemat dari flat',
                                 f"Selisih {rp(abs(selisih))} lebih rendah dibanding skema "
                                 f"flat {tarif_flat:.0f}% ({abs(selisih)/bh_flat*100:.1f}%). "
                                 f"Penyebabnya porsi <b>Interface</b> yang bertarif rendah "
                                 f"cukup besar. <b>Tindakan:</b> pastikan pembagian jenis "
                                 f"pekerjaan sudah adil bagi teknisi agar tidak memicu keluhan."))
                elif selisih > 0:
                    _ant.append(('perhatian', 'Skema aturan lebih mahal dari flat',
                                 f"Selisih {rp(selisih)} lebih tinggi dibanding skema flat "
                                 f"{tarif_flat:.0f}%. <b>Tindakan:</b> tinjau apakah tarif "
                                 f"Promo/Mati Total masih sepadan dengan nilai pekerjaannya."))

                if n_kw == 0:
                    _ant.append(('aksi', 'Penamaan barang belum memakai kata kunci',
                                 "Tidak ada item berkata kunci pada periode ini, sehingga "
                                 "seluruh skema tarif tidak berpengaruh. <b>Tindakan:</b> "
                                 "seragamkan penamaan item jasa di sistem (format "
                                 "<code>JS ... - INTERFACE/NORMAL</code>) agar aturan bagi "
                                 "hasil benar-benar berjalan."))
                elif n_kw < len(jasa) * 0.8:
                    _ant.append(('aksi', 'Penamaan belum seragam',
                                 f"Baru {n_kw/len(jasa)*100:.0f}% baris memakai penamaan "
                                 f"berkata kunci; sisanya jatuh ke tarif default "
                                 f"{tarif_lain:.0f}%. <b>Tindakan:</b> lanjutkan migrasi "
                                 f"penamaan agar perhitungan makin tepat."))

                _tk = gt[gt.index != 'TIDAK ADA TEKNISI']
                if len(_tk) >= 5:
                    _q1 = _tk['Bagi_Hasil'].quantile(0.25)
                    _bwh = _tk[_tk['Bagi_Hasil'] <= _q1]
                    _ant.append(('info', 'Sebaran bagi hasil antar teknisi',
                                 f"Tertinggi {rp(_tk['Bagi_Hasil'].max())} "
                                 f"(<b>{_tk['Bagi_Hasil'].idxmax()}</b>), terendah "
                                 f"{rp(_tk['Bagi_Hasil'].min())}. Seperempat teknisi "
                                 f"({len(_bwh)} orang) di bawah {rp(_q1)}. "
                                 f"<b>Tindakan:</b> cek apakah karena beban kerja yang "
                                 f"timpang atau memang jam kerja berbeda."))

                _tanpa = jasa.loc[jasa['TEKNISI'] == 'TIDAK ADA TEKNISI', 'TOTAL HARGA'].sum()
                if _tanpa > omzet_j * 0.03:
                    _ant.append(('aksi', 'Omzet jasa tanpa nama teknisi',
                                 f"{rp(_tanpa)} ({_tanpa/omzet_j*100:.1f}% omzet jasa) tidak "
                                 f"tercatat teknisinya, sehingga bagi hasilnya tidak bisa "
                                 f"dialokasikan. <b>Tindakan:</b> wajibkan pengisian kolom "
                                 f"teknisi saat pembuatan faktur."))
                if not _ant:
                    _ant.append(('info', 'Belum ada temuan menonjol', "Angka relatif stabil."))
                panel_analisa(_ant)

                tombol_pdf("Dashboard Bagi Hasil Teknisi", _pott, _metrik_t, _ant,
                           kpis=[{'label': 'Omzet Jasa', 'value': rp(omzet_j),
                                  'sub': f"{nfid(len(jasa))} baris", 'warna': _PN},
                                 {'label': 'Bagi Hasil (Aturan)', 'value': rp(bh),
                                  'sub': f"{pctid(bh/omzet_j*100 if omzet_j else 0)} omzet jasa",
                                  'warna': _PG},
                                 {'label': f'Pembanding Flat {tarif_flat:.0f}%',
                                  'value': rp(bh_flat), 'sub': 'skema pembanding', 'warna': _PN},
                                 {'label': 'Selisih', 'value': rp(selisih),
                                  'sub': ('aturan lebih besar' if selisih > 0 else 'flat lebih besar'),
                                  'warna': (_PA if selisih >= 0 else _PR)},
                                 {'label': 'Jumlah Teknisi', 'value': nfid(n_tek),
                                  'sub': f"rata-rata {rp(bh/n_tek if n_tek else 0)}", 'warna': _PN}],
                           ringkasan=(f"Periode {periode_txt}. Omzet jasa {rp(omzet_j)} "
                                      f"menghasilkan bagi hasil {rp(bh)}, "
                                      f"{'lebih hemat' if selisih < 0 else 'lebih besar'} "
                                      f"{rp(abs(selisih))} dibanding skema flat "
                                      f"{tarif_flat:.0f}%."),
                           metodologi=(f"Tarif: Interface {tarif_input['Interface']:.0f}%, "
                                       f"Normal {tarif_input['Normal']:.0f}%, Mati Total "
                                       f"{tarif_input['Mati Total']:.0f}%, Promo "
                                       f"{tarif_input['Promo']:.0f}%, tanpa kata kunci "
                                       f"{tarif_lain:.0f}%. Bila dua kata kunci bertemu, "
                                       f"dipakai {prioritas}. Periode penggajian memakai "
                                       f"cutoff tanggal 24 s/d 23. Angka berbasis omzet "
                                       f"jasa, belum dikurangi biaya."),
                           key='tek')

                with st.expander("ℹ️ Cara perhitungan & catatan"):
                    st.write(
                        "**Tarif bagi hasil** ditentukan dari kata kunci pada kolom NAMA BARANG. "
                        "Angka di bawah ini mengikuti isian pada panel **Pengaturan Tarif** di atas:\n"
                        f"- mengandung **Interface** → {tarif_input['Interface']:.0f}%\n"
                        f"- mengandung **Normal** → {tarif_input['Normal']:.0f}%\n"
                        f"- mengandung **Mati Total** → {tarif_input['Mati Total']:.0f}%\n"
                        f"- mengandung **Promo** → {tarif_input['Promo']:.0f}%\n"
                        f"- tidak mengandung kata kunci mana pun → **{tarif_lain:.0f}%** "
                        "(mencakup item berpola `JASA ...` seperti JASA REPAIR, JASA BATERAI, "
                        "JASA LCD 50%, yang porsinya justru mayoritas)\n\n"
                        "Bila satu nama mengandung **dua kata kunci** sekaligus (mis. "
                        f"`JS PROMO LCD 250K - NORMAL`), yang dipakai adalah **{prioritas} "
                        f"{tarif_input[prioritas]:.0f}%** sesuai pilihan prioritas.\n\n"
                        "**Periode penggajian** memakai cutoff tanggal 24 sampai 23: gaji bulan M "
                        "dihitung dari 24 bulan (M−2) sampai 23 bulan (M−1). Contoh gaji Mei 2026 "
                        "= 24 Maret 2026 s/d 23 April 2026. Tanggal acuannya adalah **TGL FAKTUR**.\n\n"
                        f"**Pembanding Flat {tarif_flat:.0f}%** = seluruh omzet jasa × {tarif_flat:.0f}%, tanpa membedakan jenis "
                        "pekerjaan. Kolom *Selisih* menunjukkan berapa lebih besar/kecil skema "
                        "aturan dibanding skema flat.\n\n"
                        "Nama teknisi diambil dari kolom **NAMA TEKNISI (FINAL)**; bila kosong, "
                        "dipakai kolom NAMA TEKNISI. Baris yang keduanya kosong dikelompokkan "
                        "sebagai *TIDAK ADA TEKNISI* dan tetap ditampilkan agar terlihat, tetapi "
                        "sebaiknya dirapikan di sumber data.\n\n"
                        "Perhitungan ini memakai **omzet jasa (TOTAL HARGA)**, belum dikurangi "
                        "biaya apa pun."
                    )

# =============================================================================
# TAB 8: BUNDLING AKSESORIS
# =============================================================================
with tab_bundling:
    st.markdown("## Bundling Aksesoris — Attach Rate per Nota")

    if sales.empty:
        st.info(
            "Data penjualan belum tersedia. Tambahkan file `data/penjualan.csv.gz` ke repo, "
            "atau upload lewat panel kiri."
        )
    elif sales_f.empty:
        st.warning("Tidak ada data penjualan untuk filter yang dipilih.")
    else:
        sb = sales_f.copy()
        sb['IS_AKS'] = sb['KATEGORI'].isin(['AKSESORIS', 'ACCESORIES'])

        # Satu nota = kombinasi CABANG + NO FAKTUR, karena penomoran faktur
        # berjalan sendiri-sendiri di tiap cabang.
        nota = sb.groupby(['CABANG', 'NO FAKTUR']).agg(
            Baris=('KATEGORI', 'size'),
            Qty=('QTY', 'sum'),
            Nilai=('TOTAL HARGA', 'sum'),
            Laba=('LABA', 'sum'),
            n_aks=('IS_AKS', 'sum'),
            Tgl=('TGL', 'min'),
            Penjual=('PENJUAL', 'first'),
        ).reset_index()
        nota['n_lain'] = nota['Baris'] - nota['n_aks']
        nota['Ada Aksesoris'] = nota['n_aks'] > 0
        nota['Ada Lainnya'] = nota['n_lain'] > 0
        nota['Bundling'] = (nota['Ada Aksesoris'] & nota['Ada Lainnya']
                            & (nota['Baris'] >= 2))
        nota['BULAN'] = nota['Tgl'].dt.month

        # nilai aksesoris per nota (untuk mengukur kontribusi aksesorisnya)
        aks_nilai = (sb[sb['IS_AKS']].groupby(['CABANG', 'NO FAKTUR'])['TOTAL HARGA']
                     .sum().rename('Nilai Aksesoris'))
        nota = nota.merge(aks_nilai, on=['CABANG', 'NO FAKTUR'], how='left')
        nota['Nilai Aksesoris'] = nota['Nilai Aksesoris'].fillna(0)

        tot_nota = len(nota)
        bund = nota[nota['Bundling']]
        non = nota[~nota['Bundling']]
        n_bund = len(bund)
        rate = (n_bund / tot_nota * 100) if tot_nota else 0

        # peluang: nota non-aksesoris yang belum dilekati aksesoris
        peluang = nota[(~nota['Ada Aksesoris']) & nota['Ada Lainnya']]

        st.markdown(kpi_html([
            {'label': 'Total Nota', 'value': f"{tot_nota:,}",
             'sub': 'sesuai filter aktif',
             'grad': 'linear-gradient(135deg,#1f3864,#2e5394)'},
            {'label': 'Nota Bundling', 'value': f"{n_bund:,}",
             'sub': f"attach rate {rate:.1f}%",
             'grad': 'linear-gradient(135deg,#16a34a,#22c55e)'},
            {'label': 'Belum Bundling', 'value': f"{len(peluang):,}",
             'sub': 'ada barang/jasa, tanpa aksesoris',
             'grad': 'linear-gradient(135deg,#e0921f,#e2b21a)'},
            {'label': 'Omzet Bundling', 'value': rp(bund['Nilai'].sum()),
             'sub': f"{(bund['Nilai'].sum()/nota['Nilai'].sum()*100 if nota['Nilai'].sum() else 0):.1f}% dari omzet",
             'grad': 'linear-gradient(135deg,#7c3aed,#a855f7)'},
            {'label': 'Rata-rata Nota Bundling', 'value': rp(bund['Nilai'].mean() if n_bund else 0),
             'sub': f"non-bundling {rp(non['Nilai'].mean() if len(non) else 0)}",
             'grad': 'linear-gradient(135deg,#0f8a82,#17a3a3)'},
            {'label': 'Aksesoris dalam Bundling', 'value': rp(bund['Nilai Aksesoris'].sum()),
             'sub': f"{(bund['Nilai Aksesoris'].sum()/bund['Nilai'].sum()*100 if n_bund and bund['Nilai'].sum() else 0):.1f}% dari nilai nota",
             'grad': 'linear-gradient(135deg,#c93fa8,#d1478d)'},
        ]), unsafe_allow_html=True)
        st.write("")

        st.caption(
            "**Definisi:** satu nota disebut *bundling* bila memuat minimal satu item "
            "berkategori **AKSESORIS** sekaligus minimal satu item kategori lain "
            "(jasa, sparepart, handphone, dll), dengan total minimal 2 produk. "
            "Satu nota dihitung sebagai kombinasi Cabang + Nomor Faktur."
        )

        b1, b2 = st.columns([1.25, 1])
        with b1:
            st.markdown("#### Tren Attach Rate per Bulan")
            tren = nota.groupby('BULAN').agg(
                total=('Bundling', 'size'), bundling=('Bundling', 'sum')).sort_index()
            tren['rate'] = (tren['bundling'] / tren['total'] * 100).round(1)
            figb = go.Figure()
            figb.add_bar(x=[BULAN_NAMES[int(b)][:3] for b in tren.index],
                         y=tren['bundling'], name='Nota bundling',
                         marker_color='#16a34a')
            figb.add_bar(x=[BULAN_NAMES[int(b)][:3] for b in tren.index],
                         y=tren['total'] - tren['bundling'], name='Bukan bundling',
                         marker_color='#d9dee9')
            figb.add_trace(go.Scatter(
                x=[BULAN_NAMES[int(b)][:3] for b in tren.index], y=tren['rate'],
                name='Attach rate (%)', yaxis='y2', mode='lines+markers',
                line=dict(color='#1f3864', width=3)))
            figb.update_layout(
                barmode='stack', height=360, margin=dict(l=10, r=10, t=10, b=10),
                legend=dict(orientation='h', y=1.12),
                yaxis=dict(title='Jumlah nota'),
                yaxis2=dict(title='Attach rate (%)', overlaying='y', side='right',
                            range=[0, 100]))
            st.plotly_chart(figb, use_container_width=True, key='bund_tren')
        with b2:
            st.markdown("#### Komposisi Nota")
            komp = pd.DataFrame({
                'Jenis': ['Bundling (aksesoris + lainnya)',
                          'Hanya aksesoris',
                          'Hanya non-aksesoris'],
                'Jumlah': [n_bund,
                           int((nota['Ada Aksesoris'] & ~nota['Ada Lainnya']).sum()),
                           int((~nota['Ada Aksesoris'] & nota['Ada Lainnya']).sum())]
            })
            figk = px.pie(komp, names='Jenis', values='Jumlah', hole=0.55,
                          color_discrete_sequence=['#16a34a', '#c93fa8', '#e0921f'])
            figk.update_layout(height=360, margin=dict(l=5, r=5, t=5, b=5),
                               legend=dict(font=dict(size=9), orientation='h', y=-0.1))
            st.plotly_chart(figk, use_container_width=True, key='bund_pie')

        st.markdown("#### Attach Rate per Cabang")
        gc = nota.groupby('CABANG').agg(
            Nota=('Bundling', 'size'), Bundling=('Bundling', 'sum'),
            Omzet=('Nilai', 'sum'), Nilai_Aksesoris=('Nilai Aksesoris', 'sum'))
        gc['Belum Bundling'] = gc['Nota'] - gc['Bundling']
        gc['Attach Rate %'] = (gc['Bundling'] / gc['Nota'] * 100).round(1)
        gc = gc.sort_values('Attach Rate %', ascending=False)

        cc1, cc2 = st.columns([1, 1.1])
        with cc1:
            st.dataframe(
                gc[['Nota', 'Bundling', 'Belum Bundling', 'Attach Rate %',
                    'Nilai_Aksesoris']].rename(
                    columns={'Nilai_Aksesoris': 'Nilai Aksesoris'}).style.format(
                    {'Nota': '{:,.0f}', 'Bundling': '{:,.0f}',
                     'Belum Bundling': '{:,.0f}', 'Nilai Aksesoris': 'Rp {:,.0f}'}),
                use_container_width=True, height=420, key='bund_cab')
        with cc2:
            gcs = gc.sort_values('Attach Rate %', ascending=True)
            rata = rate
            figc = go.Figure(go.Bar(
                y=gcs.index, x=gcs['Attach Rate %'], orientation='h',
                marker_color=['#16a34a' if v >= rata else '#e0921f'
                              for v in gcs['Attach Rate %']],
                text=[f"{v:.1f}%" for v in gcs['Attach Rate %']],
                textposition='outside'))
            figc.add_vline(x=rata, line_dash='dash', line_color='#1f3864',
                           annotation_text=f"rata-rata {rata:.1f}%")
            figc.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=10),
                               xaxis_title='Attach rate (%)',
                               xaxis=dict(range=[0, max(100, gcs['Attach Rate %'].max() * 1.15)]))
            st.plotly_chart(figc, use_container_width=True, key='bund_cab_fig')

        k1, k2 = st.columns(2)
        with k1:
            st.markdown("#### Kategori Pendamping Aksesoris")
            pend = (sb[~sb['IS_AKS']]
                    .merge(bund[['CABANG', 'NO FAKTUR']], on=['CABANG', 'NO FAKTUR'])
                    .groupby(['CABANG', 'NO FAKTUR'])['KATEGORI']
                    .apply(lambda s: list(dict.fromkeys(s))))
            from collections import Counter
            cnt = Counter(k for lst in pend for k in lst)
            dfp = (pd.DataFrame(cnt.items(), columns=['Kategori', 'Nota'])
                   .sort_values('Nota', ascending=False).head(10))
            dfp['% dari nota bundling'] = (dfp['Nota'] / n_bund * 100).round(1) if n_bund else 0
            st.dataframe(dfp.style.format({'Nota': '{:,.0f}'}),
                         use_container_width=True, height=340, hide_index=True,
                         key='bund_pendamping')
        with k2:
            st.markdown("#### Aksesoris Paling Sering Di-bundling")
            aks_top = (sb[sb['IS_AKS']]
                       .merge(bund[['CABANG', 'NO FAKTUR']], on=['CABANG', 'NO FAKTUR'])
                       .groupby('BARANG').agg(Nota=('QTY', 'size'), Qty=('QTY', 'sum'),
                                              Nilai=('TOTAL HARGA', 'sum'))
                       .sort_values('Nilai', ascending=False).head(12))
            st.dataframe(aks_top.style.format({'Nota': '{:,.0f}', 'Qty': '{:,.0f}',
                                               'Nilai': 'Rp {:,.0f}'}),
                         use_container_width=True, height=340, key='bund_aks')

        st.markdown("#### Attach Rate per Penjual (minimal 30 nota)")
        gp = nota.groupby('Penjual').agg(
            Nota=('Bundling', 'size'), Bundling=('Bundling', 'sum'),
            Omzet=('Nilai', 'sum'))
        gp = gp[gp['Nota'] >= 30]
        gp['Attach Rate %'] = (gp['Bundling'] / gp['Nota'] * 100).round(1)
        gp = gp.sort_values('Attach Rate %', ascending=False)
        if len(gp):
            st.dataframe(gp.style.format({'Nota': '{:,.0f}', 'Bundling': '{:,.0f}',
                                          'Omzet': 'Rp {:,.0f}'}),
                         use_container_width=True, height=340, key='bund_penjual')
        else:
            st.caption("Belum ada penjual dengan minimal 30 nota pada filter ini.")

        st.markdown("#### Daftar Nota")
        f1, f2 = st.columns([1, 2])
        with f1:
            saring = st.selectbox("Tampilkan", ['Semua nota', 'Hanya bundling',
                                                'Belum bundling (peluang)'],
                                  key='bund_saring')
        with f2:
            cari_b = st.text_input("Cari cabang / faktur / penjual", key='bund_cari')
        dn = nota.copy()
        if saring == 'Hanya bundling':
            dn = dn[dn['Bundling']]
        elif saring == 'Belum bundling (peluang)':
            dn = dn[(~dn['Ada Aksesoris']) & dn['Ada Lainnya']]
        kol = ['Tgl', 'CABANG', 'NO FAKTUR', 'Penjual', 'Baris', 'Qty',
               'Nilai', 'Nilai Aksesoris', 'Bundling']
        dn = dn[kol].rename(columns={'Tgl': 'Tanggal', 'CABANG': 'Cabang',
                                     'NO FAKTUR': 'No Faktur', 'Baris': 'Jml Produk'})
        if cari_b:
            m = dn.apply(lambda r: cari_b.upper() in
                         ' '.join(str(v) for v in r.values).upper(), axis=1)
            dn = dn[m]
        st.caption(f"{len(dn):,} nota (ditampilkan maksimal 1.000).")
        st.dataframe(dn.sort_values('Nilai', ascending=False).head(1000),
                     use_container_width=True, height=360, hide_index=True,
                     key='bund_detail')

        st.download_button(
            "⬇️ Unduh rekap bundling per cabang (CSV)",
            data=gc.reset_index().to_csv(index=False).encode('utf-8-sig'),
            file_name="bundling_aksesoris_per_cabang.csv", mime="text/csv",
            key='bund_unduh')

        # ---------- perbandingan periode ----------
        st.markdown("### 📊 Perbandingan Periode")
        _nota_p = nota.rename(columns={'Tgl': 'TGL'})
        _potb = potong_periode(_nota_p, 'TGL')
        _metrik_b = [
            ("Total nota", lambda x: len(x), lambda v: nfid(v), True),
            ("Nota bundling", lambda x: int(x['Bundling'].sum()),
             lambda v: nfid(v), True),
            ("Attach rate", lambda x: x['Bundling'].mean() * 100 if len(x) else 0,
             lambda v: pctid(v), True),
            ("Nilai aksesoris", lambda x: x['Nilai Aksesoris'].sum(),
             lambda v: rp(v), True),
            ("Belum bundling",
             lambda x: int(((~x['Ada Aksesoris']) & x['Ada Lainnya']).sum()),
             lambda v: nfid(v), False),
        ]
        render_banding(_potb, _metrik_b, key_prefix='bund')

        st.markdown("### 🧭 Analisa & Tindak Lanjut")
        _anb = []
        if _potb and not _potb['bulan_ini'].empty and not _potb['bulan_lalu'].empty:
            r_i = _potb['bulan_ini']['Bundling'].mean() * 100
            r_l = _potb['bulan_lalu']['Bundling'].mean() * 100
            d = r_i - r_l
            if d < -1:
                _anb.append(('aksi', f'Attach rate turun {abs(d):.1f} poin',
                             f"Dari {r_l:.1f}% ke {r_i:.1f}%. <b>Tindakan:</b> ingatkan "
                             f"kembali kebiasaan menawarkan aksesoris saat serah terima "
                             f"unit; cek juga apakah stok aksesoris sedang kosong."))
            elif d > 1:
                _anb.append(('baik', f'Attach rate naik {d:.1f} poin',
                             f"Dari {r_l:.1f}% ke {r_i:.1f}%. <b>Tindakan:</b> catat apa "
                             f"yang berubah bulan ini dan bakukan ke seluruh cabang."))

        if len(peluang):
            _pc2 = peluang['CABANG'].value_counts()
            _potensi = bund['Nilai Aksesoris'].mean() if n_bund else 0
            _anb.append(('aksi', f'{len(peluang):,} nota belum ditawari aksesoris',
                         f"Terbanyak di <b>{_pc2.index[0]}</b> ({_pc2.iloc[0]:,} nota). "
                         f"Bila separuhnya berhasil dilekati aksesoris dengan nilai "
                         f"rata-rata {rp(_potensi)}, potensi tambahan omzet sekitar "
                         f"<b>{rp(len(peluang) * 0.5 * _potensi)}</b>. <b>Tindakan:</b> "
                         f"jadikan penawaran aksesoris bagian wajib dari alur serah terima."))

        if len(gc) >= 3:
            _ter = gc.head(1)
            _bwh2 = gc.tail(2)
            _anb.append(('perhatian', 'Jarak antar cabang lebar',
                         f"<b>{_ter.index[0]}</b> mencapai {_ter['Attach Rate %'].iloc[0]:.1f}%, "
                         f"sementara {' dan '.join(f'{i} ({r:.1f}%)' for i, r in zip(_bwh2.index, _bwh2['Attach Rate %']))}"
                         f". Selisihnya "
                         f"{_ter['Attach Rate %'].iloc[0] - _bwh2['Attach Rate %'].iloc[-1]:.1f} poin. "
                         f"<b>Tindakan:</b> kirim tim {_ter.index[0]} untuk berbagi cara "
                         f"menawarkan, karena selisih sebesar ini biasanya soal kebiasaan "
                         f"kerja, bukan daya beli pelanggan."))

        if len(gp):
            _low_p = gp.tail(3)
            _anb.append(('info', 'Penjual dengan attach rate terendah',
                         f"{', '.join(f'{i} ({r:.1f}%)' for i, r in zip(_low_p.index, _low_p['Attach Rate %']))} "
                         f"— dibanding tertinggi {gp['Attach Rate %'].iloc[0]:.1f}%. "
                         f"<b>Tindakan:</b> jadikan bahan pembinaan, bukan penalti."))

        _anb.append(('info', 'Sebagian besar bundling terjadi pada servis',
                     "Pasangan aksesoris terbanyak adalah <b>jasa</b> dan <b>sparepart</b>, "
                     "artinya peluang terbesar ada saat pelanggan mengambil unit yang "
                     "selesai diperbaiki — bukan saat pembelian unit baru."))
        panel_analisa(_anb)

        tombol_pdf("Dashboard Bundling Aksesoris", _potb, _metrik_b, _anb,
                   kpis=[{'label': 'Total Nota', 'value': nfid(tot_nota),
                          'sub': 'sesuai filter', 'warna': _PN},
                         {'label': 'Nota Bundling', 'value': nfid(n_bund),
                          'sub': f"attach rate {pctid(rate)}", 'warna': _PG},
                         {'label': 'Belum Bundling', 'value': nfid(len(peluang)),
                          'sub': 'peluang penawaran', 'warna': _PA},
                         {'label': 'Omzet Bundling', 'value': rp(bund['Nilai'].sum()),
                          'sub': 'dari nota bundling', 'warna': _PN},
                         {'label': 'Nilai Aksesoris', 'value': rp(bund['Nilai Aksesoris'].sum()),
                          'sub': 'dalam nota bundling', 'warna': _PG}],
                   ringkasan=(f"Dari {nfid(tot_nota)} nota, {nfid(n_bund)} memuat aksesoris "
                              f"bersama kategori lain (attach rate {pctid(rate)}). "
                              f"Masih ada {nfid(len(peluang))} nota yang belum ditawari "
                              f"aksesoris."),
                   metodologi=("Satu nota = kombinasi Cabang + Nomor Faktur. Nota bundling "
                               "= memuat minimal satu item AKSESORIS dan minimal satu item "
                               "kategori lain, total minimal 2 produk. Nilai memakai TOTAL "
                               "HARGA (omzet), belum dikurangi biaya."),
                   key='bund')

        with st.expander("ℹ️ Cara perhitungan & catatan"):
            st.write(
                "**Satu nota** = kombinasi **Cabang + Nomor Faktur**. Ini penting karena "
                "penomoran faktur berjalan sendiri-sendiri di tiap cabang — satu nomor "
                "seperti `MF-FP.3974` bisa dipakai beberapa cabang berbeda.\n\n"
                "**Nota bundling** = memuat minimal satu item berkategori AKSESORIS "
                "(termasuk penulisan `ACCESORIES`) **dan** minimal satu item kategori "
                "lain, dengan total minimal 2 produk pada nota tersebut.\n\n"
                "**Attach rate** = jumlah nota bundling dibagi seluruh nota.\n\n"
                "**Belum bundling (peluang)** = nota yang berisi barang/jasa tapi sama "
                "sekali tidak ada aksesoris. Ini kelompok yang paling relevan untuk "
                "didorong, karena pelanggannya sudah bertransaksi tapi belum ditawari "
                "aksesoris.\n\n"
                "Perlu diketahui: pasangan aksesoris paling banyak adalah **JASA** dan "
                "**SPAREPART**, artinya sebagian besar bundling terjadi pada transaksi "
                "servis — bukan penjualan unit baru. Angka ini memakai seluruh kategori "
                "non-aksesoris sesuai kesepakatan.\n\n"
                "Nilai yang dipakai adalah **TOTAL HARGA** (omzet), belum dikurangi biaya."
            )


# =============================================================================
# TAB 5: MATI TOTAL — SUCCESS RATE JADI NOTA
# =============================================================================
KUNCI_MATI_TOTAL = 'MATI TOTAL'
KOL_NO_KIRIM_JUAL = 'Nomor # Pengiriman Pesanan'
KOL_NO_KIRIM_SERVIS = 'NOMOR PENGIRIMAN PESANAN'


@st.cache_data(show_spinner="Menggabungkan data servis & penjualan...")
def gabung_mati_total(servis: pd.DataFrame, jual: pd.DataFrame) -> pd.DataFrame:
    """Gabungkan unit MATI TOTAL dengan nota penjualannya.

    Kunci penggabungan adalah pasangan **CABANG + NOMOR PENGIRIMAN PESANAN**,
    bukan nomornya saja. Ini wajib: nomor pengiriman berjalan sendiri-sendiri di
    tiap cabang, dan ada 29.158 nomor yang dipakai lebih dari satu cabang. Kalau
    digabung hanya lewat nomor, success rate ikut membengkak palsu (79% vs 46%
    yang sebenarnya).
    """
    if (servis is None or servis.empty or jual is None or jual.empty
            or KOL_NO_KIRIM_SERVIS not in servis.columns
            or KOL_NO_KIRIM_JUAL not in jual.columns):
        return pd.DataFrame()

    mt = servis[servis['KERUSAKAN'] == KUNCI_MATI_TOTAL].copy()
    if mt.empty:
        return mt
    mt['NO_KIRIM'] = mt[KOL_NO_KIRIM_SERVIS].astype(str).str.strip().str.upper()

    j = jual.copy()
    j['NO_KIRIM'] = j[KOL_NO_KIRIM_JUAL].astype(str).str.strip().str.upper()
    j = j[~j['NO_KIRIM'].isin(['', 'NAN', 'NONE', '-'])]
    if j.empty:
        nota = pd.DataFrame(columns=['CABANG', 'NO_KIRIM', 'N_NOTA', 'OMZET',
                                     'MODAL', 'TGL_NOTA', 'NO_FAKTUR'])
    else:
        nota = (j.groupby(['CABANG', 'NO_KIRIM'])
                 .agg(N_NOTA=('NO FAKTUR', 'nunique'),
                      OMZET=('TOTAL HARGA', 'sum'),
                      MODAL=('MODAL', 'sum'),
                      TGL_NOTA=('TGL', 'min'),
                      NO_FAKTUR=('NO FAKTUR', 'first'))
                 .reset_index())

    mt = mt.merge(nota, on=['CABANG', 'NO_KIRIM'], how='left')
    mt['JADI_NOTA'] = mt['N_NOTA'].notna()
    mt['OMZET'] = pd.to_numeric(mt['OMZET'], errors='coerce').fillna(0.0)
    mt['MODAL'] = pd.to_numeric(mt['MODAL'], errors='coerce').fillna(0.0)
    mt['LABA'] = mt['OMZET'] - mt['MODAL']
    if 'TGL_NOTA' in mt.columns:
        mt['JEDA_HARI'] = (pd.to_datetime(mt['TGL_NOTA'], errors='coerce')
                           - mt['TGL PENGIRIMAN']).dt.days
    else:
        mt['JEDA_HARI'] = pd.NA
    mt.attrs['batas_data_jual'] = jual['TGL'].min()
    return mt


with tab_mati:
    st.markdown("## Dashboard Mati Total — Tingkat Keberhasilan Jadi Nota")
    st.caption(
        "Menelusuri setiap unit dengan kerusakan **MATI TOTAL**: berapa yang "
        "akhirnya benar-benar menjadi nota penjualan, dan berapa yang lepas."
    )

    if sales.empty:
        st.warning(
            "Dashboard ini menggabungkan **dua sumber**: data servis (nomor "
            "pengiriman pesanan) dan data penjualan (nota). Data penjualan belum "
            "termuat, jadi tingkat keberhasilan belum bisa dihitung. Pastikan "
            "`data/penjualan.csv.gz` ada di repo, atau upload lewat sidebar."
        )
    else:
        mt_all = gabung_mati_total(data, sales)

        if mt_all.empty:
            st.info("Tidak ada transaksi berkerusakan MATI TOTAL pada data ini.")
        else:
            batas = mt_all.attrs.get('batas_data_jual', sales['TGL'].min())
            batas = pd.Timestamp(batas).normalize().replace(day=1)

            # Unit yang datang sebelum data penjualan dimulai mustahil ketemu
            # notanya — bukan karena gagal, tapi karena notanya tidak ada di
            # sumber. Kalau ikut dihitung, success rate jadi ngawur.
            mt_ukur = mt_all[mt_all['TGL PENGIRIMAN'] >= batas].copy()
            n_luar = len(mt_all) - len(mt_ukur)

            if mt_ukur.empty:
                st.warning("Belum ada unit MATI TOTAL pada rentang data penjualan.")
            else:
                # ---- masa tunggu wajar: unit baru masuk belum sempat jadi nota
                jeda = pd.to_numeric(
                    mt_ukur.loc[mt_ukur['JADI_NOTA'], 'JEDA_HARI'], errors='coerce'
                ).dropna()
                jeda_p90 = int(jeda.quantile(0.90)) if len(jeda) >= 30 else 21
                jeda_med = float(jeda.median()) if len(jeda) else 0.0
                tgl_akhir = mt_ukur['TGL PENGIRIMAN'].max()
                ambang_matang = tgl_akhir - pd.Timedelta(days=jeda_p90)

                if n_luar:
                    st.info(
                        f"ℹ️ {nfid(n_luar)} unit MATI TOTAL bertanggal sebelum "
                        f"**{batas.strftime('%d %B %Y')}** dikeluarkan dari "
                        f"perhitungan — data penjualan belum mencakup periode itu, "
                        f"sehingga notanya mustahil ditemukan dan akan terhitung "
                        f"'gagal' secara keliru."
                    )

                sub = apply_filters(mt_ukur, f_tahun, f_bulan, f_cabang)

                if sub.empty:
                    st.warning("Tidak ada data MATI TOTAL untuk filter yang dipilih.")
                else:
                    n_unit = len(sub)
                    n_jadi = int(sub['JADI_NOTA'].sum())
                    n_gagal = n_unit - n_jadi
                    rate = n_jadi / n_unit * 100 if n_unit else 0.0
                    omzet = float(sub['OMZET'].sum())
                    laba = float(sub['LABA'].sum())
                    rata_nota = (omzet / n_jadi) if n_jadi else 0.0
                    porsi_mt = (n_unit / len(apply_filters(data, f_tahun, f_bulan, f_cabang))
                                * 100) if len(apply_filters(data, f_tahun, f_bulan, f_cabang)) else 0.0

                    warna_rate = ('linear-gradient(135deg,#16a34a,#22c55e)' if rate >= 55
                                  else ('linear-gradient(135deg,#f59e0b,#fbbf24)'
                                        if rate >= 40
                                        else 'linear-gradient(135deg,#dc2626,#ef4444)'))
                    cards = [
                        {'label': 'Unit Mati Total', 'value': nfid(n_unit),
                         'sub': f"{pctid(porsi_mt)} dari seluruh transaksi",
                         'grad': 'linear-gradient(135deg,#3b5bfd,#5a72ff)'},
                        {'label': 'Jadi Nota', 'value': nfid(n_jadi),
                         'sub': 'unit yang menghasilkan penjualan',
                         'grad': 'linear-gradient(135deg,#16a34a,#22c55e)'},
                        {'label': 'Success Rate', 'value': pctid(rate),
                         'sub': 'jadi nota ÷ unit masuk', 'grad': warna_rate},
                        {'label': 'Tidak Jadi Nota', 'value': nfid(n_gagal),
                         'sub': f"{pctid(100 - rate)} — potensi hilang",
                         'grad': 'linear-gradient(135deg,#dc2626,#ef4444)'},
                        {'label': 'Omzet dari Mati Total', 'value': rp(omzet),
                         'sub': f"laba kotor {rp(laba)} (blm potong bagi hasil)",
                         'grad': 'linear-gradient(135deg,#7c3aed,#a855f7)'},
                        {'label': 'Rata-rata per Nota', 'value': rp(rata_nota),
                         'sub': f"jeda masuk→nota: {nfid(jeda_med, 0)} hari (median)",
                         'grad': 'linear-gradient(135deg,#0f8a82,#17a3a3)'},
                    ]
                    st.markdown(kpi_html(cards), unsafe_allow_html=True)
                    st.write("")

                    # =========================================================
                    # GRAFIK BATANG PER BULAN
                    # =========================================================
                    st.markdown("#### Rekap per Bulan — Jadi Nota vs Tidak")

                    th_opts = sorted(int(t) for t in sub['TAHUN'].dropna().unique())
                    if f_tahun != 'Semua Tahun':
                        th_pilih = int(f_tahun)
                    else:
                        th_pilih = st.selectbox(
                            "Tahun grafik", th_opts, index=len(th_opts) - 1,
                            format_func=lambda x: str(x), key='mt_tahun_grafik')

                    bl = sub[sub['TAHUN'] == th_pilih].copy()
                    if bl.empty:
                        st.caption("Tidak ada data untuk tahun tersebut.")
                    else:
                        g = (bl.groupby(bl['TGL PENGIRIMAN'].dt.month)
                               .agg(unit=('JADI_NOTA', 'size'),
                                    jadi=('JADI_NOTA', 'sum'),
                                    omzet=('OMZET', 'sum'),
                                    akhir=('TGL PENGIRIMAN', 'max'))
                               .reindex(range(1, 13)))
                        g = g[g['unit'].notna()]
                        g['gagal'] = g['unit'] - g['jadi']
                        g['rate'] = g['jadi'] / g['unit'] * 100
                        g['mentah'] = g['akhir'] > ambang_matang
                        nama_bl = [BULAN_NAMES[int(i)] for i in g.index]
                        tanda = ['*' if m else '' for m in g['mentah']]
                        label_bl = [f"{n}{t}" for n, t in zip(nama_bl, tanda)]

                        fig = go.Figure()
                        fig.add_bar(x=label_bl, y=g['jadi'], name='Jadi nota',
                                    marker_color='#16a34a',
                                    text=[nfid(v) for v in g['jadi']],
                                    textposition='inside',
                                    hovertemplate='%{x}<br>Jadi nota: %{y:,.0f}<extra></extra>')
                        fig.add_bar(x=label_bl, y=g['gagal'], name='Tidak jadi nota',
                                    marker_color='#dc2626',
                                    text=[nfid(v) for v in g['gagal']],
                                    textposition='inside',
                                    hovertemplate='%{x}<br>Tidak jadi: %{y:,.0f}<extra></extra>')
                        fig.add_trace(go.Scatter(
                            x=label_bl, y=g['rate'], name='Success rate',
                            mode='lines+markers+text', yaxis='y2',
                            line=dict(color='#1f3864', width=3),
                            marker=dict(size=9, color='#1f3864'),
                            text=[pctid(v) for v in g['rate']],
                            textposition='top center',
                            textfont=dict(size=10, color='#1f3864'),
                            hovertemplate='%{x}<br>Success rate: %{y:.1f}%<extra></extra>'))
                        fig.update_layout(
                            barmode='stack', height=430,
                            margin=dict(l=10, r=10, t=30, b=10),
                            yaxis=dict(title='Jumlah unit'),
                            yaxis2=dict(title='Success rate (%)', overlaying='y',
                                        side='right', range=[0, 100], showgrid=False),
                            legend=dict(orientation='h', yanchor='bottom', y=1.02,
                                        x=0),
                            plot_bgcolor='#ffffff')
                        st.plotly_chart(fig, use_container_width=True, key='mt_bulan')

                        if bool(g['mentah'].any()):
                            st.caption(
                                f"* Bulan bertanda bintang belum matang: unit yang "
                                f"baru masuk dalam {jeda_p90} hari terakhir umumnya "
                                f"belum sempat dinotakan (90% nota terbit dalam "
                                f"{jeda_p90} hari sejak unit masuk). Success rate "
                                f"bulan itu akan naik sendiri seiring waktu — jangan "
                                f"dijadikan dasar penilaian."
                            )

                        tb = pd.DataFrame({
                            'Bulan': nama_bl,
                            'Unit Mati Total': g['unit'].astype(int).values,
                            'Jadi Nota': g['jadi'].astype(int).values,
                            'Tidak Jadi': g['gagal'].astype(int).values,
                            'Success Rate': g['rate'].values,
                            'Omzet': g['omzet'].values,
                        })
                        with st.expander("Lihat tabel angka per bulan"):
                            st.dataframe(
                                tb.style.format({
                                    'Unit Mati Total': '{:,.0f}', 'Jadi Nota': '{:,.0f}',
                                    'Tidak Jadi': '{:,.0f}', 'Success Rate': '{:.1f}%',
                                    'Omzet': 'Rp {:,.0f}'}),
                                use_container_width=True, hide_index=True,
                                key='mt_tabel_bulan')

                    # =========================================================
                    # PER TEKNISI
                    # =========================================================
                    st.markdown("---")
                    st.markdown("#### Tingkat Keberhasilan per Teknisi")

                    gt = (sub.groupby('TEKNISI')
                             .agg(Unit=('JADI_NOTA', 'size'),
                                  Jadi=('JADI_NOTA', 'sum'),
                                  Omzet=('OMZET', 'sum'),
                                  Laba=('LABA', 'sum')))
                    gt['Tidak Jadi'] = gt['Unit'] - gt['Jadi']
                    gt['Success Rate'] = gt['Jadi'] / gt['Unit'] * 100
                    gt['Omzet / Unit'] = gt['Omzet'] / gt['Unit']

                    c1, c2 = st.columns([1, 1])
                    with c1:
                        min_unit = st.number_input(
                            "Minimal unit mati total agar teknisi ditampilkan",
                            min_value=1, max_value=200, value=20, step=5,
                            key='mt_min_unit',
                            help="Menyaring teknisi bervolume kecil supaya "
                                 "persentasenya tidak menyesatkan.")
                    with c2:
                        top_n = st.slider("Jumlah teknisi pada grafik", 5, 40, 15,
                                          key='mt_topn')

                    gt_layak = gt[gt['Unit'] >= min_unit].copy()
                    if gt_layak.empty:
                        st.caption(
                            f"Tidak ada teknisi dengan minimal {min_unit} unit "
                            f"mati total pada filter ini. Turunkan angkanya.")
                    else:
                        rata_rate = rate
                        gg = gt_layak.sort_values('Success Rate', ascending=False).head(top_n)
                        gg = gg.iloc[::-1]     # plotly bar horizontal dari bawah

                        figt = go.Figure()
                        figt.add_bar(
                            y=gg.index, x=gg['Jadi'], orientation='h',
                            name='Jadi nota', marker_color='#16a34a',
                            hovertemplate='%{y}<br>Jadi nota: %{x:,.0f}<extra></extra>')
                        figt.add_bar(
                            y=gg.index, x=gg['Tidak Jadi'], orientation='h',
                            name='Tidak jadi nota', marker_color='#dc2626',
                            hovertemplate='%{y}<br>Tidak jadi: %{x:,.0f}<extra></extra>')
                        for nm, r, u in zip(gg.index, gg['Success Rate'], gg['Unit']):
                            figt.add_annotation(
                                y=nm, x=u, text=f"  {pctid(r)}", showarrow=False,
                                xanchor='left', font=dict(
                                    size=11, color=('#16a34a' if r >= rata_rate
                                                    else '#c0392b')))
                        figt.update_layout(
                            barmode='stack', height=max(320, 26 * len(gg) + 90),
                            margin=dict(l=10, r=90, t=30, b=10),
                            xaxis=dict(title='Jumlah unit mati total'),
                            legend=dict(orientation='h', yanchor='bottom', y=1.02, x=0),
                            plot_bgcolor='#ffffff')
                        st.plotly_chart(figt, use_container_width=True, key='mt_teknisi')
                        st.caption(
                            f"Diurutkan dari success rate tertinggi. Angka di ujung "
                            f"batang berwarna hijau bila di atas rata-rata "
                            f"({pctid(rata_rate)}) dan merah bila di bawahnya. "
                            f"Panjang batang menunjukkan beban unit — teknisi dengan "
                            f"batang panjang dan porsi merah besar adalah prioritas.")

                        st.markdown("##### Rekap Lengkap per Teknisi")
                        gt_show = gt_layak.copy()
                        gt_show['CABANG'] = (sub.groupby('TEKNISI')['CABANG']
                                             .agg(lambda s: ', '.join(sorted(set(s))[:3]))
                                             .reindex(gt_show.index))
                        gt_show = gt_show[['CABANG', 'Unit', 'Jadi', 'Tidak Jadi',
                                           'Success Rate', 'Omzet', 'Laba',
                                           'Omzet / Unit']]
                        gt_show = gt_show.rename(columns={'Jadi': 'Jadi Nota'})
                        gt_show = gt_show.sort_values('Success Rate', ascending=False)
                        st.dataframe(
                            gt_show.style.format({
                                'Unit': '{:,.0f}', 'Jadi Nota': '{:,.0f}',
                                'Tidak Jadi': '{:,.0f}', 'Success Rate': '{:.1f}%',
                                'Omzet': 'Rp {:,.0f}', 'Laba': 'Rp {:,.0f}',
                                'Omzet / Unit': 'Rp {:,.0f}'}),
                            use_container_width=True, height=420, key='mt_tabel_teknisi')

                        st.download_button(
                            "⬇️ Unduh rekap teknisi mati total (CSV)",
                            data=gt_show.reset_index().to_csv(index=False).encode('utf-8-sig'),
                            file_name=f"mati_total_per_teknisi_{f_tahun}.csv",
                            mime="text/csv", key='mt_unduh_teknisi')

                    # =========================================================
                    # PER CABANG + SEBAB GAGAL
                    # =========================================================
                    st.markdown("---")
                    cc1, cc2 = st.columns([1, 1])

                    with cc1:
                        st.markdown("#### Success Rate per Cabang")
                        gc = (sub.groupby('CABANG')
                                 .agg(Unit=('JADI_NOTA', 'size'),
                                      Jadi=('JADI_NOTA', 'sum'),
                                      Omzet=('OMZET', 'sum')))
                        gc['Success Rate'] = gc['Jadi'] / gc['Unit'] * 100
                        gc = gc.sort_values('Success Rate', ascending=False)
                        figc = px.bar(
                            gc.reset_index(), x='Success Rate', y='CABANG',
                            orientation='h', text=gc['Success Rate'].map(pctid).values,
                            color='Success Rate', color_continuous_scale=
                            [(0, '#c0392b'), (0.5, '#e0b31f'), (1, '#16a34a')])
                        figc.update_layout(
                            height=max(320, 22 * len(gc) + 80),
                            margin=dict(l=10, r=10, t=20, b=10),
                            yaxis=dict(autorange='reversed', title=''),
                            coloraxis_showscale=False, plot_bgcolor='#ffffff')
                        figc.update_traces(textposition='outside', cliponaxis=False)
                        st.plotly_chart(figc, use_container_width=True, key='mt_cabang')

                    with cc2:
                        st.markdown("#### Ke Mana Unit yang Tidak Jadi Nota?")
                        gagal = sub[~sub['JADI_NOTA']]
                        if gagal.empty:
                            st.caption("Semua unit mati total menjadi nota.")
                        else:
                            gs = (gagal['STATUS PENGERJAAN'].fillna('(tanpa status)')
                                  .astype(str).str.strip().str.upper()
                                  .value_counts().head(10))
                            figs = px.bar(
                                x=gs.values, y=gs.index, orientation='h',
                                text=[f"{nfid(v)} ({pctid(v / len(gagal) * 100)})"
                                      for v in gs.values],
                                color_discrete_sequence=['#c0392b'])
                            figs.update_layout(
                                height=max(320, 26 * len(gs) + 80),
                                margin=dict(l=10, r=10, t=20, b=10),
                                xaxis=dict(title='Jumlah unit'),
                                yaxis=dict(autorange='reversed', title=''),
                                plot_bgcolor='#ffffff')
                            figs.update_traces(textposition='outside', cliponaxis=False)
                            st.plotly_chart(figs, use_container_width=True,
                                            key='mt_sebab')
                            st.caption(
                                "Status pengerjaan terakhir dari unit yang tidak "
                                "pernah menjadi nota. Kelompok CANCEL adalah "
                                "kehilangan nyata; kelompok PENDING masih bisa "
                                "diselamatkan kalau segera dikerjakan.")

                    # =========================================================
                    # DETAIL
                    # =========================================================
                    st.markdown("---")
                    with st.expander("🔎 Detail unit mati total"):
                        hanya_gagal = st.checkbox(
                            "Tampilkan hanya yang belum/tidak jadi nota", value=True,
                            key='mt_hanya_gagal')
                        dd = sub[~sub['JADI_NOTA']] if hanya_gagal else sub
                        kol = [KOL_NO_KIRIM_SERVIS, 'TGL PENGIRIMAN', 'CABANG',
                               'TEKNISI', 'STATUS PENGERJAAN', 'MERK UNIT',
                               'TIPE UNIT', 'NO_FAKTUR', 'TGL_NOTA', 'OMZET',
                               'JEDA_HARI']
                        kol = [c for c in kol if c in dd.columns]
                        dd = dd[kol].rename(columns={
                            KOL_NO_KIRIM_SERVIS: 'No Pengiriman',
                            'NO_FAKTUR': 'No Faktur', 'TGL_NOTA': 'Tgl Nota',
                            'JEDA_HARI': 'Jeda (hari)'})
                        q = st.text_input("Cari nomor / teknisi / cabang",
                                          key='mt_cari')
                        if q:
                            m = dd.apply(lambda r: q.upper() in ' '.join(
                                str(v) for v in r.values).upper(), axis=1)
                            dd = dd[m]
                        dd = dd.sort_values('TGL PENGIRIMAN', ascending=False)
                        st.caption(f"{nfid(len(dd))} unit (ditampilkan maks 1.000).")
                        st.dataframe(dd.head(1000), use_container_width=True,
                                     height=380, key='mt_detail')
                        st.download_button(
                            "⬇️ Unduh daftar ini (CSV)",
                            data=dd.to_csv(index=False).encode('utf-8-sig'),
                            file_name="mati_total_detail.csv", mime="text/csv",
                            key='mt_unduh_detail')

                    # =========================================================
                    # PERBANDINGAN PERIODE
                    # =========================================================
                    st.markdown("### 📊 Perbandingan Periode")
                    _mt_cab = (mt_ukur if f_cabang == 'Semua Cabang'
                               else mt_ukur[mt_ukur['CABANG'] == f_cabang])
                    _pot_mt = potong_periode(_mt_cab, 'TGL PENGIRIMAN')
                    _metrik_mt = [
                        ("Unit Mati Total", lambda d: len(d), lambda v: nfid(v), True),
                        ("Jadi Nota", lambda d: int(d['JADI_NOTA'].sum()),
                         lambda v: nfid(v), True),
                        ("Success Rate", lambda d: (d['JADI_NOTA'].sum() / len(d) * 100)
                         if len(d) else 0, lambda v: pctid(v), True),
                        ("Omzet Mati Total", lambda d: float(d['OMZET'].sum()),
                         lambda v: rp(v), True),
                        ("Rata-rata per Nota",
                         lambda d: (float(d['OMZET'].sum()) / int(d['JADI_NOTA'].sum()))
                         if int(d['JADI_NOTA'].sum()) else 0, lambda v: rp(v), True),
                    ]
                    render_banding(_pot_mt, _metrik_mt, key_prefix='mt')
                    st.caption(
                        f"⚠️ Khusus baris **Success Rate**, bulan berjalan hampir "
                        f"selalu tampak lebih rendah — sebagian unitnya baru masuk "
                        f"dan belum sampai tahap nota (90% nota terbit dalam "
                        f"{jeda_p90} hari). Untuk penilaian yang adil, pakai bulan "
                        f"yang sudah lewat minimal {jeda_p90} hari.")

                    # =========================================================
                    # ANALISA
                    # =========================================================
                    st.markdown("### 🧭 Analisa & Tindak Lanjut")

                    # tren success rate pada bulan-bulan yang sudah matang
                    _mat = _mt_cab[_mt_cab['TGL PENGIRIMAN'] <= ambang_matang]
                    _tren = (_mat.groupby(_mat['TGL PENGIRIMAN'].dt.to_period('M'))
                                 .agg(u=('JADI_NOTA', 'size'), j=('JADI_NOTA', 'sum')))
                    _tren['r'] = _tren['j'] / _tren['u'] * 100
                    _tren = _tren[_tren['u'] >= 30]

                    _items = []

                    if len(_tren) >= 2:
                        _awal, _akhir = float(_tren['r'].iloc[0]), float(_tren['r'].iloc[-1])
                        _sel = _akhir - _awal
                        _arah = ("naik" if _sel > 1 else ("turun" if _sel < -1 else "datar"))
                        _jenis = ('baik' if _sel > 1 else ('aksi' if _sel < -1 else 'info'))
                        _items.append((
                            _jenis, f"Tren success rate {_arah}",
                            f"Pada bulan-bulan yang sudah matang, success rate bergerak "
                            f"dari <b>{pctid(_awal)}</b> ({_tren.index[0]}) ke "
                            f"<b>{pctid(_akhir)}</b> ({_tren.index[-1]}) — selisih "
                            f"<b>{pctid(abs(_sel))}</b>. Rata-rata sepanjang periode "
                            f"matang: <b>{pctid(float(_tren['j'].sum() / _tren['u'].sum() * 100))}</b>."))

                    _hilang = n_gagal * (omzet / n_jadi if n_jadi else 0)
                    _items.append((
                        'aksi', "Nilai yang belum tertutup",
                        f"<b>{nfid(n_gagal)} unit</b> mati total tidak menjadi nota. "
                        f"Dengan nilai rata-rata <b>{rp(rata_nota)}</b> per nota, itu "
                        f"setara peluang <b>{rp(_hilang)}</b> yang tidak terealisasi. "
                        f"Angka ini adalah batas atas — sebagian unit memang tidak "
                        f"layak diperbaiki — tapi tetap menunjukkan besarnya taruhan "
                        f"pada satu jenis kerusakan saja."))

                    if not gagal.empty:
                        _st = (gagal['STATUS PENGERJAAN'].fillna('(tanpa status)')
                               .astype(str).str.upper())
                        _n_cancel = int(_st.str.startswith('CANCEL').sum())
                        _n_pending = int(_st.str.startswith('PENDING').sum())
                        _items.append((
                            'perhatian', "Yang masih bisa diselamatkan",
                            f"Dari unit yang belum jadi nota, <b>{nfid(_n_cancel)}</b> "
                            f"({pctid(_n_cancel / len(gagal) * 100)}) sudah berstatus "
                            f"CANCEL — ini kehilangan yang sudah terjadi. Namun "
                            f"<b>{nfid(_n_pending)}</b> masih PENDING dan belum "
                            f"tertutup: unit inilah yang paling layak dikejar minggu "
                            f"ini, karena keputusannya belum final di tangan pelanggan."))

                    if not gt_layak.empty and len(gt_layak) >= 4:
                        _srt = gt_layak.sort_values('Success Rate')
                        _bwh = _srt.head(max(1, len(_srt) // 4))
                        _ats = _srt.tail(max(1, len(_srt) // 4))
                        _r_bwh = float(_bwh['Jadi'].sum() / _bwh['Unit'].sum() * 100)
                        _r_ats = float(_ats['Jadi'].sum() / _ats['Unit'].sum() * 100)
                        _tambah = (_r_ats - _r_bwh) / 100 * float(_bwh['Unit'].sum())
                        _items.append((
                            'aksi', "Jurang antar teknisi adalah peluang terbesar",
                            f"Kuartil teratas ({nfid(len(_ats))} teknisi) menutup "
                            f"<b>{pctid(_r_ats)}</b> unit mati total, sementara kuartil "
                            f"terbawah ({nfid(len(_bwh))} teknisi) hanya "
                            f"<b>{pctid(_r_bwh)}</b> — beda <b>{pctid(_r_ats - _r_bwh)}</b> "
                            f"pada jenis kerusakan yang sama. Kalau kelompok terbawah "
                            f"naik ke tingkat kelompok teratas, tambahannya sekitar "
                            f"<b>{nfid(_tambah, 0)} nota</b> (± {rp(_tambah * rata_nota)}). "
                            f"Terendah saat ini: <b>{_srt.index[0]}</b> "
                            f"({pctid(float(_srt['Success Rate'].iloc[0]))} dari "
                            f"{nfid(int(_srt['Unit'].iloc[0]))} unit). Selisih sebesar "
                            f"ini biasanya soal cara mendiagnosa dan menjelaskan "
                            f"estimasi ke pelanggan, bukan soal keberuntungan."))

                    if len(gc) >= 2:
                        _gc = gc[gc['Unit'] >= 30].sort_values('Success Rate')
                        if len(_gc) >= 2:
                            _items.append((
                                'perhatian', "Cabang yang perlu ditengok",
                                f"Cabang <b>{_gc.index[0]}</b> hanya menutup "
                                f"<b>{pctid(float(_gc['Success Rate'].iloc[0]))}</b> dari "
                                f"{nfid(int(_gc['Unit'].iloc[0]))} unit mati total, "
                                f"sementara <b>{_gc.index[-1]}</b> mencapai "
                                f"<b>{pctid(float(_gc['Success Rate'].iloc[-1]))}</b>. "
                                f"Perbedaan sebesar ini antar cabang biasanya berasal "
                                f"dari ketersediaan sparepart atau cara admin "
                                f"menyampaikan estimasi biaya — dua hal yang bisa "
                                f"disamakan tanpa menambah biaya."))

                    _items.append((
                        'info', "Cara angka ini dihitung",
                        f"Satu unit dianggap <b>berhasil</b> bila nomor pengiriman "
                        f"pesanannya muncul sebagai nota di data penjualan, "
                        f"dicocokkan lewat pasangan <b>Cabang + Nomor Pengiriman "
                        f"Pesanan</b>. Pasangan ini wajib karena penomoran berjalan "
                        f"sendiri-sendiri di tiap cabang. Sebagai uji kewajaran: "
                        f"unit berstatus DONE tercocokkan sekitar 92%, sedangkan "
                        f"CANCEL hanya 2% — persis seperti yang diharapkan bila "
                        f"kuncinya benar."))

                    panel_analisa(_items)

                    tombol_pdf(
                        "Dashboard Mati Total", _pot_mt, _metrik_mt,
                        temuan=[(j, ju, re.sub('<[^>]+>', '', isi))
                                for j, ju, isi in _items],
                        kpis=[{'label': 'Unit Mati Total', 'value': nfid(n_unit),
                               'sub': f"{pctid(porsi_mt)} dari seluruh transaksi",
                               'warna': _PN},
                              {'label': 'Jadi Nota', 'value': nfid(n_jadi),
                               'sub': 'menghasilkan penjualan', 'warna': _PG},
                              {'label': 'Success Rate', 'value': pctid(rate),
                               'sub': 'jadi nota / unit masuk',
                               'warna': (_PG if rate >= 55 else
                                         (_PA if rate >= 40 else _PR))},
                              {'label': 'Tidak Jadi Nota', 'value': nfid(n_gagal),
                               'sub': f"{pctid(100 - rate)} dari unit masuk",
                               'warna': _PR},
                              {'label': 'Omzet Mati Total', 'value': rp(omzet),
                               'sub': f"laba kotor {rp(laba)}", 'warna': _PN},
                              {'label': 'Rata-rata per Nota', 'value': rp(rata_nota),
                               'sub': f"jeda {nfid(jeda_med, 0)} hari (median)",
                               'warna': _PA}],
                        metodologi=(
                            f"Unit MATI TOTAL diambil dari kolom KERUSAKAN UTAMA pada "
                            f"data servis, lalu dicocokkan ke data penjualan memakai "
                            f"pasangan CABANG + NOMOR PENGIRIMAN PESANAN. Unit "
                            f"bertanggal sebelum {batas.strftime('%d %B %Y')} "
                            f"dikeluarkan karena data penjualan belum mencakup periode "
                            f"tersebut ({nfid(n_luar)} unit). Bulan berjalan ditandai "
                            f"belum matang bila unitnya masuk dalam {jeda_p90} hari "
                            f"terakhir, sebab 90% nota terbit dalam rentang itu "
                            f"(median {nfid(jeda_med, 0)} hari)."),
                        ringkasan=(
                            f"Dari {nfid(n_unit)} unit mati total, {nfid(n_jadi)} "
                            f"({pctid(rate)}) berhasil menjadi nota senilai "
                            f"{rp(omzet)}. Sisanya {nfid(n_gagal)} unit tidak tertutup."),
                        key='mt')

                    with st.expander("ℹ️ Cara perhitungan & catatan"):
                        st.write(
                            f"**Yang dihitung.** Setiap baris transaksi servis dengan "
                            f"KERUSAKAN UTAMA = `MATI TOTAL`. Baris kembar sudah "
                            f"dibuang lebih dulu saat data dimuat.\n\n"
                            f"**Kunci penggabungan.** `CABANG` + `NOMOR PENGIRIMAN "
                            f"PESANAN` (servis) dicocokkan dengan `CABANG` + "
                            f"`Nomor # Pengiriman Pesanan` (penjualan). Nomornya saja "
                            f"tidak cukup — ada 29.158 nomor yang dipakai lebih dari "
                            f"satu cabang, dan menggabungkannya secara global membuat "
                            f"success rate melonjak palsu ke 79%.\n\n"
                            f"**Batas periode.** Data penjualan mulai "
                            f"{batas.strftime('%d %B %Y')}. Unit servis sebelum "
                            f"tanggal itu dikeluarkan ({nfid(n_luar)} unit), karena "
                            f"notanya memang tidak ada di sumber — bukan karena gagal.\n\n"
                            f"**Masa tunggu.** Jeda dari unit masuk sampai terbit nota: "
                            f"median {nfid(jeda_med, 0)} hari, 90% selesai dalam "
                            f"{jeda_p90} hari. Karena itu unit yang baru masuk "
                            f"{jeda_p90} hari terakhir belum bisa dinilai.\n\n"
                            f"**Omzet.** Seluruh isi nota yang tertaut ke unit "
                            f"tersebut — termasuk sparepart dan jasa — memakai kolom "
                            f"TOTAL HARGA. Modal memakai HARGA BELI yang di sumber ini "
                            f"sudah berupa total per baris.\n\n"
                            f"**Soal angka laba.** Nilainya tinggi (sekitar 84%) karena "
                            f"sebagian besar nota mati total berisi **jasa**, dan baris "
                            f"jasa nyaris tidak punya HARGA BELI. Jadi ini laba kotor "
                            f"sebelum dipotong bagi hasil teknisi dan biaya operasional "
                            f"— bukan keuntungan bersih. Untuk perhitungan bagi hasil, "
                            f"lihat tab **Omzet & Bagi Hasil Teknisi**."
                        )
