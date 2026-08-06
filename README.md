# Dashboard Service Cabang (Offline, Python)

Aplikasi dashboard berbasis Python (Streamlit) yang jalan **sepenuhnya offline**
di komputer sendiri — tidak butuh internet, tidak butuh server, dan bisa
di-update sendiri setiap kali ada data baru cukup dengan upload ulang file
Excel dari dalam aplikasi.

Berisi 2 dashboard dalam 1 aplikasi (tab):

1. **Dashboard Utama** — status pengerjaan (Done/Pending/Cancel/Lainnya),
   tren tahunan, rata-rata transaksi/hari, rekap per cabang.
2. **Dashboard Pending** — breakdown teknisi dengan pending terbanyak dan
   jenis kerusakan yang paling sering menumpuk jadi pending, lengkap dengan
   persentase.

Semua angka mengikuti aturan yang sama seperti versi HTML sebelumnya: baris
yang seluruh kolomnya identik dihitung sebagai 1 transaksi.

## 1. Install Python (sekali saja)

Pastikan Python 3.9 ke atas sudah terpasang. Cek dengan:

```bash
python3 --version
```

Kalau belum ada, download dari https://www.python.org/downloads/

## 2. Install library yang dibutuhkan

Buka Terminal (Mac) / Command Prompt (Windows), masuk ke folder
`dashboard_app` ini, lalu jalankan:

```bash
pip install -r requirements.txt
```

Tunggu sampai selesai (butuh internet untuk sekali ini saja, saat install).

## 3. Jalankan dashboard

Masih di folder yang sama, jalankan:

```bash
streamlit run app.py
```

Browser akan otomatis terbuka ke `http://localhost:8501` — itulah
dashboard-nya. Semua proses jalan lokal di komputer, tidak terkirim ke
mana-mana.

## 4. Upload data

Di panel kiri, klik **"Upload file Excel"** dan pilih file data (format harus
sama seperti `Gabungan_Semua_Cabang.xlsx`: satu sheet per cabang, dengan
kolom seperti `NOMOR PENGIRIMAN PESANAN`, `TGL PENGIRIMAN`,
`STATUS PENGERJAAN`, `NAMA TEKNISI`, `KERUSAKAN UTAMA`, dst).

Setelah upload, dashboard otomatis menghitung ulang. Gunakan filter Tahun /
Bulan / Cabang di panel kiri untuk mempersempit tampilan.

## 5. Cara Update Data (Offline)

Dashboard ini **tidak menyimpan data sendiri** — setiap kali dibuka, dia
membaca ulang file Excel yang Anda upload. Jadi "update database" artinya
cukup: siapkan file Excel terbaru, lalu upload ulang. Tidak ada instalasi
ulang, tidak ada ubah kode.

### Skenario A — Anda punya file export baru dari sistem (paling disarankan)

Kalau sumber datanya adalah export/laporan dari sistem service (seperti
`Gabungan_Semua_Cabang.xlsx` yang sudah ada), biasanya setiap kali export
baru sudah otomatis berisi data lama **+** data baru dalam satu file.

Langkah-langkahnya:

1. Export ulang laporan dari sistem seperti biasa, mencakup rentang tanggal
   terbaru (bisa dari awal lagi atau digabung dengan data sebelumnya).
2. Pastikan hasil export masih format yang sama: **satu sheet per cabang**,
   dengan nama kolom persis seperti sebelumnya (`NOMOR PENGIRIMAN PESANAN`,
   `TGL PENGIRIMAN`, `STATUS PENGERJAAN`, `NAMA TEKNISI`,
   `KERUSAKAN UTAMA`, dst).
3. Buka dashboard (`streamlit run app.py` kalau belum jalan).
4. Di panel kiri, klik **"Browse files"** pada uploader dan pilih file
   export terbaru — ini akan **menimpa** file lama yang sedang dibuka.
5. Tunggu proses "Membaca & memproses file Excel..." selesai (bisa 30–90
   detik untuk file besar). Semua angka, filter Tahun/Bulan/Cabang, dan
   ranking otomatis terhitung ulang dari file baru.

Tidak perlu khawatir soal data dobel: baris yang isinya identik persis tetap
otomatis dihitung sebagai 1 transaksi, jadi meng-upload file yang berisi
gabungan data lama + baru itu aman.

### Skenario B — Anda menambah data secara manual di Excel

Kalau tidak ada sistem export dan data ditambah manual:

1. Buka file Excel sumber (mis. `Gabungan_Semua_Cabang.xlsx`) langsung di
   Excel/Google Sheets.
2. Masuk ke **sheet cabang yang sesuai** (nama sheet = nama cabang, ini yang
   dipakai dashboard sebagai label Cabang).
3. Tambahkan baris baru **di bawah baris terakhir**, isi setiap kolom sesuai
   urutan header yang sudah ada — jangan mengubah nama header, jangan
   menyisipkan kolom baru di tengah.
4. Pastikan kolom `TGL PENGIRIMAN` diisi sebagai **tanggal (Date)**, bukan
   teks biasa — kalau formatnya teks, filter Tahun/Bulan di dashboard tidak
   akan mengenali baris tsb.
5. Simpan file (tetap format `.xlsx`, jangan `.xls` atau `.csv`).
6. Upload ulang file tsb ke dashboard seperti Skenario A langkah 3–5.

### Hal yang perlu dihindari

- Jangan mengganti/mengetik ulang nama kolom (header) — kalau berbeda dari
  aslinya, dashboard akan menampilkan pesan error "kolom tidak ditemukan".
- Jangan mengganti nama sheet cabang secara tiba-tiba tanpa alasan — nama
  sheet itu yang muncul sebagai pilihan filter **Cabang**.
- Kalau menambah cabang baru, cukup buat sheet baru dengan nama cabang tsb
  dan header kolom yang sama — dashboard otomatis mendeteksinya tanpa
  perlu ubah kode.

## Menghentikan aplikasi

Kembali ke Terminal/Command Prompt tempat `streamlit run app.py` dijalankan,
tekan `Ctrl + C`.

---

**Catatan:** file besar (ratusan ribu baris, puluhan MB) bisa butuh waktu
sekitar 30–90 detik untuk diproses saat pertama kali diupload. Setelah itu,
mengganti filter akan terasa instan karena data sudah di-cache di memori.
