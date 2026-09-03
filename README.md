# M-Flash Dashboard Builder — versi 4.0 (11 slide)

Aplikasi untuk mengubah file mentah Accurate menjadi dashboard interaktif **dan** presentasi
PowerPoint 11 slide yang mengikuti template standar weekly meeting M-Flash (latar, logo, warna,
dan tata letak yang sama persis).

---

## 1. Cara menjalankan (sekali setup)

**Windows** — klik dua kali `jalankan.bat`
**macOS** — klik dua kali `jalankan.command`

Pertama kali dijalankan, aplikasi mengunduh komponen yang dibutuhkan (butuh internet, ±2 menit).
Setelah itu browser terbuka otomatis di `http://localhost:8501`.

Kalau lebih suka manual:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Syarat: Python 3.9 atau lebih baru.

---

### Kalau hasil PPT terasa "tidak berubah"

1. **Hentikan aplikasi lama** — tutup jendela terminal/CMD yang sedang menjalankan Streamlit
   (tekan `Ctrl + C`), lalu jalankan lagi dari folder hasil ekstrak yang baru.
2. Pastikan sidebar bertuliskan **Versi 4.0 · 11 slide**. Kalau masih versi lama, berarti aplikasi
   masih dijalankan dari folder lama.
3. Setelah mengubah isian, klik **Buat PPTX** lagi — tombol unduh sengaja dikosongkan setiap kali
   ada perubahan supaya file lama tidak ikut terunduh.
4. Halaman **cover** file PPT memuat tulisan `Dibuat otomatis dari dashboard v4.0` —
   cara tercepat memastikan file berasal dari versi terbaru.
5. Untuk memastikan lewat properti file: klik kanan file → *Properties/Get Info* → bagian
   *Comments* tertulis `M-Flash Dashboard Builder v4.0 — 11 slide`.

### Menjalankan di Streamlit Cloud

Struktur repositori harus **persis** seperti ini — `app.py` di akar, dan folder `mflash`
serta `assets` ikut ter-upload:

```
app.py
requirements.txt
runtime.txt
cek_lingkungan.py
TEMPLATE_STRUKTUR_ORGANISASI.xlsx
mflash/
    __init__.py   loader.py   metrics.py   context.py
    charts.py     theme.py    deck.py      template.py
assets/
    bg.jpg   logo_mflash.png   logo_madinah.png
```

Kesalahan paling sering: folder `mflash` atau `assets` tidak ikut ter-push (mengunggah lewat
tombol *Add file → Upload files* di GitHub kadang melewatkan isi subfolder — seret seluruh
foldernya, atau gunakan `git add -A`). Gejalanya `ImportError` pada baris
`from mflash import ...`.

Setelah mengubah isi repositori, buka **Manage app → Reboot app** agar Streamlit Cloud
memasang ulang dependensinya.

Untuk memeriksa kelengkapan di komputer sendiri:

```bash
python cek_lingkungan.py
```

Skrip itu menyebut satu per satu file yang hilang dan pustaka yang gagal dimuat. Versi 3.2 juga
menampilkan diagnosa serupa langsung di halaman aplikasi bila impor gagal, lengkap dengan daftar
file yang terbaca — jadi penyebabnya terlihat tanpa perlu membuka log.

## 2. Cara pakai

1. Unggah file mentah di panel kiri (boleh beberapa sekaligus):
   - `rincian_pengiriman_pesanan_*.xlsx` → **unit masuk** per hari/bulan, klasifikasi Done, Pending, Cancel
   - `rincian_faktur_penjualan_*.xlsx` → **penjualan** per hari/bulan, breakdown kategori, voucher, bagi hasil
2. Atur filter: tahun, bulan, kategori penjualan, dan dimensi pembanding (cabang/teknisi/admin).
3. Telusuri dashboard di tab yang tersedia.
4. Isi tab **Slide manual**: 4 goal, foto Measure Activity, foto AR, foto + kesimpulan
   Improvement & Efficiency, foto To Do List, tabel Support Needs, dan tabel komitmen.
5. Buka tab **⬇️ Unduh PPT** → klik *Buat PPTX* → *Unduh PPTX*.

**Format yang diterima:** `.xlsx`, `.xls`, `.xlsm`, `.csv`, `.tsv`, `.txt`, **`.csv.gz`**, `.zip`,
dan `.parquet`. Untuk CSV, pemisah kolom (`,` `;` tab `|`) dan encoding dideteksi otomatis, jadi
file gabungan hasil ekspor sistem lain tetap terbaca.

**Beberapa berkas sejenis otomatis digabung.** Unggah sekaligus file gabungan per bulan atau per
cabang — aplikasi menyatukannya sebelum menghitung, dan panel kiri menyebutkan berapa berkas yang
tergabung. Untuk data pengiriman pesanan, baris yang identik tetap dihitung satu unit meskipun
berasal dari berkas berbeda.

Aplikasi membaca kolom berdasarkan nama, jadi ekspor Accurate bulan depan bisa langsung dipakai
tanpa mengubah apa pun.

---

## 3. Isi deck yang dihasilkan (11 slide)

| # | Slide | Sumber |
|---|-------|--------|
| 1 | Cover | otomatis + judul/penyaji |
| 2 | Pencapaian Goal — 4 gauge: Gross Profit, Omset Aksesoris, Tingkat Kepuasan Pelanggan, Google Ulasan | input manual |
| 3 | Ringkasan kinerja | pengiriman pesanan |
| 4 | **Perkembangan Pekanan — total transaksi & omzet per pekan** | pengiriman + faktur |
| 5 | Measure Activity | unggah foto |
| 6 | AR | unggah foto |
| 7 | Improvement & Efficiency | 4 foto + kesimpulan teks |
| 8 | To Do List | unggah foto |
| 9 | Support Needs — No, Divisi, Needs | tabel manual |
| 10 | Komitmen — pencapaian, komitmen, target | input manual |
| 11 | Penutup | otomatis |

### Perkembangan pekanan

Satu pekan dihitung **Senin–Minggu**. Rentangnya mengikuti filter bulan di panel kiri — pilih
Juli dan Agustus untuk melihat perkembangan dua bulan itu.

Yang ditampilkan: total transaksi dan total omzet seluruh pekan, rata-rata per pekan, angka pekan
terakhir beserta perubahannya terhadap pekan sebelumnya (hijau bila naik, merah bila turun), lalu
dua grafik batang — jumlah transaksi per pekan dan omzet per pekan dalam juta rupiah.

Transaksi dihitung dari file pengiriman pesanan, omzet dari file faktur penjualan. Bila file
faktur belum diunggah, grafik omzet diganti keterangan singkat. Pekan pertama dan terakhir bisa
belum genap tujuh hari karena mengikuti rentang data yang ada — catatan ini juga tercetak di
kaki slide.

Dashboard punya tab **Perkembangan Pekanan** sendiri dengan grafik gabungan (batang transaksi +
garis omzet) dan tabel lengkap berikut persentase perubahan tiap pekan.

## 4. Aturan perhitungan

- **Unit unik**: baris yang identik pada file pengiriman pesanan (satu nota tercetak berulang)
  dihitung satu. 46.612 baris → 26.280 unit unik pada data contoh.
- **Status**: `DONE*`/`COMPLAIN DONE` → Done · `CANCEL*` → Cancel · `PENDING*` → Pending · sisanya Lainnya.
- **Perbandingan bulan**: bulan berjalan dipotong sampai H-1 dari tanggal data terakhir, dan bulan
  sebelumnya dipotong pada tanggal yang sama agar setara. Pending tidak ikut dibandingkan karena
  angkanya kondisi terkini, bukan kejadian bulan tersebut.
- **Laba kotor** = `TOTAL HARGA` − `HARGA BELI`. Kategori JASA umumnya bermodal nol, sehingga
  marginnya tampil mendekati 100% — biaya tenaga kerja belum dibebankan di sana.
- **Bagi hasil**: Interface 20% · Normal 30% · Mati Total 32% · Promo 60% · Lainnya 30%.
  Bila dua kata kunci muncul bersamaan, Normal diprioritaskan.

---

## 5. Membuat PPT tanpa membuka aplikasi

```bash
python buat_ppt.py "rincian_pengiriman_pesanan.xlsx" "rincian_faktur_penjualan.xlsx" \
  --tahun 2026 --lingkup "Cabang Klender" \
  --judul "WEEKLY MEETING - PEKAN 6" --penyaji "NAMA – JABATAN" \
  --out WEEKLY_PEKAN_6.pptx --isi isi.json
```

`isi.json` (opsional) untuk mengisi slide manual:

```json
{
  "goals": [
    {"nama": "GROSS PROFIT", "nilai": 84.31, "ket": ""},
    {"nama": "OMSET AKSESORIS", "nilai": 59.5, "ket": ""},
    {"nama": "TINGKAT KEPUASAN PELANGGAN", "nilai": 124.0, "ket": ""}
  ],
  "komitmen": [{"pencapaian": "", "komitmen": "", "target": ""}],
  "foto_measure": ["foto/measure1.jpg"],
  "foto_ar": ["foto/ar1.jpg"]
}
```

---

## 6. Struktur file

```
app.py              antarmuka Streamlit
buat_ppt.py         mode baris perintah
mflash/loader.py    baca & bersihkan Excel
mflash/metrics.py   seluruh perhitungan
mflash/context.py   perangkai angka + input manual
mflash/deck.py      penyusun 21 slide
mflash/charts.py    grafik native PowerPoint
mflash/theme.py     warna, font, kartu, header (dari template)
assets/             latar & logo diambil dari template asli
```

Untuk mengubah warna atau tata letak, sunting `mflash/theme.py` dan `mflash/deck.py`.
