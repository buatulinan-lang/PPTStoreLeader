# M-Flash Dashboard Builder — versi 3.1 (21 slide)

Aplikasi untuk mengubah file mentah Accurate menjadi dashboard interaktif **dan** presentasi
PowerPoint 21 slide yang mengikuti template standar weekly meeting M-Flash (latar, logo, warna,
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
2. Pastikan sidebar bertuliskan **Versi 3.1 · 21 slide**. Kalau masih versi lama, berarti aplikasi
   masih dijalankan dari folder lama.
3. Setelah mengubah isian, klik **Buat PPTX** lagi — tombol unduh sengaja dikosongkan setiap kali
   ada perubahan supaya file lama tidak ikut terunduh.
4. Untuk memastikan file hasil unduhan: klik kanan file → *Properties/Get Info* → bagian
   *Comments* tertulis `M-Flash Dashboard Builder v3.1 — 21 slide`.

## 2. Cara pakai

1. Unggah file mentah di panel kiri (boleh dua-duanya sekaligus):
   - `rincian_pengiriman_pesanan_*.xlsx` → **unit masuk** per hari/bulan, klasifikasi Done, Pending, Cancel
   - `rincian_faktur_penjualan_*.xlsx` → **penjualan** per hari/bulan, breakdown kategori, voucher, bagi hasil
2. Atur filter: tahun, bulan, kategori penjualan, dan dimensi pembanding (cabang/teknisi/admin).
3. Telusuri dashboard di tab yang tersedia.
4. Isi tab **Slide manual**: 4 goal, tabel struktur organisasi, foto Measure Activity, foto AR,
   dan tabel komitmen.
5. Buka tab **⬇️ Unduh PPT** → klik *Buat PPTX* → *Unduh PPTX*.

Aplikasi membaca kolom berdasarkan nama, jadi ekspor Accurate bulan depan bisa langsung dipakai
tanpa mengubah apa pun.

---

## 3. Isi deck yang dihasilkan (21 slide)

| # | Slide | Sumber |
|---|-------|--------|
| 1 | Cover | otomatis + judul/penyaji |
| 2 | Pencapaian Goal — 4 gauge: Gross Profit, Omset Aksesoris, Tingkat Kepuasan Pelanggan, Google Ulasan | input manual |
| 3 | Ringkasan kinerja | pengiriman pesanan |
| 4 | Komposisi status pengerjaan | pengiriman pesanan |
| 5 | Bulan berjalan vs bulan sebelumnya (dibandingkan setara) | pengiriman + faktur |
| 6 | Rekap unit masuk harian | pengiriman pesanan |
| 7 | Hari dengan unit masuk tertinggi | pengiriman pesanan |
| 8 | Kinerja per cabang/teknisi | pengiriman pesanan |
| 9–11 | Rincian Pending, Done, Cancel | pengiriman pesanan |
| 12 | Penjualan — modal, omzet & laba | faktur penjualan |
| 13 | Rekap penjualan harian & bulanan | faktur penjualan |
| 14 | Penjualan per kategori (service HP, service laptop, penjualan HP, dll) | faktur penjualan |
| 15 | Voucher | faktur penjualan |
| 16 | Bagi hasil teknisi | faktur penjualan |
| 17 | Struktur organisasi | tabel nama & jabatan |
| 18 | Measure Activity | unggah foto |
| 19 | AR | unggah foto |
| 20 | Komitmen — pencapaian, komitmen, target | input manual |
| 21 | Penutup | otomatis |

### Struktur organisasi

Cukup isi nama lengkap dan jabatan; bagan disusun otomatis berdasarkan kata pada kolom jabatan:

```
                 Ustadz Pembina Cabang
                          |
                    Store Leader
                          |
   ┌──────┬──────┬────────┼────────┬────────────┬──────┐
 Service Aksesoris Pengadaan Penyewaan Maintenance  ISP     (Supervisor)
   |                   └────────┴────────────┴──────┘
 Admin, Sales,                    Sales Corporate
 Teknisi, dll
```

- Jabatan mengandung **"Ustadz"** atau **"Pembina"** → puncak bagan
- **"Store Leader"** (atau Store Manager / Kepala Toko / Pimpinan) → di bawahnya
- **"Supervisor …"** → berjajar sesuai urutan Service, Aksesoris, Pengadaan, Penyewaan, Maintenance, ISP
- **"Sales Corporate"** → di bawah Supervisor Pengadaan, Penyewaan, Maintenance, dan ISP
- jabatan lain (Admin, Sales, Teknisi, Kasir, …) → di bawah Supervisor Service (maksimal 8 kotak)

Baris yang namanya dikosongkan tetap ditampilkan sebagai kotak jabatan kosong; hapus barisnya
bila posisi itu memang tidak ada.

**Template Excel.** Di tab *Slide manual* → *Struktur organisasi* ada tombol
**Unduh template Excel**. File itu sudah berisi 12 baris jabatan baku, daftar pilihan jabatan,
dan sheet PETUNJUK. Isi kolom NAMA LENGKAP, simpan, lalu unggah kembali di kolom sebelahnya —
tabel di aplikasi langsung terisi. Lewat baris perintah gunakan
`--struktur TEMPLATE_STRUKTUR_ORGANISASI.xlsx`.

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
  "struktur": [
    {"nama": "Nama Ustadz", "jabatan": "Ustadz Pembina Cabang"},
    {"nama": "Nama Store Leader", "jabatan": "Store Leader"},
    {"nama": "Nama Supervisor", "jabatan": "Supervisor"},
    {"nama": "Nama Sales Corporate", "jabatan": "Sales Corporate"},
    {"nama": "Nama Teknisi", "jabatan": "Teknisi"}
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
