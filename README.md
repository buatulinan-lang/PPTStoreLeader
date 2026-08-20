# M-Flash Dashboard Builder

Aplikasi untuk mengubah file mentah Accurate menjadi dashboard interaktif **dan** presentasi
PowerPoint 23 slide yang mengikuti template standar weekly meeting M-Flash (latar, logo, warna,
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

## 2. Cara pakai

1. Unggah file mentah di panel kiri (boleh dua-duanya sekaligus):
   - `rincian_pengiriman_pesanan_*.xlsx` → **unit masuk** per hari/bulan, klasifikasi Done, Pending, Cancel
   - `rincian_faktur_penjualan_*.xlsx` → **penjualan** per hari/bulan, breakdown kategori, voucher, bagi hasil
2. Atur filter: tahun, bulan, kategori penjualan, dan dimensi pembanding (cabang/teknisi/admin).
3. Telusuri dashboard di tab yang tersedia.
4. Isi tab **Slide manual**: 4 goal, catatan pekan, tabel struktur organisasi, foto Measure Activity,
   foto AR, tabel komitmen, dan kesimpulan.
5. Buka tab **⬇️ Unduh PPT** → klik *Buat PPTX* → *Unduh PPTX*.

Aplikasi membaca kolom berdasarkan nama, jadi ekspor Accurate bulan depan bisa langsung dipakai
tanpa mengubah apa pun.

---

## 3. Isi deck yang dihasilkan (23 slide)

| # | Slide | Sumber |
|---|-------|--------|
| 1 | Cover | otomatis + judul/penyaji |
| 2 | Pencapaian Goal — 4 gauge: Gross Profit, Omset Aksesoris, Tingkat Kepuasan Pelanggan, Google Ulasan | input manual |
| 3 | Catatan pekan ini | input manual |
| 4 | Ringkasan kinerja | pengiriman pesanan |
| 5 | Komposisi status pengerjaan | pengiriman pesanan |
| 6 | Bulan berjalan vs bulan sebelumnya (dibandingkan setara) | pengiriman + faktur |
| 7 | Rekap unit masuk harian | pengiriman pesanan |
| 8 | Hari dengan unit masuk tertinggi | pengiriman pesanan |
| 9 | Kinerja per cabang/teknisi | pengiriman pesanan |
| 10–12 | Rincian Pending, Done, Cancel | pengiriman pesanan |
| 13 | Penjualan — modal, omzet & laba | faktur penjualan |
| 14 | Rekap penjualan harian & bulanan | faktur penjualan |
| 15 | Penjualan per kategori (service HP, service laptop, penjualan HP, dll) | faktur penjualan |
| 16 | Voucher | faktur penjualan |
| 17 | Bagi hasil teknisi | faktur penjualan |
| 18 | **Struktur organisasi** | tabel nama & jabatan |
| 19 | **Measure Activity** | unggah foto |
| 20 | **AR** | unggah foto |
| 21 | Komitmen — pencapaian, komitmen, target | input manual |
| 22 | Kesimpulan & tindak lanjut | otomatis, bisa diedit |
| 23 | Penutup | otomatis |

### Struktur organisasi

Cukup isi nama lengkap dan jabatan; bagan disusun otomatis:

- **Ustadz Pembina Cabang** (jabatan mengandung kata "Ustadz" atau "Pembina") → paling atas
- **Store Leader** (atau Store Manager / Kepala Toko / Pimpinan) → di bawah ustadz pembina
- **Supervisor** dan **Sales Corporate** → sejajar tepat di bawahnya
- jabatan lain → di bawah Supervisor

Kalau tidak ada baris berjabatan Store Leader, baris pertama otomatis dijadikan puncak.

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
  "catatan": ["Poin pertama", "Poin kedua", "Poin ketiga"],
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
