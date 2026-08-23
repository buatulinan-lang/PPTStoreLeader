"""Pembacaan & pembersihan file mentah M-Flash (Excel export Accurate)."""
from __future__ import annotations
import gzip
import io
import os
import re
import zipfile

import pandas as pd


def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().upper()


EKSTENSI_TABEL = (".csv", ".tsv", ".txt", ".gz", ".zip", ".xlsx", ".xls", ".xlsm", ".parquet")


def _baca_teks(buf, nama=""):
    """Baca CSV/TSV apa pun: pemisah dan encoding dideteksi otomatis."""
    if hasattr(buf, "read"):
        data = buf.read()
    else:
        with open(buf, "rb") as fh:
            data = fh.read()
    kesalahan = None
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        for sep in (None, ",", ";", "\t", "|"):
            try:
                df = pd.read_csv(io.BytesIO(data), sep=sep, encoding=enc,
                                 engine="python" if sep is None else "c",
                                 on_bad_lines="skip", low_memory=False)
                if df.shape[1] > 1:
                    return df
                kesalahan = kesalahan or ValueError("hanya satu kolom terbaca")
            except Exception as e:  # noqa: BLE001
                kesalahan = e
    raise ValueError(f"Tidak bisa membaca {nama or 'file'}: {kesalahan}")


def read_excel_any(file) -> pd.DataFrame:
    """Baca xlsx/xls/csv/tsv/csv.gz/zip/parquet.

    Kolom kosong 'Unnamed' yang jadi pemisah pada ekspor Accurate ikut dibuang.
    """
    nama = str(getattr(file, "name", file))
    rendah = nama.lower()

    if rendah.endswith(".gz"):                       # csv.gz — file gabungan terkompresi
        isi = file.read() if hasattr(file, "read") else open(file, "rb").read()
        df = _baca_teks(io.BytesIO(gzip.decompress(isi)), nama)
    elif rendah.endswith(".zip"):
        isi = file.read() if hasattr(file, "read") else open(file, "rb").read()
        with zipfile.ZipFile(io.BytesIO(isi)) as z:
            anggota = [m for m in z.namelist()
                       if m.lower().endswith((".csv", ".tsv", ".txt", ".xlsx", ".xls"))
                       and not m.startswith("__MACOSX")]
            if not anggota:
                raise ValueError(f"{os.path.basename(nama)}: tidak ada file tabel di dalam zip")
            bagian = []
            for m in anggota:
                with z.open(m) as fh:
                    data = fh.read()
                if m.lower().endswith((".xlsx", ".xls")):
                    bagian.append(pd.read_excel(io.BytesIO(data)))
                else:
                    bagian.append(_baca_teks(io.BytesIO(data), m))
            df = pd.concat(bagian, ignore_index=True) if len(bagian) > 1 else bagian[0]
    elif rendah.endswith(".parquet"):
        df = pd.read_parquet(file)
    elif rendah.endswith((".csv", ".tsv", ".txt")):
        df = _baca_teks(file, nama)
    else:
        df = pd.read_excel(file)

    df = df.loc[:, [c for c in df.columns if not str(c).startswith("Unnamed")]]
    df.columns = [_norm(c) for c in df.columns]
    df = df.dropna(axis=1, how="all")
    return df


def gabung(bagian) -> pd.DataFrame:
    """Satukan beberapa berkas sejenis menjadi satu tabel."""
    bagian = [b for b in bagian if b is not None and len(b)]
    if not bagian:
        return pd.DataFrame()
    if len(bagian) == 1:
        return bagian[0]
    return pd.concat(bagian, ignore_index=True, sort=False)


def detect_kind(df: pd.DataFrame) -> str:
    cols = set(df.columns)
    if "STATUS PENGERJAAN" in cols:
        return "pengiriman"
    if "NO FAKTUR" in cols or "TOTAL HARGA" in cols:
        return "faktur"
    return "unknown"


STATUS_ORDER = ["Done", "Cancel", "Pending", "Lainnya"]


def classify_status(v) -> str:
    s = _norm(v) if pd.notna(v) else ""
    if not s or s == "NAN":
        return "Lainnya"
    if "CANCEL" in s or "BATAL" in s:
        return "Cancel"
    if "PENDING" in s or "TERTAHAN" in s:
        return "Pending"
    if "DONE" in s or "SELESAI" in s:
        return "Done"
    return "Lainnya"


def prep_pengiriman(df: pd.DataFrame) -> pd.DataFrame:
    """Unit masuk. Baris identik (satu nota tercetak berulang) dihitung satu."""
    n_raw = len(df)
    df = df.drop_duplicates().copy()
    df["TGL PENGIRIMAN"] = pd.to_datetime(df.get("TGL PENGIRIMAN"), errors="coerce")
    if "TANGGAL STATUS PENGERJAAN" in df:
        df["TANGGAL STATUS PENGERJAAN"] = pd.to_datetime(df["TANGGAL STATUS PENGERJAAN"], errors="coerce")
    df = df[df["TGL PENGIRIMAN"].notna()]
    df["STATUS"] = df["STATUS PENGERJAAN"].map(classify_status) if "STATUS PENGERJAAN" in df else "Lainnya"
    df["TANGGAL"] = df["TGL PENGIRIMAN"].dt.normalize()
    df["TAHUN"] = df["TGL PENGIRIMAN"].dt.year
    df["BULAN"] = df["TGL PENGIRIMAN"].dt.month
    df["PERIODE"] = df["TGL PENGIRIMAN"].dt.to_period("M").astype(str)
    df["HARI"] = df["TGL PENGIRIMAN"].dt.dayofweek
    for c in ["NAMA TEKNISI", "NAMA ADMIN", "KERUSAKAN UTAMA", "MERK UNIT",
              "KATEGORI PENJUALAN", "KATEGORI PENGERJAAN UNIT", "CABANG"]:
        if c in df:
            df[c] = df[c].fillna("TIDAK ADA DATA").map(_norm)
    df.attrs["n_raw"] = n_raw
    df.attrs["n_unik"] = len(df)
    return df


def prep_faktur(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["TGL FAKTUR"] = pd.to_datetime(df.get("TGL FAKTUR"), errors="coerce")
    df = df[df["TGL FAKTUR"].notna()]
    for c in ["HARGA BELI", "QTY", "@HARGA", "TOTAL HARGA"]:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["OMZET"] = df["TOTAL HARGA"] if "TOTAL HARGA" in df else 0
    df["MODAL"] = df["HARGA BELI"] if "HARGA BELI" in df else 0
    df["LABA"] = df["OMZET"] - df["MODAL"]
    df["TANGGAL"] = df["TGL FAKTUR"].dt.normalize()
    df["TAHUN"] = df["TGL FAKTUR"].dt.year
    df["BULAN"] = df["TGL FAKTUR"].dt.month
    df["PERIODE"] = df["TGL FAKTUR"].dt.to_period("M").astype(str)
    for c in ["KATEGORI BARANG", "KATEGORI PENJUALAN", "NAMA BARANG", "NAMA TEKNISI (FINAL)",
              "NAMA TEKNISI", "KERUSAKAN UTAMA", "NAMA ADMIN", "CABANG"]:
        if c in df:
            df[c] = df[c].fillna("TIDAK ADA DATA").map(_norm)
    if "NAMA TEKNISI (FINAL)" in df:
        df["TEKNISI"] = df["NAMA TEKNISI (FINAL)"]
    elif "NAMA TEKNISI" in df:
        df["TEKNISI"] = df["NAMA TEKNISI"]
    else:
        df["TEKNISI"] = "TIDAK ADA DATA"
    return df


def group_options(dfp: pd.DataFrame):
    prefer = ["CABANG", "NAMA TEKNISI", "NAMA ADMIN", "KATEGORI PENJUALAN",
              "KATEGORI PENGERJAAN UNIT", "MERK UNIT", "KERUSAKAN UTAMA"]
    return [c for c in prefer if c in dfp.columns]
