# -*- coding: utf-8 -*-
"""Buat PPT tanpa membuka aplikasi (mode baris perintah).

Contoh:
    python buat_ppt.py "rincian_pengiriman_pesanan.xlsx" "rincian_faktur_penjualan.xlsx" \
        --tahun 2026 --lingkup "Cabang Klender" --out WEEKLY.pptx --isi isi.json
"""
import argparse, datetime as dt, json, sys
from mflash import loader, context as CTX, deck, metrics as M


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+", help="file pengiriman pesanan &/atau faktur penjualan")
    ap.add_argument("--tahun", type=int, nargs="*", default=None)
    ap.add_argument("--bulan", nargs="*", default=None, help="format 2026-08")
    ap.add_argument("--dim", default=None, help="dimensi pembanding, mis. 'NAMA TEKNISI'")
    ap.add_argument("--lingkup", default="Seluruh Data")
    ap.add_argument("--judul", default="WEEKLY MEETING")
    ap.add_argument("--penyaji", default="")
    ap.add_argument("--isi", default=None, help="file JSON berisi goal/komitmen/foto")
    ap.add_argument("--struktur", default=None, help="template Excel struktur organisasi")
    ap.add_argument("--kolom-pilar", dest="kolom_pilar", default=None,
                    help="nama kolom kategori pilar pada file faktur")
    ap.add_argument("--out", default=f"WEEKLY_MEETING_MFLASH_{dt.date.today():%Y%m%d}.pptx")
    a = ap.parse_args()

    dfp = dff = None
    n_raw = 0
    for f in a.files:
        raw = loader.read_excel_any(f)
        kind = loader.detect_kind(raw)
        if kind == "pengiriman":
            dfp = loader.prep_pengiriman(raw)
            n_raw = dfp.attrs.get("n_raw", len(raw))
            print(f"  pengiriman : {len(raw):,} baris → {len(dfp):,} unit unik")
        elif kind == "faktur":
            dff = loader.prep_faktur(raw)
            print(f"  faktur     : {len(dff):,} baris penjualan")
        else:
            print(f"  ! {f}: format tidak dikenali, dilewati")
    if dfp is None:
        sys.exit("File rincian pengiriman pesanan wajib ada.")

    tahun = a.tahun or [int(dfp['TAHUN'].max())]
    bulan = a.bulan or sorted(dfp[dfp["TAHUN"].isin(tahun)]["PERIODE"].unique())
    dim = a.dim or (loader.group_options(dfp) or [None])[0]

    manual = dict(judul=a.judul, penyaji=a.penyaji, tarif=M.DEFAULT_TARIF, flat=30.0,
                  voucher_kata="VOUCHER",
                  goals=[{"nama": g, "nilai": 0.0, "ket": ""} for g in
                         ("GROSS PROFIT", "OMSET AKSESORIS", "TINGKAT KEPUASAN PELANGGAN",
                          "GOOGLE ULASAN")],
                  catatan=[], struktur=[], komitmen=[], foto_measure=[], foto_ar=[],
                  kolom_pilar=a.kolom_pilar)
    if a.isi:
        manual.update(json.load(open(a.isi, encoding="utf-8")))
    # foto boleh ditulis sebagai path file di isi.json
    for k in ("foto_measure", "foto_ar"):
        manual[k] = [open(f, "rb").read() if isinstance(f, str) else f
                     for f in (manual.get(k) or [])]

    if a.struktur:
        from mflash import template as TPL
        manual["struktur"] = TPL.baca_struktur(a.struktur)
        print(f"  struktur   : {len(manual['struktur'])} baris jabatan")

    flt = dict(tahun=tahun, periode=bulan, dim=dim, lingkup=a.lingkup)
    c = CTX.build(dfp, dff, flt, manual, {"pengiriman_raw": n_raw})
    data = deck.build(c)
    with open(a.out, "wb") as fh:
        fh.write(data)
    from pptx import Presentation
    jml = len(Presentation(a.out).slides._sldIdLst)
    print(f"✅ {a.out} — {jml} slide (v{deck.VERSI}), periode {c['periode_label']}")


if __name__ == "__main__":
    main()
