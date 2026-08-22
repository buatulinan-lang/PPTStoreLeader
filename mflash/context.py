"""Merangkai seluruh angka + input manual menjadi satu 'context' untuk dashboard & deck."""
from __future__ import annotations
import datetime as dt
import pandas as pd
from . import metrics as M
from .metrics import n, pct, rp, tgl, periode_label


def _bersih(v):
    """Buang NaN/None dari input tabel manual."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v).strip()


def _bersih_rows(rows, kunci):
    out = []
    for r in rows or []:
        if isinstance(r, dict):
            d = {k: _bersih(r.get(k)) for k in kunci}
        else:
            d = {k: _bersih(v) for k, v in zip(kunci, list(r) + [""] * len(kunci))}
        if any(d.values()):
            out.append(d)
    return out


def _bersih_kesimpulan(rows):
    out = []
    for r in rows or []:
        a, b = (r[0], r[1]) if not isinstance(r, dict) else (r.get("judul"), r.get("isi"))
        a, b = _bersih(a), _bersih(b)
        if a or b:
            out.append((a, b))
    return out


LEVEL0 = ("USTADZ", "PEMBINA")
LEVEL1 = ("STORE LEADER", "STORE MANAGER", "KEPALA TOKO", "PIMPINAN")
URUTAN_SPV = ["SERVICE", "AKSESORIS", "PENGADAAN", "PENYEWAAN", "MAINTENANCE", "ISP"]
SPV_KANAN = ("PENGADAAN", "PENYEWAAN", "MAINTENANCE", "ISP")


def _tipe_spv(jabatan):
    j = jabatan.upper()
    for t in URUTAN_SPV:
        if t in j:
            return t
    return "LAINNYA"


def susun_struktur(rows):
    """Susun bagan: ustadz pembina → store leader → para supervisor →
    tim service (admin/sales/teknisi) dan sales corporate."""
    orang = _bersih_rows(rows, ["nama", "jabatan"])
    pembina, pimpinan, spv, corporate, tim = [], [], [], [], []
    for o in orang:
        j = o["jabatan"].upper()
        if any(k in j for k in LEVEL0) and not pembina:
            pembina.append(o)
        elif any(k in j for k in LEVEL1) and not pimpinan:
            pimpinan.append(o)
        elif "SUPERVISOR" in j or j.startswith("SPV"):
            spv.append(dict(o, tipe=_tipe_spv(j)))
        elif "CORPORATE" in j or "KORPORAT" in j:
            corporate.append(o)
        else:
            tim.append(o)
    spv.sort(key=lambda m: URUTAN_SPV.index(m["tipe"]) if m["tipe"] in URUTAN_SPV else 99)
    if not pimpinan and spv:
        pass
    return dict(pembina=pembina, pimpinan=pimpinan, supervisor=spv,
                corporate=corporate, tim=tim)


def periode_teks(p, flt):
    if not len(p):
        return "Tidak ada data"
    a, b = p["TANGGAL"].min(), p["TANGGAL"].max()
    lingkup = flt.get("lingkup") or "Seluruh Data"
    return f"{tgl(a)} – {tgl(b)} · {lingkup}"


def build(dfp, dff, flt, manual, raw_counts=None):
    p, f = M.apply_filter(dfp, dff, flt)
    dim = flt.get("dim") or ("NAMA TEKNISI" if "NAMA TEKNISI" in p else None)
    dim_label = (dim or "UNIT KERJA").replace("NAMA ", "")

    r = M.ringkasan(p)
    per_dim_df = M.per_dim(p, dim) if dim else pd.DataFrame()
    detail = {s: M.detail_status(p, s) for s in ("Done", "Cancel", "Pending")}

    dim_top = {}
    for s in ("Done", "Cancel", "Pending"):
        d = p[p["STATUS"] == s]
        if dim and len(d):
            vc = d[dim].value_counts()
            dim_top[s] = (vc.index[0], int(vc.iloc[0]), vc.iloc[0] / len(d) * 100)
        else:
            dim_top[s] = ("-", 0, 0.0)

    jual = M.penjualan(f)
    vch = M.voucher(f, manual.get("voucher_kata", "VOUCHER"))
    bagi = M.bagi_hasil(f, manual.get("tarif"), manual.get("flat", 30.0))
    mm = M.mom(p, f)

    nraw = (raw_counts or {}).get("pengiriman_raw", dfp.attrs.get("n_raw", len(dfp)))
    sumber = (f"Sumber: {n(nraw)} baris mentah → {n(len(dfp))} unit unik (baris identik dihitung satu). "
              f"Dibuat otomatis dari dashboard {dt.date.today().strftime('%d/%m/%Y')}.")

    bulan_top = M.per_bulan(p)
    catatan_ringkas = [
        f"Tingkat penyelesaian {pct(r['p_done'])} dari {n(r['total'])} unit masuk.",
        f"Pembatalan {pct(r['p_cancel'])} — {n(r['cancel'])} unit tidak jadi dikerjakan.",
        f"{n(r['pending'])} unit masih tertahan; terbanyak di {dim_top['Pending'][0]} "
        f"({n(dim_top['Pending'][1])} unit).",
        (f"Bulan tersibuk: {periode_label(bulan_top.idxmax())} dengan {n(bulan_top.max())} unit."
         if len(bulan_top) else "-"),
    ]

    kes = _bersih_kesimpulan(manual.get("kesimpulan")) or M.kesimpulan(
        r, detail["Pending"], detail["Cancel"], dim_label, per_dim_df)
    c = dict(
        p=p, f=f, r=r, dim=dim, dim_label=dim_label, per_dim=per_dim_df, dim_top=dim_top,
        detail=detail, jual=jual, voucher=vch, bagi=bagi, mom=mm,
        per_bulan=bulan_top, per_bulan_status=M.per_bulan_status(p),
        harian=M.harian(p), harian_status=M.harian_status(p),
        weekday=M.per_weekday(p), top_hari=M.top_hari(p, 10),
        periode_label=periode_teks(p, flt), sumber=sumber,
        catatan_ringkas=catatan_ringkas, kesimpulan=kes,
        voucher_kata=manual.get("voucher_kata", "VOUCHER"),
        voucher_judul=(manual.get("voucher_kata", "VOUCHER").upper()
                       if "VOUCHER" in manual.get("voucher_kata", "VOUCHER").upper()
                       else f"VOUCHER {manual.get('voucher_kata','').upper()}"),
        judul=manual.get("judul", "WEEKLY MEETING"),
        penyaji=manual.get("penyaji", ""),
        lingkup=(flt.get("lingkup") or "SELURUH DATA").upper(),
        dibuat=dt.date.today().strftime("%d %B %Y"),
        goal_judul=manual.get("goal_judul", "PENCAPAIAN GOAL"),
        goal_sub=manual.get("goal_sub", "Pencapaian tertinggi dibatasi 100%"),
        goals=manual.get("goals", []),
        catatan_judul=manual.get("catatan_judul", "CATATAN PEKAN INI"),
        catatan=[_bersih(x) for x in manual.get("catatan", []) if _bersih(x)],
        todo=_bersih_rows(manual.get("todo"), ["goal", "measure", "todo", "pic"]),
        komitmen=_bersih_rows(manual.get("komitmen"), ["pencapaian", "komitmen", "target"]),
        struktur=susun_struktur(manual.get("struktur")),
        foto_measure=[f for f in (manual.get("foto_measure") or []) if f],
        foto_ar=[f for f in (manual.get("foto_ar") or []) if f],
        foto_improvement=[f for f in (manual.get("foto_improvement") or []) if f],
        foto_todo=[f for f in (manual.get("foto_todo") or []) if f],
        teks_improvement=_bersih(manual.get("teks_improvement")),
        support=_bersih_rows(manual.get("support"), ["divisi", "needs"]),
        komitmen_intro=manual.get("komitmen_intro", "Pekan depan insyaAllah akan mencapai target sebesar:"),
    )
    return c
