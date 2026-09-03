"""Semua perhitungan angka untuk dashboard & deck."""
from __future__ import annotations
import calendar
import pandas as pd
import numpy as np
from .loader import STATUS_ORDER

BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]
BULAN_S = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]


# ------------------------------------------------------------------ format
def n(v, dec=0):
    if v is None or (isinstance(v, float) and (np.isnan(v))):
        return "-"
    s = f"{v:,.{dec}f}"
    return s.replace(",", "@").replace(".", ",").replace("@", ".")


def pct(v, dec=1):
    return f"{n(v, dec)}%"


def rp(v):
    v = float(v or 0)
    a = abs(v)
    sign = "-" if v < 0 else ""
    if a >= 1e12:
        return f"{sign}Rp {n(a/1e12, 2)} T"
    if a >= 1e9:
        return f"{sign}Rp {n(a/1e9, 2)} M"
    if a >= 1e6:
        return f"{sign}Rp {n(a/1e6, 1)} jt"
    if a >= 1e3:
        return f"{sign}Rp {n(a/1e3, 1)} rb"
    return f"{sign}Rp {n(a)}"


def tgl(d):
    d = pd.Timestamp(d)
    return f"{d.day} {BULAN[d.month-1]} {d.year}"


def tgl_s(d):
    d = pd.Timestamp(d)
    return f"{d.day} {BULAN_S[d.month-1]}"


def periode_label(p):
    y, m = str(p).split("-")
    return f"{BULAN[int(m)-1]} {y}"


def safe_div(a, b):
    return (a / b * 100) if b else 0.0


# ------------------------------------------------------------------ filter
def apply_filter(dfp, dff, flt):
    p, f = dfp, dff
    if flt.get("tahun"):
        p = p[p["TAHUN"].isin(flt["tahun"])]
        if f is not None and len(f):
            f = f[f["TAHUN"].isin(flt["tahun"])]
    if flt.get("periode"):
        p = p[p["PERIODE"].isin(flt["periode"])]
        if f is not None and len(f):
            f = f[f["PERIODE"].isin(flt["periode"])]
    for col in ("KATEGORI PENJUALAN",):
        val = flt.get(col)
        if val:
            if col in p:
                p = p[p[col].isin(val)]
            if f is not None and col in f:
                f = f[f[col].isin(val)]
    dim, vals = flt.get("dim"), flt.get("dim_vals")
    if dim and vals and dim in p:
        p = p[p[dim].isin(vals)]
    return p, f


# ------------------------------------------------------------------ inti
def ringkasan(p):
    total = len(p)
    c = p["STATUS"].value_counts()
    done, cancel, pending = int(c.get("Done", 0)), int(c.get("Cancel", 0)), int(c.get("Pending", 0))
    lain = total - done - cancel - pending
    hari = p["TANGGAL"].nunique()
    span = (p["TANGGAL"].max() - p["TANGGAL"].min()).days + 1 if total else 0
    return dict(total=total, done=done, cancel=cancel, pending=pending, lain=lain,
                p_done=safe_div(done, total), p_cancel=safe_div(cancel, total),
                p_pending=safe_div(pending, total), p_lain=safe_div(lain, total),
                hari_aktif=hari, hari_periode=span,
                rata_hari=(total / span if span else 0),
                rata_hari_aktif=(total / hari if hari else 0),
                mulai=p["TANGGAL"].min() if total else None,
                akhir=p["TANGGAL"].max() if total else None)


def per_bulan(p):
    g = p.groupby("PERIODE").size().sort_index()
    return g


def per_bulan_status(p):
    return p.pivot_table(index="PERIODE", columns="STATUS", values="TANGGAL",
                         aggfunc="count", fill_value=0).reindex(columns=STATUS_ORDER, fill_value=0).sort_index()


def harian(p):
    g = p.groupby("TANGGAL").size().sort_index()
    return g


def harian_status(p):
    return p.pivot_table(index="TANGGAL", columns="STATUS", values="TAHUN",
                         aggfunc="count", fill_value=0).reindex(columns=STATUS_ORDER, fill_value=0).sort_index()


def per_weekday(p):
    d = p.groupby(["HARI", "TANGGAL"]).size().reset_index(name="c")
    r = d.groupby("HARI")["c"].mean().reindex(range(7)).fillna(0)
    return r


def top_hari(p, k=10):
    g = harian(p)
    if not len(g):
        return pd.DataFrame(columns=["TANGGAL", "JUMLAH", "HARI", "VS"])
    avg = g.mean()
    t = g.sort_values(ascending=False).head(k).reset_index()
    t.columns = ["TANGGAL", "JUMLAH"]
    t["HARI"] = t["TANGGAL"].dt.dayofweek.map(lambda i: HARI[i])
    t["VS"] = (t["JUMLAH"] / avg - 1) * 100
    return t


def per_dim(p, dim, k=None):
    if dim not in p:
        return pd.DataFrame()
    g = p.groupby(dim).agg(UNIT=("STATUS", "size"))
    st = p.pivot_table(index=dim, columns="STATUS", values="TAHUN", aggfunc="count", fill_value=0)
    for s in STATUS_ORDER:
        g[s] = st[s] if s in st else 0
    g["P_DONE"] = g["Done"] / g["UNIT"] * 100
    g["P_CANCEL"] = g["Cancel"] / g["UNIT"] * 100
    g["P_PENDING"] = g["Pending"] / g["UNIT"] * 100
    g = g.sort_values("UNIT", ascending=False)
    return g.head(k) if k else g


def detail_status(p, status, dim_teknisi="NAMA TEKNISI"):
    d = p[p["STATUS"] == status]
    total = len(p)
    out = dict(jumlah=len(d), persen=safe_div(len(d), total))
    for key, col in (("teknisi", dim_teknisi), ("kerusakan", "KERUSAKAN UTAMA"),
                     ("kategori", "KATEGORI PENJUALAN"), ("merk", "MERK UNIT")):
        if col in d and len(d):
            vc = d[col].value_counts()
            out[key] = vc
            out[key + "_top"] = (vc.index[0], int(vc.iloc[0]), safe_div(int(vc.iloc[0]), len(d)))
        else:
            out[key] = pd.Series(dtype=int)
            out[key + "_top"] = ("-", 0, 0.0)
    span = (p["TANGGAL"].max() - p["TANGGAL"].min()).days + 1 if total else 0
    out["rata_hari"] = len(d) / span if span else 0
    out["hari_periode"] = span
    out["harian"] = d.groupby("TANGGAL").size().sort_index() if len(d) else pd.Series(dtype=int)
    out["bulanan"] = d.groupby("PERIODE").size().sort_index() if len(d) else pd.Series(dtype=int)
    return out


# ------------------------------------------------------------------ bulan berjalan
def mom(p, f=None):
    if not len(p):
        return None
    last = p["TANGGAL"].max()
    cur = last.to_period("M")
    prev = cur - 1
    cutoff = last.day - 1 if last.day > 1 else last.day
    cur_d = p[(p["TGL PENGIRIMAN"].dt.to_period("M") == cur) & (p["TGL PENGIRIMAN"].dt.day <= cutoff)]
    prev_d = p[(p["TGL PENGIRIMAN"].dt.to_period("M") == prev) & (p["TGL PENGIRIMAN"].dt.day <= cutoff)]

    def stat(d):
        c = d["STATUS"].value_counts()
        return dict(total=len(d), done=int(c.get("Done", 0)), cancel=int(c.get("Cancel", 0)),
                    rata=len(d) / cutoff if cutoff else 0)
    a, b = stat(prev_d), stat(cur_d)
    om_a = om_b = lb_a = lb_b = 0.0
    if f is not None and len(f):
        fa = f[(f["TGL FAKTUR"].dt.to_period("M") == prev) & (f["TGL FAKTUR"].dt.day <= cutoff)]
        fb = f[(f["TGL FAKTUR"].dt.to_period("M") == cur) & (f["TGL FAKTUR"].dt.day <= cutoff)]
        om_a, om_b = fa["OMZET"].sum(), fb["OMZET"].sum()
        lb_a, lb_b = fa["LABA"].sum(), fb["LABA"].sum()
    hari_bulan = calendar.monthrange(last.year, last.month)[1]
    proyeksi = b["total"] / cutoff * hari_bulan if cutoff else 0
    daily_cur = cur_d.groupby(cur_d["TGL PENGIRIMAN"].dt.day).size().reindex(range(1, cutoff + 1), fill_value=0)
    daily_prev = prev_d.groupby(prev_d["TGL PENGIRIMAN"].dt.day).size().reindex(range(1, cutoff + 1), fill_value=0)
    return dict(cur=str(cur), prev=str(prev), cur_label=periode_label(cur), prev_label=periode_label(prev),
                cutoff=cutoff, a=a, b=b, omzet_a=om_a, omzet_b=om_b, laba_a=lb_a, laba_b=lb_b,
                proyeksi=proyeksi, daily_cur=daily_cur, daily_prev=daily_prev,
                last=last, hari_bulan=hari_bulan)


def delta(cur, prev):
    if not prev:
        return None
    return (cur - prev) / prev * 100


# ------------------------------------------------------------------ pekanan
def pekanan(p, f=None):
    """Perkembangan per pekan (Senin–Minggu): jumlah transaksi & omzet."""
    if p is None or not len(p):
        return None
    d = p.copy()
    d["_AWAL"] = d["TANGGAL"] - pd.to_timedelta(d["TANGGAL"].dt.dayofweek, unit="D")
    trx = d.groupby("_AWAL").size().rename("TRANSAKSI")
    tabel = trx.to_frame()
    tabel["OMZET"] = 0.0
    tabel["LABA"] = 0.0
    if f is not None and len(f):
        g = f.copy()
        g["_AWAL"] = g["TANGGAL"] - pd.to_timedelta(g["TANGGAL"].dt.dayofweek, unit="D")
        uang = g.groupby("_AWAL")[["OMZET", "LABA"]].sum()
        tabel = tabel.join(uang, how="outer", rsuffix="_f")
        tabel["OMZET"] = tabel["OMZET_f"].fillna(0) if "OMZET_f" in tabel else tabel["OMZET"]
        tabel["LABA"] = tabel["LABA_f"].fillna(0) if "LABA_f" in tabel else tabel["LABA"]
        tabel = tabel.drop(columns=[c for c in tabel.columns if c.endswith("_f")])
        tabel["TRANSAKSI"] = tabel["TRANSAKSI"].fillna(0).astype(int)
    tabel = tabel.sort_index()
    if not len(tabel):
        return None
    tabel["AKHIR"] = tabel.index + pd.Timedelta(days=6)
    tabel["LABEL"] = [_label_pekan(a, b) for a, b in zip(tabel.index, tabel["AKHIR"])]
    tabel["D_TRX"] = tabel["TRANSAKSI"].pct_change() * 100
    tabel["D_OMZET"] = tabel["OMZET"].pct_change() * 100

    akhir = tabel.iloc[-1]
    sebelum = tabel.iloc[-2] if len(tabel) > 1 else None
    return dict(tabel=tabel, jumlah=len(tabel),
                total_trx=int(tabel["TRANSAKSI"].sum()), total_omzet=float(tabel["OMZET"].sum()),
                rata_trx=float(tabel["TRANSAKSI"].mean()), rata_omzet=float(tabel["OMZET"].mean()),
                puncak_trx=tabel["TRANSAKSI"].idxmax(), puncak_omzet=tabel["OMZET"].idxmax(),
                terakhir=akhir, sebelum=sebelum,
                mulai=tabel.index[0], selesai=tabel["AKHIR"].iloc[-1],
                ada_omzet=bool(tabel["OMZET"].sum()))


def _label_pekan(a, b):
    a, b = pd.Timestamp(a), pd.Timestamp(b)
    if a.month == b.month:
        return f"{a.day}–{b.day} {BULAN_S[a.month - 1]}"
    return f"{a.day} {BULAN_S[a.month - 1]}–{b.day} {BULAN_S[b.month - 1]}"


# ------------------------------------------------------------------ penjualan
def penjualan(f):
    if f is None or not len(f):
        return None
    omzet, modal = f["OMZET"].sum(), f["MODAL"].sum()
    laba = omzet - modal
    faktur = f["NO FAKTUR"].nunique() if "NO FAKTUR" in f else len(f)
    unit = f["QTY"].sum() if "QTY" in f else len(f)
    key = "KATEGORI BARANG" if "KATEGORI BARANG" in f else "KATEGORI PENJUALAN"
    def _grp(col):
        if col not in f:
            return pd.DataFrame()
        gg = f.groupby(col).agg(OMZET=("OMZET", "sum"), MODAL=("MODAL", "sum"), LABA=("LABA", "sum"),
                                QTY=("QTY", "sum") if "QTY" in f else ("OMZET", "size"),
                                FAKTUR=("NO FAKTUR", "nunique") if "NO FAKTUR" in f else ("OMZET", "size"))
        gg["MARGIN"] = gg["LABA"] / gg["OMZET"].replace(0, np.nan) * 100
        return gg.sort_values("OMZET", ascending=False)

    g = _grp(key)
    g_jual = _grp("KATEGORI PENJUALAN")
    harian_jual = f.groupby("TANGGAL")[["OMZET", "MODAL", "LABA"]].sum().sort_index()
    harian_jual["FAKTUR"] = f.groupby("TANGGAL")["NO FAKTUR"].nunique() if "NO FAKTUR" in f else f.groupby("TANGGAL").size()
    bulan_kat = (f.pivot_table(index="PERIODE", columns="KATEGORI PENJUALAN", values="OMZET",
                               aggfunc="sum", fill_value=0).sort_index()
                 if "KATEGORI PENJUALAN" in f else pd.DataFrame())
    return dict(omzet=omzet, modal=modal, laba=laba, margin=safe_div(laba, omzet),
                faktur=faktur, unit=unit, baris=len(f), per_kategori=g,
                per_kategori_jual=g_jual, harian=harian_jual, bulan_kategori=bulan_kat,
                per_bulan=f.groupby("PERIODE")[["OMZET", "MODAL", "LABA"]].sum().sort_index(),
                faktur_bulan=(f.groupby("PERIODE")["NO FAKTUR"].nunique().sort_index()
                              if "NO FAKTUR" in f else f.groupby("PERIODE").size().sort_index()),
                rata_faktur=(omzet / faktur if faktur else 0),
                laba_faktur=(laba / faktur if faktur else 0),
                laba_unit=(laba / unit if unit else 0),
                p_modal=safe_div(modal, omzet))


def voucher(f, kata="VOUCHER"):
    if f is None or not len(f) or "NAMA BARANG" not in f:
        return None
    d = f[f["NAMA BARANG"].str.contains(kata, na=False)]
    if not len(d):
        return None
    qty = d["QTY"].sum() if "QTY" in d else len(d)
    omzet, modal = d["OMZET"].sum(), d["MODAL"].sum()
    hari = d["TANGGAL"].nunique()
    dim = "CABANG" if "CABANG" in d else ("NAMA ADMIN" if "NAMA ADMIN" in d else "NAMA BARANG")
    return dict(qty=qty, trx=(d["NO FAKTUR"].nunique() if "NO FAKTUR" in d else len(d)),
                omzet=omzet, modal=modal, laba=omzet - modal,
                margin=safe_div(omzet - modal, omzet), p_modal=safe_div(modal, omzet),
                hari=hari, rata_hari=(qty / hari if hari else 0),
                per_dim=_vdim(d, dim),
                dim=dim,
                per_bulan=d.groupby("PERIODE")["QTY"].sum().sort_index() if "QTY" in d else d["PERIODE"].value_counts().sort_index(),
                mulai=d["TANGGAL"].min(), akhir=d["TANGGAL"].max(),
                rata_harga=(omzet / qty if qty else 0))


def _vdim(d, dim):
    s = d.groupby(dim)["QTY"].sum() if "QTY" in d else d[dim].value_counts()
    s = s[[i for i in s.index if str(i) != "TIDAK ADA DATA"]] if len(s) > 1 else s
    return s.sort_values(ascending=False)


KOSONG = {"", "NAN", "NONE", "-", "0", "TIDAK ADA DATA", "TIDAK ADA", "NULL"}


def kolom_pilar(f, kolom=None):
    """Cari kolom kategori pilar; kembalikan None bila tidak ada."""
    if f is None or not len(f):
        return None
    if kolom and kolom in f.columns:
        return kolom
    for c in f.columns:
        cc = str(c).upper()
        if "PILAR" in cc or "PILLAR" in cc:
            return c
    return None


def pilar(f, kolom=None):
    """Omzet, modal, laba per kategori pilar.

    Baris tanpa kategori pilar tidak ikut dihitung — hanya baris yang punya
    kategori pilar yang masuk pencapaian.
    """
    kol = kolom_pilar(f, kolom)
    if kol is None:
        return None
    d = f.copy()
    d["_PILAR"] = d[kol].astype(str).str.strip().str.upper()
    d = d[~d["_PILAR"].isin(KOSONG) & d[kol].notna()]
    if not len(d):
        return None
    g = d.groupby("_PILAR").agg(OMZET=("OMZET", "sum"), MODAL=("MODAL", "sum"),
                                LABA=("LABA", "sum"))
    g["FAKTUR"] = (d.groupby("_PILAR")["NO FAKTUR"].nunique() if "NO FAKTUR" in d
                   else d.groupby("_PILAR").size())
    g["MARGIN"] = g["LABA"] / g["OMZET"].replace(0, np.nan) * 100
    g["KONTRIBUSI"] = g["OMZET"] / g["OMZET"].sum() * 100
    g = g.sort_values("OMZET", ascending=False)
    omzet, modal = g["OMZET"].sum(), g["MODAL"].sum()
    laba = omzet - modal
    omzet_semua = f["OMZET"].sum()
    return dict(kolom=str(kol), per_pilar=g, omzet=omzet, modal=modal, laba=laba,
                margin=safe_div(laba, omzet), jumlah=len(g),
                baris=len(d), baris_total=len(f),
                cakupan=safe_div(omzet, omzet_semua),
                per_bulan=d.groupby(["PERIODE", "_PILAR"])["OMZET"].sum().unstack(fill_value=0).sort_index()
                if "PERIODE" in d else None)


DEFAULT_TARIF = {"INTERFACE": 20.0, "NORMAL": 30.0, "MATI TOTAL": 32.0, "PROMO": 60.0, "LAINNYA": 30.0}


def bagi_hasil(f, tarif=None, flat=30.0):
    """Bagi hasil teknisi dari omzet JASA."""
    if f is None or not len(f):
        return None
    tarif = {**DEFAULT_TARIF, **(tarif or {})}
    key = "KATEGORI BARANG" if "KATEGORI BARANG" in f else None
    d = f[f[key] == "JASA"] if key else f
    if not len(d):
        return None

    def kat(row):
        s = str(row.get("KERUSAKAN UTAMA", "")) + " " + str(row.get("NAMA BARANG", ""))
        s = s.upper()
        if "NORMAL" in s:
            return "NORMAL"
        for k in ("PROMO", "INTERFACE", "MATI TOTAL"):
            if k in s:
                return k
        return "LAINNYA"
    d = d.copy()
    d["KAT_TARIF"] = d.apply(kat, axis=1)
    d["TARIF"] = d["KAT_TARIF"].map(tarif).fillna(tarif["LAINNYA"])
    d["BAGI"] = d["OMZET"] * d["TARIF"] / 100
    per_kat = d.groupby("KAT_TARIF").agg(TARIF=("TARIF", "first"), OMZET=("OMZET", "sum"),
                                         BAGI=("BAGI", "sum")).sort_values("OMZET", ascending=False)
    per_tek = d.groupby("TEKNISI").agg(OMZET=("OMZET", "sum"), BAGI=("BAGI", "sum")).sort_values("BAGI", ascending=False)
    omzet = d["OMZET"].sum()
    total = d["BAGI"].sum()
    return dict(omzet=omzet, total=total, persen=safe_div(total, omzet), baris=len(d),
                flat=omzet * flat / 100, flat_rate=flat, selisih=total - omzet * flat / 100,
                teknisi=len(per_tek), per_kat=per_kat, per_teknisi=per_tek,
                rata=(total / len(per_tek) if len(per_tek) else 0), tarif=tarif)


# ------------------------------------------------------------------ kesimpulan otomatis
def kesimpulan(r, dpen, dcan, dim_label, per_dim_df):
    out = []
    out.append(("Jaga tingkat penyelesaian",
                f"Saat ini {pct(r['p_done'])} unit tuntas ({n(r['done'])} unit). Pertahankan kapasitas "
                f"dan pastikan tidak turun saat volume naik."))
    kk = dcan.get("kerusakan_top", ("-", 0, 0))
    out.append(("Telusuri penyebab pembatalan",
                f"Pembatalan {pct(r['p_cancel'])} ({n(r['cancel'])} unit), paling sering pada kerusakan "
                f"{kk[0]} ({n(kk[1])} unit). Estimasi biaya lebih awal berpotensi menekan angka ini."))
    pt = dpen.get("teknisi_top", ("-", 0, 0))
    pk = dpen.get("kerusakan_top", ("-", 0, 0))
    out.append(("Selesaikan unit yang tertahan",
                f"{n(r['pending'])} unit pending. Prioritaskan {pt[0]} ({n(pt[1])} unit) dan kerusakan "
                f"{pk[0]} ({n(pk[1])} unit) sebelum berkembang jadi komplain."))
    if per_dim_df is not None and len(per_dim_df):
        d = per_dim_df[per_dim_df["UNIT"] >= max(20, per_dim_df["UNIT"].median())]
        d = d if len(d) else per_dim_df
        w = d["P_CANCEL"].idxmax()
        out.append((f"Tinjau {dim_label.lower()} dengan pembatalan tertinggi",
                    f"{w} mencatat pembatalan {pct(d.loc[w,'P_CANCEL'])}, di atas rata-rata {pct(r['p_cancel'])}. "
                    f"Perlu ditelusuri apakah soal harga, waktu tunggu, atau sparepart."))
    else:
        out.append(("Rapikan pencatatan data",
                    "Sebagian baris tidak memiliki dimensi pembanding sehingga analisis per unit kerja terbatas."))
    return out
