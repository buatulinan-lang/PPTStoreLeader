# -*- coding: utf-8 -*-
"""M-Flash Dashboard Builder — dashboard interaktif + ekspor PPT standar weekly meeting."""
import datetime as dt
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from mflash import loader, metrics as M, context as CTX, deck
from mflash.metrics import n, pct, rp, tgl, periode_label

st.set_page_config(page_title="M-Flash Dashboard Builder", page_icon="📊", layout="wide")

NAVY, GREEN, RED, AMBER, BLUE, MUTED = "#1F3864", "#16A34A", "#C0392B", "#E18C1F", "#2E9BD6", "#6B7280"
STATUS_COLOR = {"Done": GREEN, "Cancel": RED, "Pending": AMBER, "Lainnya": BLUE}

st.markdown(f"""
<style>
 .stApp {{ background:#FAFBFE; }}
 h1,h2,h3 {{ color:{NAVY}; font-family:Calibri,Segoe UI,sans-serif; }}
 div[data-testid="stMetric"] {{ background:#F6F8FC; border:1px solid #D5DDEB; border-radius:10px; padding:12px 14px; }}
 div[data-testid="stMetricLabel"] p {{ color:{MUTED}; font-size:.78rem; font-weight:700; letter-spacing:.04em; }}
</style>""", unsafe_allow_html=True)


def px_style(fig, h=330, legend=True):
    fig.update_layout(height=h, margin=dict(l=10, r=10, t=30, b=10),
                      font=dict(family="Calibri, Segoe UI", size=13, color="#20242E"),
                      plot_bgcolor="white", paper_bgcolor="white",
                      showlegend=legend, legend=dict(orientation="h", y=-0.18, title=""))
    fig.update_xaxes(showgrid=False, linecolor="#D5DDEB")
    fig.update_yaxes(gridcolor="#EDF1F8", zeroline=False)
    return fig


# ------------------------------------------------------------------ data
@st.cache_data(show_spinner="Membaca & merapikan file… (file besar butuh 20–40 detik)")
def _load(file_bytes, name):
    """Baca + siapkan sekali saja; hasilnya di-cache supaya filter terasa instan."""
    import io
    src = (io.StringIO(file_bytes.decode("utf-8", "ignore")) if name.lower().endswith(".csv")
           else io.BytesIO(file_bytes))
    src.name = name
    df = loader.read_excel_any(src)
    kind = loader.detect_kind(df)
    if kind == "pengiriman":
        out = loader.prep_pengiriman(df)
        return out, kind, {"raw": out.attrs.get("n_raw", len(df))}
    if kind == "faktur":
        return loader.prep_faktur(df), kind, {"raw": len(df)}
    return df, kind, {"raw": len(df)}


st.sidebar.title("📊 M-Flash Dashboard")
st.sidebar.caption("Unggah file mentah → dashboard → unduh PPT")

up = st.sidebar.file_uploader("File mentah (bisa 2 sekaligus)", type=["xlsx", "xls", "csv"],
                              accept_multiple_files=True)

if "store" not in st.session_state:
    st.session_state.store = {}
for uf in up or []:
    df, kind, meta = _load(uf.getvalue(), uf.name)
    if kind == "unknown":
        st.sidebar.warning(f"{uf.name}: format tidak dikenali, dilewati.")
    else:
        st.session_state.store[kind] = (df, uf.name, meta)

store = st.session_state.store
if "pengiriman" not in store:
    st.title("M-Flash Dashboard Builder")
    st.info("Unggah minimal file **rincian pengiriman pesanan** di panel kiri. "
            "File **rincian faktur penjualan** melengkapi slide omzet, laba, voucher, dan bagi hasil.")
    st.stop()

dfp, nama_p, meta_p = store["pengiriman"]
dff = store["faktur"][0] if "faktur" in store else None
n_raw = meta_p.get("raw", len(dfp))

st.sidebar.success(f"Pengiriman: {n(n_raw)} baris → {n(len(dfp))} unit unik")
if dff is not None:
    st.sidebar.success(f"Faktur: {n(len(dff))} baris penjualan")
else:
    st.sidebar.info("File faktur penjualan belum diunggah — slide omzet/laba/voucher akan kosong.")

# ------------------------------------------------------------------ filter
st.sidebar.header("Filter")
tahun_all = sorted(dfp["TAHUN"].unique())
tahun = st.sidebar.multiselect("Tahun", tahun_all, default=[tahun_all[-1]])
per_all = sorted(dfp[dfp["TAHUN"].isin(tahun)]["PERIODE"].unique()) if tahun else sorted(dfp["PERIODE"].unique())
periode = st.sidebar.multiselect("Bulan", per_all, default=per_all,
                                 format_func=periode_label)
kat_all = sorted(dfp["KATEGORI PENJUALAN"].dropna().unique()) if "KATEGORI PENJUALAN" in dfp else []
kat = st.sidebar.multiselect("Kategori penjualan", kat_all, default=[])
dims = loader.group_options(dfp)
dim = st.sidebar.selectbox("Dimensi pembanding (slide 'kinerja per …')", dims) if dims else None
dim_vals = st.sidebar.multiselect(f"Filter {dim}", sorted(dfp[dim].unique()), default=[]) if dim else []
lingkup = st.sidebar.text_input("Lingkup (tampil di cover)", "Cabang Klender")

flt = dict(tahun=tahun, periode=periode, **{"KATEGORI PENJUALAN": kat},
           dim=dim, dim_vals=dim_vals, lingkup=lingkup)

# ------------------------------------------------------------------ input manual
st.sidebar.header("Isi slide manual")
judul = st.sidebar.text_input("Judul cover", "WEEKLY MEETING")
penyaji = st.sidebar.text_input("Penyaji / jabatan", "")
voucher_kata = st.sidebar.text_input("Kata kunci voucher", "VOUCHER")
flat = st.sidebar.number_input("Pembanding bagi hasil flat (%)", 0.0, 100.0, 30.0, 1.0)

GOAL_DEFAULT = ["GROSS PROFIT", "OMSET AKSESORIS", "TINGKAT KEPUASAN PELANGGAN", "GOOGLE ULASAN"]
JABATAN_DEFAULT = ["Ustadz Pembina Cabang", "Store Leader",
                   "Supervisor Service", "Supervisor Aksesoris", "Supervisor Pengadaan",
                   "Supervisor Penyewaan", "Supervisor Maintenance", "Supervisor ISP",
                   "Admin", "Sales", "Teknisi", "Sales Corporate"]

man = st.session_state.setdefault("manual", {
    "goals": [{"nama": g, "nilai": 0.0, "ket": ""} for g in GOAL_DEFAULT],
    "struktur": [{"nama": "", "jabatan": j} for j in JABATAN_DEFAULT],
    "komitmen": [{"pencapaian": "", "komitmen": "", "target": ""} for _ in range(4)],
    "foto_measure": [],
    "foto_ar": [],
})
man.setdefault("foto_measure", [])
man.setdefault("foto_ar", [])
while len(man["goals"]) < 4:
    man["goals"].append({"nama": GOAL_DEFAULT[len(man["goals"])], "nilai": 0.0, "ket": ""})

manual = dict(judul=judul, penyaji=penyaji, voucher_kata=voucher_kata, flat=flat,
              tarif=M.DEFAULT_TARIF, **man)

c = CTX.build(dfp, dff, flt, manual, {"pengiriman_raw": n_raw})
r = c["r"]
p = c["p"]

# ------------------------------------------------------------------ header
st.title("M-Flash Dashboard Builder")
st.caption(c["periode_label"])

k = st.columns(5)
k[0].metric("TOTAL UNIT MASUK", n(r["total"]))
k[1].metric("SELESAI (DONE)", n(r["done"]), pct(r["p_done"]))
k[2].metric("BATAL (CANCEL)", n(r["cancel"]), pct(r["p_cancel"]))
k[3].metric("PENDING", n(r["pending"]), pct(r["p_pending"]))
k[4].metric("RATA-RATA / HARI", n(r["rata_hari"], 1))

tabs = st.tabs(["Ringkasan", "Harian & Bulanan", "Status", f"Per {c['dim_label'].title()}",
                "Penjualan", "Voucher & Bagi Hasil", "Slide manual", "⬇️ Unduh PPT"])

# --- ringkasan
with tabs[0]:
    a, b = st.columns([2, 1])
    pbs = c["per_bulan_status"]
    pbs.index = [periode_label(i) for i in pbs.index]
    fig = go.Figure()
    for s_ in ["Done", "Cancel", "Pending", "Lainnya"]:
        if pbs[s_].sum():
            fig.add_bar(x=pbs.index, y=pbs[s_], name=s_, marker_color=STATUS_COLOR[s_])
    fig.update_layout(barmode="stack")
    a.plotly_chart(px_style(fig, 380), use_container_width=True)
    b.subheader("Catatan utama")
    for x in c["catatan_ringkas"]:
        b.markdown(f"- {x}")
    st.dataframe(pbs, use_container_width=True)

# --- harian
with tabs[1]:
    h = c["harian"]
    fig = px.line(x=h.index, y=h.values, labels={"x": "Tanggal", "y": "Unit masuk"})
    fig.update_traces(line_color=NAVY)
    st.plotly_chart(px_style(fig, 340, False), use_container_width=True)
    a, b = st.columns(2)
    wd = c["weekday"]
    f2 = px.bar(x=[M.HARI[i] for i in range(7)], y=wd.values, labels={"x": "", "y": "Rata-rata"})
    f2.update_traces(marker_color=NAVY)
    a.subheader("Rata-rata per hari")
    a.plotly_chart(px_style(f2, 320, False), use_container_width=True)
    b.subheader("10 tanggal tertinggi")
    t = c["top_hari"].copy()
    t["TANGGAL"] = t["TANGGAL"].map(tgl)
    t["VS"] = t["VS"].map(lambda v: f"+{pct(v,0)}")
    b.dataframe(t.rename(columns={"VS": "VS RATA-RATA"}), use_container_width=True, hide_index=True)
    m = c["mom"]
    if m:
        st.subheader(f"{m['cur_label']} vs {m['prev_label']} (tanggal 1–{m['cutoff']})")
        cols = st.columns(6)
        for i, (l, cur, prev, fmt) in enumerate([
                ("Unit masuk", m["b"]["total"], m["a"]["total"], n),
                ("Done", m["b"]["done"], m["a"]["done"], n),
                ("Cancel", m["b"]["cancel"], m["a"]["cancel"], n),
                ("Rata/hari", m["b"]["rata"], m["a"]["rata"], lambda v: n(v, 1)),
                ("Omzet", m["omzet_b"], m["omzet_a"], rp),
                ("Laba", m["laba_b"], m["laba_a"], rp)]):
            d = M.delta(cur, prev)
            cols[i].metric(l, fmt(cur), None if d is None else f"{'+' if d>=0 else ''}{pct(d)}")
        f3 = go.Figure()
        f3.add_scatter(x=list(m["daily_prev"].index), y=m["daily_prev"].values, name=m["prev_label"],
                       line=dict(color=MUTED))
        f3.add_scatter(x=list(m["daily_cur"].index), y=m["daily_cur"].values, name=m["cur_label"],
                       line=dict(color=NAVY, width=3))
        st.plotly_chart(px_style(f3, 320), use_container_width=True)

# --- status
with tabs[2]:
    vals = [(k_, r[v], r[pv]) for k_, v, pv in [("Done", "done", "p_done"), ("Cancel", "cancel", "p_cancel"),
                                                ("Pending", "pending", "p_pending"), ("Lainnya", "lain", "p_lain")]]
    vals = [v for v in vals if v[1] > 0]
    a, b = st.columns([1, 1])
    fig = go.Figure(go.Pie(labels=[v[0] for v in vals], values=[v[1] for v in vals], hole=.58,
                           marker_colors=[STATUS_COLOR[v[0]] for v in vals]))
    a.plotly_chart(px_style(fig, 380), use_container_width=True)
    sel = b.selectbox("Lihat rincian status", ["Pending", "Done", "Cancel"])
    d = c["detail"][sel]
    b.metric(f"Jumlah {sel}", n(d["jumlah"]), pct(d["persen"]))
    for lab, key in [("Teknisi terbanyak", "teknisi_top"), ("Kerusakan terbanyak", "kerusakan_top")]:
        v = d[key]
        b.write(f"**{lab}:** {v[0]} — {n(v[1])} unit ({pct(v[2])})")
    x, y = st.columns(2)
    for col, key, title in [(x, "teknisi", "Teknisi"), (y, "kerusakan", "Kerusakan")]:
        s_ = d[key].head(10).iloc[::-1]
        f4 = px.bar(x=s_.values, y=[str(i)[:24] for i in s_.index], orientation="h")
        f4.update_traces(marker_color=STATUS_COLOR[sel])
        col.subheader(f"{title} — {sel}")
        col.plotly_chart(px_style(f4, 380, False), use_container_width=True)

# --- per dimensi
with tabs[3]:
    g = c["per_dim"]
    if len(g):
        st.dataframe(g.assign(**{"% DONE": g["P_DONE"].map(lambda v: pct(v)),
                                 "% CANCEL": g["P_CANCEL"].map(lambda v: pct(v)),
                                 "% PENDING": g["P_PENDING"].map(lambda v: pct(v))})
                     [["UNIT", "Done", "Cancel", "Pending", "% DONE", "% CANCEL", "% PENDING"]],
                     use_container_width=True)
        gg = g.head(12).iloc[::-1]
        fig = go.Figure()
        for s_ in ["Done", "Cancel", "Pending"]:
            fig.add_bar(y=[str(i)[:22] for i in gg.index], x=gg[s_], name=s_,
                        orientation="h", marker_color=STATUS_COLOR[s_])
        fig.update_layout(barmode="stack")
        st.plotly_chart(px_style(fig, 480), use_container_width=True)
    else:
        st.info("Tidak ada dimensi pembanding pada data ini.")

# --- penjualan
with tabs[4]:
    j = c["jual"]
    if not j:
        st.info("Unggah file **rincian faktur penjualan** untuk mengisi bagian ini.")
    else:
        cc = st.columns(5)
        cc[0].metric("OMZET", rp(j["omzet"]))
        cc[1].metric("MODAL", rp(j["modal"]), pct(j["p_modal"]))
        cc[2].metric("LABA KOTOR", rp(j["laba"]), f"margin {pct(j['margin'])}")
        cc[3].metric("FAKTUR", n(j["faktur"]))
        cc[4].metric("LABA / UNIT", rp(j["laba_unit"]))
        hj = j["harian"]
        fig = px.line(x=hj.index, y=hj["OMZET"] / 1e6, labels={"x": "Tanggal", "y": "Omzet (juta)"})
        fig.update_traces(line_color=NAVY)
        st.subheader("Omzet harian")
        st.plotly_chart(px_style(fig, 330, False), use_container_width=True)
        a, b = st.columns(2)
        pb = j["per_bulan"]
        f5 = go.Figure()
        f5.add_bar(x=[periode_label(i) for i in pb.index], y=pb["OMZET"] / 1e6, name="Omzet (jt)", marker_color=NAVY)
        f5.add_bar(x=[periode_label(i) for i in pb.index], y=pb["LABA"] / 1e6, name="Laba (jt)", marker_color=GREEN)
        a.subheader("Per bulan")
        a.plotly_chart(px_style(f5, 340), use_container_width=True)
        gk = j["per_kategori_jual"]
        b.subheader("Per kategori penjualan")
        if len(gk):
            b.dataframe(pd.DataFrame({"FAKTUR": gk["FAKTUR"].map(n), "OMZET": gk["OMZET"].map(rp),
                                      "LABA": gk["LABA"].map(rp),
                                      "MARGIN": gk["MARGIN"].map(lambda v: pct(v))}),
                        use_container_width=True)
        st.subheader("Per kategori barang")
        gb = j["per_kategori"]
        st.dataframe(pd.DataFrame({"OMZET": gb["OMZET"].map(rp), "MODAL": gb["MODAL"].map(rp),
                                   "LABA": gb["LABA"].map(rp), "MARGIN": gb["MARGIN"].map(lambda v: pct(v))}),
                     use_container_width=True)

# --- voucher & bagi hasil
with tabs[5]:
    v, bg = c["voucher"], c["bagi"]
    a, b = st.columns(2)
    a.subheader(f"Voucher ({voucher_kata})")
    if v:
        a.metric("Voucher terjual", n(v["qty"]), f"omzet {rp(v['omzet'])}")
        gv = v["per_dim"].head(10).iloc[::-1]
        fv = px.bar(x=gv.values, y=[str(i)[:22] for i in gv.index], orientation="h")
        fv.update_traces(marker_color=NAVY)
        a.plotly_chart(px_style(fv, 340, False), use_container_width=True)
    else:
        a.info("Tidak ada baris voucher pada filter/kata kunci ini.")
    b.subheader("Bagi hasil teknisi")
    if bg:
        b.metric("Bagi hasil (aturan)", rp(bg["total"]), f"{pct(bg['persen'])} dari omzet jasa")
        gt = bg["per_teknisi"].head(10).iloc[::-1]
        ft = px.bar(x=gt["BAGI"] / 1e6, y=[str(i)[:22] for i in gt.index], orientation="h")
        ft.update_traces(marker_color=NAVY)
        b.plotly_chart(px_style(ft, 340, False), use_container_width=True)
        b.dataframe(pd.DataFrame({"TARIF": bg["per_kat"]["TARIF"].map(lambda x_: pct(x_, 0)),
                                  "OMZET": bg["per_kat"]["OMZET"].map(rp),
                                  "BAGI HASIL": bg["per_kat"]["BAGI"].map(rp)}), use_container_width=True)
    else:
        b.info("Data jasa tidak tersedia.")

# --- slide manual
with tabs[6]:
    st.subheader("Slide 2 — Pencapaian Goal")
    st.caption("Isi pencapaian dalam persen. Di bawah 85% tampil merah, 85–100% kuning, di atas 100% hijau.")
    gcols = st.columns(4)
    for i in range(4):
        with gcols[i]:
            man["goals"][i]["nama"] = st.text_input(f"Goal {i+1}", man["goals"][i]["nama"], key=f"gn{i}")
            man["goals"][i]["nilai"] = st.number_input(f"Pencapaian (%) {i+1}", 0.0, 1000.0,
                                                       float(man["goals"][i]["nilai"]), 0.01, key=f"gv{i}")
            man["goals"][i]["ket"] = st.text_input(f"Keterangan {i+1}", man["goals"][i]["ket"], key=f"gk{i}")

    st.divider()
    st.subheader("Slide Struktur Organisasi")
    st.caption("Urutan otomatis: Ustadz Pembina Cabang → Store Leader → para Supervisor "
               "(Service, Aksesoris, Pengadaan, Penyewaan, Maintenance, ISP). Admin, Sales, dan "
               "Teknisi masuk di bawah Supervisor Service; Sales Corporate di bawah Supervisor "
               "Pengadaan, Penyewaan, Maintenance, dan ISP.")
    man["struktur"] = st.data_editor(
        pd.DataFrame(man["struktur"]), num_rows="dynamic", use_container_width=True, key="str_ed",
        column_config={"nama": st.column_config.TextColumn("Nama lengkap", width="large"),
                       "jabatan": st.column_config.TextColumn("Jabatan", width="medium")}
    ).to_dict("records")

    st.divider()
    st.subheader("Slide Measure Activity — foto")
    fm = st.file_uploader("Unggah foto measure activity (maksimal 4)", type=["png", "jpg", "jpeg"],
                          accept_multiple_files=True, key="up_measure")
    if fm:
        man["foto_measure"] = [f.getvalue() for f in fm[:4]]
    if man["foto_measure"]:
        st.image(man["foto_measure"], width=180)
        if st.button("Hapus foto measure activity"):
            man["foto_measure"] = []

    st.subheader("Slide AR — foto")
    fa = st.file_uploader("Unggah foto AR (maksimal 4)", type=["png", "jpg", "jpeg"],
                          accept_multiple_files=True, key="up_ar")
    if fa:
        man["foto_ar"] = [f.getvalue() for f in fa[:4]]
    if man["foto_ar"]:
        st.image(man["foto_ar"], width=180)
        if st.button("Hapus foto AR"):
            man["foto_ar"] = []

    st.divider()
    st.subheader("Slide Komitmen")
    man["komitmen"] = st.data_editor(
        pd.DataFrame(man["komitmen"]), num_rows="dynamic", use_container_width=True, key="kom_ed",
        column_config={"pencapaian": st.column_config.TextColumn("Pencapaian", width="large"),
                       "komitmen": st.column_config.TextColumn("Komitmen", width="large"),
                       "target": st.column_config.TextColumn("Target", width="medium")}
    ).to_dict("records")

# --- unduh
with tabs[7]:
    st.subheader("Unduh presentasi")
    st.write("Deck mengikuti template standar M-Flash: latar, logo, warna, dan tata letak yang sama.")
    nama = st.text_input("Nama file", f"WEEKLY_MEETING_MFLASH_{dt.date.today():%Y%m%d}.pptx")
    if st.button("🛠️ Buat PPTX", type="primary"):
        manual2 = dict(manual)
        manual2.update(man)
        c2 = CTX.build(dfp, dff, flt, manual2, {"pengiriman_raw": n_raw})
        data = deck.build(c2)
        st.session_state["pptx"] = data
        st.success(f"Selesai — {len(data)/1e6:.1f} MB, 21 slide siap diunduh.")
    if st.session_state.get("pptx"):
        st.download_button("⬇️ Unduh PPTX", st.session_state["pptx"], file_name=nama,
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
