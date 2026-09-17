#!/usr/bin/env python3
"""Hazine Kesin Alım Satım — Günlük Değişim Raporu (BV Portföy formatı).

Kullanım:
  python hazine_rapor.py                 # bugünden önceki son işlem günü
  python hazine_rapor.py --date 16.09.2026 [--excel]
"""
import argparse
import datetime as dt
import html
import io
import json
import sys
import zipfile
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import pandas as pd
import requests
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle, KeepTogether)

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
OUT = ROOT / "output"
UA = {"User-Agent": "Mozilla/5.0"}
BIST = "https://www.borsaistanbul.com"

RED, DARK = "#C00000", "#8B0000"
GREEN = "#2E7D32"
GREY = "#7F7F7F"

MARKETS = {"BAP KES NORMAL EMIRLER PZ (OPSN)", "BAP KES KUCUK EMIRLER PZR (OPSS)"}
TLREF_GT = "TLREF e  Endeksli / Indexed to TLREF"
HAZINE = "Hazine / Turkish Treasury"
SABIT, ISKONTO = "Sabit / Fixed", "İskontolu / Discounted"
DT_MK, HB_MK = "Devlet Tahvili / Government Bond", "Hazine Bonosu/Treasury Bill"

FONT_DIR = Path(matplotlib.get_data_path()) / "fonts" / "ttf"
pdfmetrics.registerFont(TTFont("DejaVu", str(FONT_DIR / "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("DejaVu-Bold", str(FONT_DIR / "DejaVuSans-Bold.ttf")))
pdfmetrics.registerFontFamily("DejaVu", normal="DejaVu", bold="DejaVu-Bold", italic="DejaVu", boldItalic="DejaVu-Bold")
plt.rcParams["font.family"] = "DejaVu Sans"


# ----------------------------------------------------------------------------- yardımcılar
def tr(x, nd=2, sign=False):
    """Türkçe sayı formatı: 1.234,56"""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "–"
    if round(x, nd) == 0:
        x = 0.0
    s = f"{x:+,.{nd}f}" if sign else f"{x:,.{nd}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def dstr(d):
    return pd.Timestamp(d).strftime("%d.%m.%Y") if pd.notna(d) else "–"


# ----------------------------------------------------------------------------- veri
def ttb_path(d):
    return DATA / "ttb" / f"ttb{d:%Y%m%d}3.csv"


def load_ttb(d, required=False):
    """BİST borçlanma araçları günlük bülteni; yoksa indir. Yayımlanmamışsa None."""
    d = pd.Timestamp(d)
    p = ttb_path(d)
    if not p.exists():
        url = f"{BIST}/data/ttb/{d:%Y}/{d:%m}/ttb{d:%Y%m%d}3.zip"
        try:
            r = requests.get(url, headers=UA, timeout=60)
        except requests.RequestException as e:
            if required:
                raise
            print(f"uyarı: {url} alınamadı: {e}", file=sys.stderr)
            return None
        if r.status_code != 200 or r.content[:2] != b"PK":
            if required:
                raise FileNotFoundError(f"Bülten bulunamadı: {url} ({r.status_code})")
            return None
        p.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(io.BytesIO(r.content)) as z:
            name = next(n for n in z.namelist() if n.lower().endswith(".csv"))
            p.write_bytes(z.read(name))
    df = pd.read_csv(p, sep=";", skiprows=[1], dtype=str, encoding_errors="replace")
    num = ["KAPANIS", "AG.ORT. FIYAT/ORAN/SWAP PUANI", "ONCEKI AG.ORT. FIYAT", "ONCEKI KAPANIS",
           "KAPANIS BILESIK GETIRI", "AG. ORT. BILESIK GETIRI", "MIKTAR", "ISLEM HACMI", "ISLEM SAYISI"]
    for c in num:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ["ISLEM TARIHI", "ONCEKI ISLEM TARIHI", "VALOR1"]:
        df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


def load_tbliste():
    """Borçlanma Araçları Listesi — BİST'ten en güncel liste; alınamazsa son kayıtlı liste."""
    folder = DATA / "tbliste"
    folder.mkdir(parents=True, exist_ok=True)
    note = None
    try:
        r = requests.get(f"{BIST}/datum/tbliste.zip", headers=UA, timeout=60)
        r.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(r.content)) as z:
            for n in z.namelist():
                if n.lower().endswith((".xls", ".xlsx")):
                    (folder / Path(n).name).write_bytes(z.read(n))
    except Exception as e:  # noqa: BLE001
        note = f"Borçlanma Araçları Listesi indirilemedi ({e}); son kayıtlı liste kullanıldı."
    files = sorted(folder.glob("tbliste_*.xls*"))
    if not files:
        raise FileNotFoundError("tbliste bulunamadı; data/tbliste klasörüne yükleyin.")
    f = files[-1]
    lst = pd.read_excel(f, sheet_name="Borçlanma Araçları", header=0)
    lst.columns = [str(c).split("\n")[0].strip() for c in lst.columns]
    lst = lst.rename(columns={"ISIN Kodu": "ISIN", "İhraçcı Kurum": "Ihracci", "Getiri Türü": "GetiriTuru",
                              "MK Türü": "MKTuru", "İtfa Tarihi": "Vade", "Para Birimi": "ParaBirimi"})
    lst["Vade"] = pd.to_datetime(lst["Vade"], errors="coerce")
    return lst[["ISIN", "Ihracci", "GetiriTuru", "MKTuru", "Vade", "ParaBirimi"]].drop_duplicates("ISIN"), f.name, note


def last_bulletin_before(today):
    d = pd.Timestamp(today).normalize() - pd.Timedelta(days=1)
    for _ in range(10):
        if d.weekday() < 5 and load_ttb(d) is not None:
            return d
        d -= pd.Timedelta(days=1)
    raise FileNotFoundError("Son 10 günde yayımlanmış bülten bulunamadı.")


# ----------------------------------------------------------------------------- hesaplama
def wavg(v, w):
    v, w = np.asarray(v, float), np.asarray(w, float)
    m = ~np.isnan(v) & ~np.isnan(w) & (w > 0)
    return float(np.sum(v[m] * w[m]) / np.sum(w[m])) if m.any() else np.nan


def aggregate(report_date, lst):
    today = load_ttb(report_date, required=True)
    x = today[today["PAZAR ISMI"].isin(MARKETS) & (today["ISLEM SAYISI"] > 0)].copy()
    x = x.sort_values(["ISIN/KOD", "VALOR1", "ISLEM KODU"])
    cache = {}

    def prev_row(r):
        pdte = r["ONCEKI ISLEM TARIHI"]
        if pd.isna(pdte):
            return None
        if pdte not in cache:
            cache[pdte] = load_ttb(pdte)
        pdf_ = cache[pdte]
        if pdf_ is None:
            return None
        m = pdf_[(pdf_["ISLEM KODU"] == r["ISLEM KODU"]) & (pdf_["PAZAR ISMI"] == r["PAZAR ISMI"])
                 & (pdf_["ISLEM SAYISI"] > 0)]
        return m.iloc[0] if len(m) else None

    rows = []
    for isin, g in x.groupby("ISIN/KOD", sort=False):
        prevs = [prev_row(r) for _, r in g.iterrows()]
        last = g.iloc[-1]
        lastprev = prevs[-1]
        pv = [p for p in prevs if p is not None]
        prev_wavg = wavg([p["AG.ORT. FIYAT/ORAN/SWAP PUANI"] for p in pv], [p["ISLEM HACMI"] for p in pv]) if pv else np.nan
        if np.isnan(prev_wavg) and g["ONCEKI AG.ORT. FIYAT"].notna().any():  # bülten bulunamazsa yedek
            prev_wavg = wavg(g["ONCEKI AG.ORT. FIYAT"], g["ISLEM HACMI"])
        rows.append(dict(
            ISIN=isin,
            Hacim=g["ISLEM HACMI"].sum(), Islem=int(g["ISLEM SAYISI"].sum()), Nominal=g["MIKTAR"].sum(),
            AgOrt=wavg(g["AG.ORT. FIYAT/ORAN/SWAP PUANI"], g["ISLEM HACMI"]),
            OncAgOrt=prev_wavg,
            Kapanis=last["KAPANIS"], OncKapanis=last["ONCEKI KAPANIS"],
            KapGetiri=last["KAPANIS BILESIK GETIRI"],
            OncKapGetiri=lastprev["KAPANIS BILESIK GETIRI"] if lastprev is not None else np.nan,
            AgOrtGetiri=wavg(g["AG. ORT. BILESIK GETIRI"], g["ISLEM HACMI"]),
            OncAgOrtGetiri=(wavg([p["AG. ORT. BILESIK GETIRI"] for p in pv], [p["ISLEM HACMI"] for p in pv])
                            if pv else np.nan),
            OncGun=last["ONCEKI ISLEM TARIHI"],
        ))
    a = pd.DataFrame(rows).merge(lst, on="ISIN", how="left")
    a["KalanGun"] = (a["Vade"] - pd.Timestamp(report_date)).dt.days
    a["FiyatDeg"] = (a["AgOrt"] - a["OncAgOrt"]).round(6)
    a["KapDeg"] = (a["Kapanis"] - a["OncKapanis"]).round(6)
    a["KapGetiriDegBp"] = ((a["KapGetiri"] - a["OncKapGetiri"]) * 100).round(4)
    a["AgOrtGetiriDegBp"] = ((a["AgOrtGetiri"] - a["OncAgOrtGetiri"]) * 100).round(4)

    tlref = a[a["GetiriTuru"] == TLREF_GT].sort_values("Vade").reset_index(drop=True)
    isB = (a["Ihracci"] == HAZINE) & (
        ((a["GetiriTuru"] == SABIT) & (a["MKTuru"] == DT_MK))
        | ((a["GetiriTuru"] == ISKONTO) & a["MKTuru"].isin([DT_MK, HB_MK])))
    b = a[isB].copy()
    b["Tur"] = np.where(b["GetiriTuru"] == SABIT, "Sabit", "İskontolu")
    b = b.sort_values("Vade").reset_index(drop=True)
    b, excluded = exclude_outliers(b)
    return tlref, b, excluded


def exclude_outliers(b):
    """Tek işlem + nominal < 1 mn TL + kapanış bileşik getirisi komşu vadelerin ortalamasından >10 puan sapma."""
    cand = (b["Islem"] == 1) & (b["Nominal"] < 1_000_000) & b["KapGetiri"].notna()
    ref = b[b["KapGetiri"].notna() & ~cand]
    if len(ref) < 2:
        ref = b[b["KapGetiri"].notna()]
    drop, info = [], []
    for i in b.index[cand]:
        r = ref.drop(index=i, errors="ignore")
        before = r[r["Vade"] <= b.at[i, "Vade"]].tail(1)
        after = r[r["Vade"] > b.at[i, "Vade"]].head(1)
        nb = pd.concat([before, after])
        if nb.empty:
            continue
        avg = nb["KapGetiri"].mean()
        if abs(b.at[i, "KapGetiri"] - avg) > 10:
            drop.append(i)
            info.append(dict(ISIN=b.at[i, "ISIN"], Getiri=b.at[i, "KapGetiri"], Komsu=avg,
                             Nominal=b.at[i, "Nominal"]))
    return b.drop(index=drop).reset_index(drop=True), info


# ----------------------------------------------------------------------------- grafikler
def place_labels(ax, xs, ys, labels, polylines, fontsize, radii=(14, 24, 36, 50, 66, 84), leader_from=24):
    """Etiketleri çizgiler, noktalar ve birbirleriyle çakışmayacak şekilde yerleştirir (gerekirse leader line)."""
    fig = ax.figure
    fig.canvas.draw()
    rend = fig.canvas.get_renderer()
    k = fig.dpi / 72.0
    to_disp = ax.transData.transform
    abox = ax.get_window_extent(rend)
    cloud = []
    for lx, ly in polylines:
        pts = to_disp(np.column_stack([lx, ly]))
        for a, b in zip(pts[:-1], pts[1:]):
            n = max(int(np.hypot(*(b - a)) / 2), 2)
            cloud.append(a + (b - a) * np.linspace(0, 1, n)[:, None])
    cloud = np.vstack(cloud) if cloud else np.empty((0, 2))
    marks = to_disp(np.column_stack([xs, ys]))
    placed = []
    angles = [90, -90, 60, 120, -60, -120, 30, 150, -30, -150, 0, 180]
    order = np.argsort(xs)
    for i in order:
        t = ax.text(0, 0, labels[i], fontsize=fontsize, ha="center", va="center", zorder=6)
        bb = t.get_window_extent(rend)
        w, h = bb.width + 4, bb.height + 2
        px, py = marks[i]
        best = None
        for r in radii:
            for ang in angles:
                cx = px + np.cos(np.radians(ang)) * (r * k + w / 2 * abs(np.cos(np.radians(ang))))
                cy = py + np.sin(np.radians(ang)) * (r * k * 0.8 + h / 2 * abs(np.sin(np.radians(ang))))
                x0, x1, y0, y1 = cx - w / 2, cx + w / 2, cy - h / 2, cy + h / 2
                cost = r * 0.6
                if x0 < abox.x0 or x1 > abox.x1 or y0 < abox.y0 or y1 > abox.y1:
                    cost += 1e5
                if len(cloud):
                    cost += 400 * np.sum((cloud[:, 0] > x0) & (cloud[:, 0] < x1) & (cloud[:, 1] > y0) & (cloud[:, 1] < y1))
                cost += 3000 * np.sum((marks[:, 0] > x0 - 4) & (marks[:, 0] < x1 + 4)
                                      & (marks[:, 1] > y0 - 4) & (marks[:, 1] < y1 + 4))
                for (a0, a1, b0, b1) in placed:
                    ox = max(0, min(x1, a1) - max(x0, a0))
                    oy = max(0, min(y1, b1) - max(y0, b0))
                    cost += 50 * ox * oy + (2000 if ox * oy > 0 else 0)
                if best is None or cost < best[0]:
                    best = (cost, cx, cy, r, (x0, x1, y0, y1))
            if best[0] < 60:
                break
        _, cx, cy, r, box = best
        placed.append(box)
        dx, dy = ax.transData.inverted().transform((cx, cy))
        t.set_position((dx, dy))
        if r >= leader_from:
            ex = min(max(px, box[0]), box[1])
            ey = min(max(py, box[2]), box[3])
            lx, ly = ax.transData.inverted().transform([(px, py), (ex, ey)]).T
            ax.plot(lx, ly, color="#9A9A9A", lw=0.5, zorder=4)


def _date_axis(ax):
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m.%Y"))
    ax.grid(True, color="#E6E6E6", lw=0.6)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)


def chart_tlref(t, report_date, path):
    t = t.dropna(subset=["AgOrt"])
    fig, ax = plt.subplots(figsize=(11, 3.6), dpi=160)
    p = t.dropna(subset=["OncAgOrt"])
    ax.plot(p["Vade"], p["OncAgOrt"], ls="--", color=GREY, lw=1.2, label="Önceki işlem günü")
    ax.plot(t["Vade"], t["AgOrt"], color=RED, lw=1.6, label=dstr(report_date))
    col = [GREEN if (not np.isnan(v) and v >= 0) else (RED if not np.isnan(v) else GREY) for v in t["FiyatDeg"]]
    ax.scatter(t["Vade"], t["AgOrt"], c=col, s=34, zorder=5, edgecolor="white", lw=0.6)
    xs = mdates.date2num(t["Vade"])
    labels = [f"{r.ISIN}\n{tr(r.FiyatDeg, 3, True)}" for r in t.itertuples()]
    ax.set_ylabel("Ağ.Ort. Temiz Fiyat")
    ax.set_title("TLREF'e endeksli kağıtlar — Ağ.Ort. temiz fiyat (etiket: fiyat değişimi, puan)",
                 fontsize=9.5, color=DARK, loc="left")
    _date_axis(ax)
    ax.legend(fontsize=7.5, frameon=False, loc="lower right", bbox_to_anchor=(1, 1.0), ncol=2)
    lo, hi = t["AgOrt"].min(), t["AgOrt"].max()
    pad = max((hi - lo) * 0.35, 0.3)
    ax.set_ylim(lo - pad, hi + pad)
    span = xs.max() - xs.min() if len(xs) > 1 else 60
    ax.set_xlim(xs.min() - span * 0.06, xs.max() + span * 0.06)
    fig.tight_layout()
    place_labels(ax, xs, t["AgOrt"].to_numpy(), labels,
                 [(xs, t["AgOrt"].to_numpy()), (mdates.date2num(p["Vade"]), p["OncAgOrt"].to_numpy())], 6.3)
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


def chart_b_curve(b, report_date, path, figsize):
    b = b.dropna(subset=["KapGetiri"])
    fig, ax = plt.subplots(figsize=figsize, dpi=160)
    p = b.dropna(subset=["OncKapGetiri"])
    ax.plot(p["Vade"], p["OncKapGetiri"], ls="--", color=GREY, lw=1.2, label="Önceki işlem günü")
    ax.plot(b["Vade"], b["KapGetiri"], color=RED, lw=1.6, label=dstr(report_date))
    for tur, mk in (("Sabit", "o"), ("İskontolu", "s")):
        s = b[b["Tur"] == tur]
        ax.scatter(s["Vade"], s["KapGetiri"], marker=mk, s=30, color=DARK, zorder=5,
                   edgecolor="white", lw=0.6, label=tur)
    xs = mdates.date2num(b["Vade"])
    labels = [f"{r.ISIN} {tr(r.KapGetiriDegBp, 1, True)}" for r in b.itertuples()]
    ax.set_ylabel("Kapanış bileşik getiri (%)")
    ax.set_title("Sabit kuponlu + İskontolu — Kapanış bileşik getiri (etiket: değişim, bp)",
                 fontsize=9.5, color=DARK, loc="left")
    _date_axis(ax)
    ax.legend(fontsize=7, frameon=False, loc="lower right", bbox_to_anchor=(1, 1.0), ncol=4)
    lo, hi = b["KapGetiri"].min(), b["KapGetiri"].max()
    ax.set_ylim(lo - (hi - lo) * 0.3 - 0.4, hi + (hi - lo) * 0.35 + 0.4)
    span = xs.max() - xs.min() if len(xs) > 1 else 60
    ax.set_xlim(xs.min() - span * 0.08, xs.max() + span * 0.06)
    fig.tight_layout()
    place_labels(ax, xs, b["KapGetiri"].to_numpy(), labels,
                 [(xs, b["KapGetiri"].to_numpy()), (mdates.date2num(p["Vade"]), p["OncKapGetiri"].to_numpy())],
                 6.0, radii=(20, 30, 42, 56, 72, 90, 110, 132), leader_from=0)
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


def chart_b_bars(b, path, figsize):
    s = b.dropna(subset=["KapGetiriDegBp"]).sort_values("KapGetiriDegBp", ascending=True)
    fig, ax = plt.subplots(figsize=figsize, dpi=160)
    col = [RED if v > 0 else GREEN for v in s["KapGetiriDegBp"]]
    ax.barh(s["ISIN"], s["KapGetiriDegBp"], color=col, height=0.7)
    for i, v in enumerate(s["KapGetiriDegBp"]):
        ax.text(v, i, f" {tr(v, 1, True)} ", va="center", ha="left" if v >= 0 else "right", fontsize=6)
    ax.axvline(0, color="#444444", lw=0.6)
    ax.set_xlabel("Kapanış bileşik getiri değişimi (bp)")
    ax.set_title("Kapanış getiri değişimi (bp) — büyükten küçüğe", fontsize=9.5, color=DARK, loc="left")
    ax.tick_params(axis="y", labelsize=6.2)
    ax.grid(True, axis="x", color="#E6E6E6", lw=0.6)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    m = max(abs(s["KapGetiriDegBp"]).max(), 1) * 1.25
    ax.set_xlim(-m, m)
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


# ----------------------------------------------------------------------------- metin
def bucket(days):
    if days < 365:
        return "0-1 yıl"
    if days < 3 * 365:
        return "1-3 yıl"
    if days < 5 * 365:
        return "3-5 yıl"
    return "5 yıl+"


def summary_numbers(t, b):
    return dict(
        tlref_hacim=t["Hacim"].sum(), tlref_islem=int(t["Islem"].sum()),
        tlref_fiyat_deg=wavg(t["FiyatDeg"], t["Hacim"]),
        b_hacim=b["Hacim"].sum(), b_islem=int(b["Islem"].sum()),
        b_bp=wavg(b["KapGetiriDegBp"], b["Hacim"]),
        b_agort_bp=wavg(b["AgOrtGetiriDegBp"], b["Hacim"]),
    )


def highlights(t, b, s):
    out = []
    bb = b.dropna(subset=["KapGetiriDegBp"])
    if len(bb):
        up, dn = bb.loc[bb["KapGetiriDegBp"].idxmax()], bb.loc[bb["KapGetiriDegBp"].idxmin()]
        n_up, n_dn = int((bb["KapGetiriDegBp"] > 0).sum()), int((bb["KapGetiriDegBp"] < 0).sum())
        out.append(
            f"<b>En sert hareketler:</b> Sabit kuponlu/iskontolu grupta kapanış getirisindeki en büyük artış "
            f"{up.ISIN} ({up.Tur}, vade {dstr(up.Vade)}) kağıdında {tr(up.KapGetiriDegBp, 1, True)} bp "
            f"(%{tr(up.OncKapGetiri)} → %{tr(up.KapGetiri)}), en büyük düşüş {dn.ISIN} ({dn.Tur}, vade "
            f"{dstr(dn.Vade)}) kağıdında {tr(dn.KapGetiriDegBp, 1, True)} bp (%{tr(dn.OncKapGetiri)} → "
            f"%{tr(dn.KapGetiri)}). Karşılaştırılabilir {len(bb)} kağıt: getirisi artan {n_up}, düşen {n_dn}"
            f"{', değişmeyen ' + str(len(bb) - n_up - n_dn) if len(bb) - n_up - n_dn else ''}.")
        g = bb.assign(B=bb["KalanGun"].map(bucket)).groupby("B", sort=False)
        order = ["0-1 yıl", "1-3 yıl", "3-5 yıl", "5 yıl+"]
        parts = []
        for k in order:
            if k in g.groups:
                d = g.get_group(k)
                parts.append(f"{k}: {tr(wavg(d['KapGetiriDegBp'], d['Hacim']), 1, True)} bp "
                             f"({tr(d['Hacim'].sum() / 1e6, 0)} mn TL, {len(d)} kağıt)")
        out.append("<b>Vade bölgeleri (hacim ağırlıklı kapanış getiri değişimi):</b> " + "; ".join(parts) + ".")
    tt = t.dropna(subset=["AgOrt"]).sort_values("Hacim", ascending=False).head(3)
    if len(tt):
        share = tt["Hacim"].sum() / s["tlref_hacim"] * 100
        parts = [f"{r.ISIN} ({tr(r.Hacim / 1e6, 0)} mn TL; Ağ.Ort. {tr(r.AgOrt, 3)}, "
                 f"{tr(r.FiyatDeg, 3, True)} puan)" for r in tt.itertuples()]
        out.append(f"<b>TLREF'te en likit kağıtlar:</b> " + "; ".join(parts) +
                   f". Bu üç kağıdın TLREF hacmindeki payı %{tr(share, 1)}; grubun hacim ağırlıklı "
                   f"fiyat değişimi {tr(s['tlref_fiyat_deg'], 3, True)} puan.")
    if len(bb):
        yon = "yükseldi" if s["b_bp"] > 0 else "geriledi"
        bz = bb.assign(B=bb["KalanGun"].map(bucket)).groupby("B")["Hacim"].sum()
        out.append(
            f"<b>Kısa yorum:</b> Sabit kuponlu/iskontolu grupta getiriler hacim ağırlıklı olarak "
            f"{tr(abs(s['b_bp']), 1)} bp {yon} (Ağ.Ort. getiri bazında {tr(s['b_agort_bp'], 1, True)} bp). "
            f"Hacimde en büyük pay %{tr(bz.max() / bz.sum() * 100, 0)} ile {bz.idxmax()} vade bölgesinde.")
    return out


# ----------------------------------------------------------------------------- PDF
PAGE = landscape(A4)
ST = dict(
    h1=ParagraphStyle("h1", fontName="DejaVu-Bold", fontSize=16, textColor=colors.HexColor(DARK), leading=20),
    h2=ParagraphStyle("h2", fontName="DejaVu-Bold", fontSize=11, textColor=colors.HexColor(RED), leading=14,
                      spaceBefore=6, spaceAfter=3),
    p=ParagraphStyle("p", fontName="DejaVu", fontSize=8.6, leading=11.6, spaceAfter=4),
    small=ParagraphStyle("s", fontName="DejaVu", fontSize=7.4, leading=9.6, spaceAfter=2),
    kpi_l=ParagraphStyle("kl", fontName="DejaVu", fontSize=8, textColor=colors.HexColor("#555555"), leading=10),
    kpi_v=ParagraphStyle("kv", fontName="DejaVu-Bold", fontSize=17, textColor=colors.HexColor(DARK), leading=21),
)


def page_deco(report_date, source_note):
    def deco(c, doc):
        w, h = PAGE
        c.saveState()
        c.setFillColor(colors.HexColor(RED))
        c.rect(0, h - 26, w, 26, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("DejaVu-Bold", 10.5)
        c.drawString(14 * mm, h - 17, "BV Portföy | Hazine Kesin Alım Satım — Günlük Değişim Raporu")
        c.setFont("DejaVu", 10)
        c.drawRightString(w - 14 * mm, h - 17, dstr(report_date))
        c.setStrokeColor(colors.HexColor("#CCCCCC"))
        c.line(14 * mm, 20, w - 14 * mm, 20)
        c.setFillColor(colors.HexColor("#555555"))
        c.setFont("DejaVu", 6.8)
        c.drawString(14 * mm, 11, source_note + "  Yalnızca bilgi amaçlıdır; yatırım tavsiyesi değildir.")
        c.drawRightString(w - 14 * mm, 11, f"Sayfa {doc.page}")
        c.restoreState()
    return deco


def colored(v, txt, good_if_positive):
    if v is None or np.isnan(v) or v == 0:
        return txt
    good = (v > 0) == good_if_positive
    return f'<font color="{GREEN if good else RED}">{txt}</font>'


def make_table(header, rows, col_w, font=6.6, total=True, pad=1.2):
    cell = ParagraphStyle("c", fontName="DejaVu", fontSize=font, leading=font + 1.6, alignment=2)
    cell_l = ParagraphStyle("cl", parent=cell, alignment=0)
    hdr = ParagraphStyle("h", fontName="DejaVu-Bold", fontSize=font, leading=font + 1.6, alignment=1,
                         textColor=colors.white)
    data = [[Paragraph(h, hdr) for h in header]]
    for r in rows:
        data.append([Paragraph(str(v), cell_l if i == 0 or (i == 1 and not str(v)[:1].isdigit() and str(v) not in ("–",))
                               else cell) for i, v in enumerate(r)])
    tb = Table(data, colWidths=col_w, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(DARK)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F0F0")]),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
        ("TOPPADDING", (0, 0), (-1, -1), pad), ("BOTTOMPADDING", (0, 0), (-1, -1), pad),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.5), ("RIGHTPADDING", (0, 0), (-1, -1), 2.5),
    ]
    if total:
        style += [("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#EBD5D5")),
                  ("LINEABOVE", (0, -1), (-1, -1), 0.8, colors.HexColor(DARK))]
    tb.setStyle(TableStyle(style))
    return tb


def tlref_table(t, s):
    header = ["ISIN", "Vade", "Kalan Gün", "Hacim mn TL", "İşlem", "Önceki Ağ.Ort.", "Ağ.Ort.", "Değişim",
              "Önceki Kapanış", "Kapanış", "Değişim", "Önceki İşlem Günü"]
    rows = [[r.ISIN, dstr(r.Vade), tr(r.KalanGun, 0), tr(r.Hacim / 1e6, 1), r.Islem, tr(r.OncAgOrt, 3),
             tr(r.AgOrt, 3), colored(r.FiyatDeg, tr(r.FiyatDeg, 3, True), True), tr(r.OncKapanis, 3),
             tr(r.Kapanis, 3), colored(r.KapDeg, tr(r.KapDeg, 3, True), True), dstr(r.OncGun)]
            for r in t.itertuples()]
    kd = wavg(t["KapDeg"], t["Hacim"])
    rows.append(["<b>Toplam</b>", "", "", f"<b>{tr(s['tlref_hacim'] / 1e6, 1)}</b>", f"<b>{s['tlref_islem']}</b>",
                 "", "", colored(s["tlref_fiyat_deg"], f"<b>{tr(s['tlref_fiyat_deg'], 3, True)}*</b>", True),
                 "", "", colored(kd, f"<b>{tr(kd, 3, True)}*</b>", True), ""])
    w = [66, 52, 40, 52, 34, 56, 52, 50, 56, 52, 50, 62]
    return make_table(header, rows, [x * 1.1 for x in w])


def b_table(b, s, font):
    header = ["ISIN", "Tür", "Vade", "Kalan Gün", "Hacim mn TL", "İşlem", "Ağ.Ort. Fiyat", "Fiyat Değ.",
              "Önceki Kap. Getiri", "Kapanış Getiri", "Kap. Getiri Değ. bp", "Ağ.Ort. Getiri",
              "Ağ.Ort. Getiri Değ. bp", "Önceki İşlem Günü"]
    rows = [[r.ISIN, r.Tur, dstr(r.Vade), tr(r.KalanGun, 0), tr(r.Hacim / 1e6, 1), r.Islem, tr(r.AgOrt, 3),
             colored(r.FiyatDeg, tr(r.FiyatDeg, 3, True), True), tr(r.OncKapGetiri), tr(r.KapGetiri),
             colored(r.KapGetiriDegBp, tr(r.KapGetiriDegBp, 1, True), False), tr(r.AgOrtGetiri),
             colored(r.AgOrtGetiriDegBp, tr(r.AgOrtGetiriDegBp, 1, True), False), dstr(r.OncGun)]
            for r in b.itertuples()]
    fd = wavg(b["FiyatDeg"], b["Hacim"])
    rows.append(["<b>Toplam</b>", "", "", "", f"<b>{tr(s['b_hacim'] / 1e6, 1)}</b>", f"<b>{s['b_islem']}</b>", "",
                 colored(fd, f"<b>{tr(fd, 3, True)}*</b>", True), "", "",
                 colored(s["b_bp"], f"<b>{tr(s['b_bp'], 1, True)}*</b>", False), "",
                 colored(s["b_agort_bp"], f"<b>{tr(s['b_agort_bp'], 1, True)}*</b>", False), ""])
    w = [64, 46, 50, 40, 50, 32, 52, 48, 54, 50, 54, 50, 56, 58]
    return make_table(header, rows, [x * 1.05 for x in w], font=font, pad=2.2 if font >= 7.4 else 1.2)


def kpi_boxes(s):
    items = [("TLREF'e endeksli hacim", f"{tr(s['tlref_hacim'] / 1e6, 0)} mn TL", f"{s['tlref_islem']} işlem"),
             ("TLREF hacim ağırlıklı fiyat değişimi", f"{tr(s['tlref_fiyat_deg'], 3, True)} puan", "Ağ.Ort. temiz fiyat"),
             ("Sabit kuponlu + İskontolu hacim", f"{tr(s['b_hacim'] / 1e6, 0)} mn TL", f"{s['b_islem']} işlem"),
             ("Sabit + İskontolu hacim ağırlıklı getiri değişimi", f"{tr(s['b_bp'], 1, True)} bp",
              "Kapanış bileşik getiri")]
    cells = [[Paragraph(l, ST["kpi_l"]), Paragraph(v, ST["kpi_v"]), Paragraph(n, ST["kpi_l"])] for l, v, n in items]
    inner = [Table([[c[0]], [c[1]], [c[2]]], colWidths=[180]) for c in cells]
    tb = Table([inner], colWidths=[193] * 4)
    tb.setStyle(TableStyle([
        ("BOX", (0, 0), (0, 0), 0.8, colors.HexColor(RED)), ("BOX", (1, 0), (1, 0), 0.8, colors.HexColor(RED)),
        ("BOX", (2, 0), (2, 0), 0.8, colors.HexColor(RED)), ("BOX", (3, 0), (3, 0), 0.8, colors.HexColor(RED)),
        ("LINEBEFORE", (0, 0), (-1, -1), 3, colors.HexColor(RED)),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
    return tb


def build_pdf(report_date, t, b, excluded, s, hl, list_name, list_note, pdf_path, tmp):
    src = (f"Kaynak: Borsa İstanbul BAP günlük bültenleri (OPSN+OPSS kesin alım satım), Borçlanma Araçları "
           f"Listesi ({list_name}); BV Portföy hesaplamaları.")
    doc = SimpleDocTemplate(str(pdf_path), pagesize=PAGE, leftMargin=14 * mm, rightMargin=14 * mm,
                            topMargin=34, bottomMargin=28, title="Hazine Günlük Değişim Raporu",
                            author="BV Portföy")
    avail_w, avail_h = PAGE[0] - 28 * mm, PAGE[1] - 34 - 28
    story = []

    # Sayfa 1
    story += [Paragraph("Hazine Kağıtları Kesin Alım Satım — Günlük Değişim", ST["h1"]),
              Paragraph(f"İşlem tarihi: {dstr(report_date)} · Karşılaştırma: her kağıdın bir önceki işlem günü",
                        ST["small"]), Spacer(1, 6), kpi_boxes(s), Spacer(1, 4),
              Paragraph("Öne çıkanlar", ST["h2"])]
    story += [Paragraph(x, ST["p"]) for x in hl]
    exc = ", ".join(f"{e['ISIN']} (kapanış getirisi %{tr(e['Getiri'])}, komşu ort. %{tr(e['Komsu'])})"
                    for e in excluded) or "yok"
    method = [
        "Pazar: BAP Kesin Alım Satım Normal (OPSN) ve Küçük Emirler (OPSS); işlem sayısı > 0. T0 ve T1 valörlü "
        "işlem kodları ISIN bazında birleştirildi.",
        "Kategori A: Getiri türü “TLREF’e endeksli” kağıtlar (TLREF’e dayalı değişken faizli ve TLREFK kağıtları "
        "hariç); BİST bileşik getiri yayımlamadığından değişim temiz fiyat bazındadır. Kategori B: Hazine sabit "
        "kuponlu devlet tahvilleri ve iskontolu bono/tahviller (kira sertifikaları hariç).",
        "Değişimler BİST'in yüzde değişim kolonları yerine, aynı işlem kodunun önceki işlem günü bülten "
        "satırından hesaplandı: Ağ.Ort. fiyat işlem hacmi ağırlıklı; kapanış son valörlü koddan. Fiyat "
        "değişimi puan, getiri değişimi baz puan (bp). Toplam satırlarında * hacim ağırlıklı ortalamadır.",
        f"Hariç tutulanlar (tek işlem, nominal < 1 mn TL ve komşu vadelere göre >10 puan sapma): {exc}.",
    ]
    if list_note:
        method.append(list_note)
    story += [Paragraph("Kapsam ve yöntem", ST["h2"])] + [Paragraph("• " + m, ST["small"]) for m in method]
    story.append(PageBreak())

    # Sayfa 2: TLREF tablo + G1
    tt = tlref_table(t, s)
    _, th = tt.wrap(avail_w, avail_h)
    g1 = tmp / "g1.png"
    chart_tlref(t, report_date, g1)
    head = Paragraph("Kategori A — TLREF'e endeksli kağıtlar", ST["h2"])
    rem = avail_h - th - 30
    story += [head, tt, Spacer(1, 6), fit_image(g1, avail_w, max(rem, 120))]
    story.append(PageBreak())

    # Sayfa 3: Kategori B tablosu (tek sayfaya sığacak font)
    for font in (8.0, 7.4, 6.8, 6.2, 5.6, 5.0):
        bt = b_table(b, s, font)
        _, bh = bt.wrap(avail_w, avail_h)
        if bh <= avail_h - 24:
            break
    story += [Paragraph("Kategori B — Sabit kuponlu ve İskontolu Hazine kağıtları (vadeye göre)", ST["h2"]), bt,
              PageBreak()]

    # Sayfa 4: G2 + G3
    g2, g3 = tmp / "g2.png", tmp / "g3.png"
    nb = b["KapGetiriDegBp"].notna().sum()
    chart_b_curve(b, report_date, g2, figsize=(12, 4.6))
    chart_b_bars(b, g3, figsize=(12, max(2.6, 0.13 * nb + 0.9)))
    story += [fit_image(g2, avail_w, avail_h * 0.55), Spacer(1, 4), fit_image(g3, avail_w, avail_h * 0.43)]
    deco = page_deco(report_date, src)
    doc.build(story, onFirstPage=deco, onLaterPages=deco)
    return [g1, g2, g3]


def fit_image(path, max_w, max_h):
    from PIL import Image as PImage
    with PImage.open(path) as im:
        w, h = im.size
    k = min(max_w / w, max_h / h)
    return Image(str(path), width=w * k, height=h * k)


# ----------------------------------------------------------------------------- Excel
def write_excel(t, b, path):
    tc = {"ISIN": "ISIN", "Vade": "Vade", "KalanGun": "Kalan Gün", "Hacim": "Hacim TL", "Islem": "İşlem",
          "OncAgOrt": "Önceki Ağ.Ort.", "AgOrt": "Ağ.Ort.", "FiyatDeg": "Ağ.Ort. Değişim (puan)",
          "OncKapanis": "Önceki Kapanış", "Kapanis": "Kapanış", "KapDeg": "Kapanış Değişim (puan)",
          "OncGun": "Önceki İşlem Günü"}
    bc = {"ISIN": "ISIN", "Tur": "Tür", "Vade": "Vade", "KalanGun": "Kalan Gün", "Hacim": "Hacim TL",
          "Islem": "İşlem", "AgOrt": "Ağ.Ort. Fiyat", "FiyatDeg": "Fiyat Değ. (puan)",
          "OncKapGetiri": "Önceki Kap. Getiri", "KapGetiri": "Kapanış Getiri",
          "KapGetiriDegBp": "Kap. Getiri Değ. bp", "AgOrtGetiri": "Ağ.Ort. Getiri",
          "AgOrtGetiriDegBp": "Ağ.Ort. Getiri Değ. bp", "OncGun": "Önceki İşlem Günü"}
    with pd.ExcelWriter(path) as xw:
        t[list(tc)].rename(columns=tc).to_excel(xw, sheet_name="TLREF", index=False)
        b[list(bc)].rename(columns=bc).to_excel(xw, sheet_name="Sabit+Iskontolu", index=False)


# ----------------------------------------------------------------------------- e-posta içeriği
def mail_summary(t, b, excluded, s):
    """E-posta için 3-4 cümlelik, tamamen tablolardan türetilmiş özet."""
    out = []
    bb = b.dropna(subset=["KapGetiriDegBp"])
    if len(bb):
        m = bb.loc[bb["KapGetiriDegBp"].abs().idxmax()]
        out.append(f"Günün en büyük hareketi {m.ISIN} ({m.Tur}, vade {dstr(m.Vade)}) kağıdında: kapanış bileşik "
                   f"getirisi {tr(m.KapGetiriDegBp, 1, True)} bp (%{tr(m.OncKapGetiri)} → %{tr(m.KapGetiri)}).")
        n_up, n_dn = int((bb["KapGetiriDegBp"] > 0).sum()), int((bb["KapGetiriDegBp"] < 0).sum())
        out.append(f"Sabit kuponlu + iskontolu kağıtlarda {tr(s['b_hacim'] / 1e6, 0)} mn TL hacimle kapanış getirisi "
                   f"hacim ağırlıklı {tr(s['b_bp'], 1, True)} bp değişti (getirisi artan {n_up}, düşen {n_dn} kağıt).")
    tt = t.dropna(subset=["FiyatDeg"]).sort_values("Hacim", ascending=False)
    if len(tt):
        top = tt.iloc[0]
        out.append(f"TLREF'e endeksli kağıtlarda {tr(s['tlref_hacim'] / 1e6, 0)} mn TL hacimle Ağ.Ort. fiyat hacim "
                   f"ağırlıklı {tr(s['tlref_fiyat_deg'], 3, True)} puan değişti; en likit kağıt {top.ISIN} "
                   f"({tr(top.Hacim / 1e6, 0)} mn TL) {tr(top.FiyatDeg, 3, True)} puan.")
    if excluded:
        out.append("Piyasa dışı fiyat görünümlü tek işlemler nedeniyle hariç tutulanlar: "
                   + ", ".join(e["ISIN"] for e in excluded) + ".")
    return out


def mail_content(rd, t, b, excluded, s, list_note):
    kpis = [("TLREF'e endeksli hacim", f"{tr(s['tlref_hacim'] / 1e6, 0)} mn TL", f"{s['tlref_islem']} işlem"),
            ("TLREF hacim ağırlıklı fiyat değişimi", f"{tr(s['tlref_fiyat_deg'], 3, True)} puan", "Ağ.Ort. temiz fiyat"),
            ("Sabit kuponlu + İskontolu hacim", f"{tr(s['b_hacim'] / 1e6, 0)} mn TL", f"{s['b_islem']} işlem"),
            ("Sabit + İskontolu hacim ağırlıklı getiri değişimi", f"{tr(s['b_bp'], 1, True)} bp",
             "Kapanış bileşik getiri")]
    summ = mail_summary(t, b, excluded, s)
    subject = f"Hazine Günlük Değişim Raporu — {dstr(rd)}"
    cells = "".join(
        f'<td style="width:25%;padding:10px 12px;border:1px solid {RED};border-left:4px solid {RED};'
        f'vertical-align:top;font-family:Arial,Helvetica,sans-serif">'
        f'<div style="font-size:12px;color:#555">{html.escape(l)}</div>'
        f'<div style="font-size:22px;font-weight:bold;color:{DARK};margin:4px 0">{html.escape(v)}</div>'
        f'<div style="font-size:11px;color:#777">{html.escape(n)}</div></td>' for l, v, n in kpis)
    paras = "".join(f'<p style="margin:0 0 8px">{html.escape(x)}</p>' for x in summ)
    note = f'<p style="color:#8B0000">{html.escape(list_note)}</p>' if list_note else ""
    body_html = f"""<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;max-width:860px">
<div style="background:{RED};color:#fff;padding:10px 14px;font-weight:bold">BV Portföy | Hazine Kesin Alım Satım — Günlük Değişim Raporu
<span style="float:right;font-weight:normal">{dstr(rd)}</span></div>
<p>Merhaba,</p>
<p>{dstr(rd)} tarihli Hazine Kesin Alım Satım Günlük Değişim Raporu ektedir (PDF ve Excel).</p>
<table cellspacing="6" cellpadding="0" style="border-collapse:separate;width:100%"><tr>{cells}</tr></table>
<h3 style="color:{RED};margin:14px 0 6px">Özet</h3>{paras}{note}
<p style="font-size:11px;color:#777;border-top:1px solid #ddd;padding-top:6px">Kaynak: Borsa İstanbul BAP günlük
bültenleri (OPSN+OPSS kesin alım satım); BV Portföy hesaplamaları. Değişimler her kağıdın bir önceki işlem
gününe göredir. Yalnızca bilgi amaçlıdır; yatırım tavsiyesi değildir.</p></div>"""
    body_txt = "\n".join(
        ["Merhaba,", "", f"{dstr(rd)} tarihli Hazine Kesin Alım Satım Günlük Değişim Raporu ektedir (PDF ve Excel).", ""]
        + [f"- {l}: {v} ({n})" for l, v, n in kpis] + [""] + summ + ([list_note] if list_note else [])
        + ["", "Kaynak: Borsa İstanbul BAP günlük bültenleri; BV Portföy hesaplamaları.",
           "Yalnızca bilgi amaçlıdır; yatırım tavsiyesi değildir."])
    return subject, body_html, body_txt, summ


# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", help="GG.AA.YYYY (varsayılan: bugünden önceki son bülten günü)")
    ap.add_argument("--excel", action="store_true")
    args = ap.parse_args()
    rd = (pd.Timestamp(dt.datetime.strptime(args.date, "%d.%m.%Y")) if args.date
          else last_bulletin_before(dt.date.today()))

    lst, list_name, list_note = load_tbliste()
    t, b, excluded = aggregate(rd, lst)
    s = summary_numbers(t, b)
    hl = highlights(t, b, s)

    OUT.mkdir(exist_ok=True)
    tmp = OUT / f"_work_{rd:%Y%m%d}"
    tmp.mkdir(exist_ok=True)
    pdf = OUT / f"Hazine_Gunluk_Degisim_Raporu_{rd:%Y%m%d}.pdf"
    build_pdf(rd, t, b, excluded, s, hl, list_name, list_note, pdf, tmp)
    xlsx = None
    if args.excel:
        xlsx = OUT / f"Hazine_Gunluk_Degisim_Raporu_{rd:%Y%m%d}.xlsx"
        write_excel(t, b, xlsx)

    # sayfa görüntüleri (kontrol için)
    import pypdfium2 as pdfium
    pages = []
    for i, page in enumerate(pdfium.PdfDocument(str(pdf))):
        p = tmp / f"page{i + 1}.png"
        page.render(scale=1.6).to_pil().save(p)
        pages.append(str(p))

    subject, body_html, body_txt, summ = mail_content(rd, t, b, excluded, s, list_note)
    (OUT / "mail.html").write_text(body_html)
    (OUT / "mail.txt").write_text(body_txt)

    rel = lambda q: str(Path(q).relative_to(ROOT)) if q else None  # noqa: E731
    result = dict(
        rapor_tarihi=dstr(rd), rapor_tarihi_iso=f"{rd:%Y%m%d}", pdf=rel(pdf), excel=rel(xlsx),
        sayfa_goruntuleri=[rel(q) for q in pages],
        grafikler=[rel(tmp / f"g{i}.png") for i in (1, 2, 3)],
        mail_konu=subject, mail_html="output/mail.html", mail_txt="output/mail.txt", mail_ozet=summ,
        tlref_kagit=len(t), b_kagit=len(b),
        hariç_tutulanlar=[e["ISIN"] for e in excluded], liste=list_name, liste_notu=list_note,
        ozet={k: (round(v, 4) if isinstance(v, float) else v) for k, v in s.items()},
        one_cikanlar=[x.replace("<b>", "").replace("</b>", "") for x in hl],
    )
    (tmp / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2))
    (OUT / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2))
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
