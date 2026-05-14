#!/usr/bin/env python3
"""
CX Dashboard Update Pipeline — Pollo Campero USA
Run every Monday after updating the 2026 data files.
Usage: python update.py
"""

import pandas as pd
import json
import os
import subprocess
import sys
import warnings
from datetime import datetime, date, timedelta

warnings.filterwarnings("ignore")

# ── Configuration ──────────────────────────────────────────────────────────────
DATA_DIR  = r"C:\Users\ITAdmin\Desktop\Claude CX Desktop\Data"
REPO_DIR  = r"C:\Users\ITAdmin\Desktop\cx-dashboard"
GOALS     = {"osat": 90.0, "rating": 4.3, "ms": 90.0}

# ── Helpers ────────────────────────────────────────────────────────────────────
def log(msg): print(f"  {msg}")

def safe_round(v, d=1):
    try: return round(float(v), d)
    except: return None

def week_label(d):
    return d.strftime("%G-W%V")   # ISO week

def top2_pct(series):
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) == 0: return None
    return safe_round((s >= 4).sum() / len(s) * 100)

def avg(series):
    s = pd.to_numeric(series, errors="coerce").dropna()
    return safe_round(s.mean()) if len(s) > 0 else None

def cnt(series):
    return int(series.count())


# ── Step 1: Stores Mapping ──────────────────────────────────────────────────────
def load_mapping():
    log("Loading stores mapping…")
    df = pd.read_excel(os.path.join(DATA_DIR, "Stores mapping.xlsx"))
    df.columns = ["store_id","store_name","area","area_manager",
                  "region_manager","region","country","open"]
    df = df[df["open"] == "Y"].copy()
    df["store_name_lc"] = df["store_name"].str.strip().str.lower()
    return df


# ── Step 2: OSAT ───────────────────────────────────────────────────────────────
def load_osat():
    log("Loading OSAT 2026…")
    df = pd.read_csv(
        os.path.join(DATA_DIR, "OSAT", "OSAT 2026.csv"),
        encoding="latin-1", low_memory=False,
        usecols=["StoreId","VisitDate","Overall Satisfaction"]
    )
    df.columns = ["store_id","visit_date","score"]
    df["score"]      = pd.to_numeric(df["score"], errors="coerce")
    df["visit_date"] = pd.to_datetime(df["visit_date"], errors="coerce")
    df = df.dropna(subset=["store_id","visit_date","score"])
    df["week"]  = df["visit_date"].apply(week_label)
    df["month"] = df["visit_date"].dt.strftime("%Y-%m")
    df["top2"]  = (df["score"] >= 4).astype(float)
    log(f"  OSAT rows: {len(df):,}")
    return df


# ── Step 3: Reviews ────────────────────────────────────────────────────────────
def load_reviews():
    log("Loading Reviews 2026…")
    df = pd.read_csv(
        os.path.join(DATA_DIR, "Social Media", "Reviews 2026.csv"),
        encoding="latin-1", low_memory=False,
        usecols=["Client Location ID","Rating","Review Date"]
    )
    df.columns = ["store_id","rating","review_date"]
    df["rating"]      = pd.to_numeric(df["rating"], errors="coerce")
    df["review_date"] = pd.to_datetime(df["review_date"], errors="coerce")
    df = df.dropna(subset=["store_id","review_date","rating"])
    df["week"]  = df["review_date"].apply(week_label)
    df["month"] = df["review_date"].dt.strftime("%Y-%m")
    log(f"  Reviews rows: {len(df):,}")
    return df


# ── Step 4: Mystery Shopper ────────────────────────────────────────────────────
def load_ms():
    log("Loading Mystery Shopper 2026 + April update…")
    def _read(path):
        return pd.read_excel(path, header=3)

    ms26  = _read(os.path.join(DATA_DIR, "Mystery Shopper", "Mystery Shopper 2026.xlsx"))
    msapr = _read(os.path.join(DATA_DIR, "Mystery Shopper", "Mystery Shopper April 26th to date.xlsx"))

    ms = pd.concat([ms26, msapr], ignore_index=True)
    # Deduplicate by Survey ID if column exists
    if "Survey ID" in ms.columns:
        ms = ms.drop_duplicates(subset=["Survey ID"])

    keep = ["Location ID","Location Name","Region","Director of Operations",
            "Area Manager Region","Date","Total"]
    ms = ms[[c for c in keep if c in ms.columns]].copy()
    ms["date"]  = pd.to_datetime(ms["Date"], errors="coerce")
    ms["Total"] = pd.to_numeric(ms["Total"], errors="coerce")
    ms = ms.dropna(subset=["date","Total","Location Name"])
    ms["month"]    = ms["date"].dt.strftime("%Y-%m")
    ms["name_lc"]  = ms["Location Name"].str.strip().str.lower()
    log(f"  MS rows: {len(ms):,}")
    return ms


# ── Step 5: Atel (Claims) ──────────────────────────────────────────────────────
def load_atel():
    log("Loading Atel 2026…")
    df = pd.read_excel(os.path.join(DATA_DIR, "Atel", "Atel 2026.xlsx"))
    store_col = "Restaurante que Atendión (Reclamo)"
    date_col  = "Hora de inicio"
    df = df.dropna(subset=[date_col, store_col])
    df["date"] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=["date"])
    df["week"]  = df["date"].apply(week_label)
    df["month"] = df["date"].dt.strftime("%Y-%m")

    def _strip_state(s):
        p = str(s).split(" - ", 1)
        return p[1].strip().lower() if len(p) == 2 else str(s).strip().lower()

    df["name_lc"] = df[store_col].apply(_strip_state)
    log(f"  Atel rows: {len(df):,}")
    return df


# ── Step 6: Build lookup from name → store_id ──────────────────────────────────
def build_name_map(mapping_df):
    m = {}
    for _, row in mapping_df.iterrows():
        m[row["store_name_lc"]] = row["store_id"]
    return m


# ── Step 7: Aggregate helpers ──────────────────────────────────────────────────
def agg_osat(df, group_col):
    """Returns dict: group → {ytd, n, weeks:{}, months:{}}"""
    out = {}
    for key, grp in df.groupby(group_col):
        ytd = top2_pct(grp["top2"] * 100)  # already 0/1 so multiply by 100
        # Actually top2 is 0/1 float, top2_pct expects raw score. Use directly:
        ytd_v = safe_round(grp["top2"].mean() * 100)
        n     = len(grp)
        wks   = {w: safe_round(g["top2"].mean()*100) for w,g in grp.groupby("week")}
        mths  = {m: safe_round(g["top2"].mean()*100) for m,g in grp.groupby("month")}
        out[key] = {"ytd": ytd_v, "n": n, "weeks": wks, "months": mths}
    return out

def agg_rating(df, group_col):
    out = {}
    for key, grp in df.groupby(group_col):
        out[key] = {
            "ytd":   avg(grp["rating"]),
            "n":     cnt(grp["rating"]),
            "weeks": {w: avg(g["rating"]) for w,g in grp.groupby("week")},
            "months":{m: avg(g["rating"]) for m,g in grp.groupby("month")}
        }
    return out

def agg_ms(df, group_col):
    out = {}
    for key, grp in df.groupby(group_col):
        out[key] = {
            "ytd":   avg(grp["Total"]),
            "n":     cnt(grp["Total"]),
            "months":{m: avg(g["Total"]) for m,g in grp.groupby("month")}
        }
    return out

def agg_claims(df, group_col):
    out = {}
    for key, grp in df.groupby(group_col):
        out[key] = {
            "ytd":   len(grp),
            "weeks": {w: len(g) for w,g in grp.groupby("week")},
            "months":{m: len(g) for m,g in grp.groupby("month")}
        }
    return out


# ── Step 8: Compute L4W ────────────────────────────────────────────────────────
def last_n_weeks(weeks_dict, n, all_weeks):
    """Return average of last n weeks present in all_weeks."""
    avail = [w for w in all_weeks if w in weeks_dict]
    recent = avail[-n:] if len(avail) >= n else avail
    vals = [weeks_dict[w] for w in recent if weeks_dict[w] is not None]
    return safe_round(sum(vals)/len(vals)) if vals else None


# ── Step 9: Main build ─────────────────────────────────────────────────────────
def build(mapping_df, osat_df, reviews_df, ms_df, atel_df):
    log("Building JSON…")
    name_map = build_name_map(mapping_df)

    # Attach store_id to MS and Atel via name lookup
    ms_df["store_id"]    = ms_df["name_lc"].map(name_map)
    atel_df["store_id"]  = atel_df["name_lc"].map(name_map)

    ms_df    = ms_df.dropna(subset=["store_id"])
    atel_df  = atel_df.dropna(subset=["store_id"])

    # Join region/area onto each dataset
    meta_cols = ["store_id","area","region"]
    m_idx = mapping_df.set_index("store_id")[["area","region"]]

    def _enrich(df):
        return df.join(m_idx, on="store_id")

    osat_df    = _enrich(osat_df)
    reviews_df = _enrich(reviews_df)
    ms_df      = _enrich(ms_df)
    atel_df    = _enrich(atel_df)

    # All time-axis values
    all_weeks  = sorted(set(osat_df["week"].tolist() + reviews_df["week"].tolist()))
    all_months = sorted(set(osat_df["month"].tolist() + reviews_df["month"].tolist()
                            + ms_df["month"].tolist()))

    # ── National ──
    def nat_kpi(o_agg, r_agg, m_agg, c_agg):
        o = o_agg.get("__ALL__", {})
        r = r_agg.get("__ALL__", {})
        ms_v = m_agg.get("__ALL__", {})
        c = c_agg.get("__ALL__", {})

        return {
            "osat":   {**o,   "l4w": last_n_weeks(o.get("weeks",{}),  4, all_weeks)},
            "rating": {**r,   "l4w": last_n_weeks(r.get("weeks",{}),  4, all_weeks)},
            "ms":     {**ms_v},
            "claims": {**c,   "l4w": last_n_weeks(c.get("weeks",{}),  4, all_weeks)}
        }

    osat_df["__ALL__"]    = "__ALL__"
    reviews_df["__ALL__"] = "__ALL__"
    ms_df["__ALL__"]      = "__ALL__"
    atel_df["__ALL__"]    = "__ALL__"

    nat_o = agg_osat(osat_df,    "__ALL__")
    nat_r = agg_rating(reviews_df,"__ALL__")
    nat_m = agg_ms(ms_df,        "__ALL__")
    nat_c = agg_claims(atel_df,  "__ALL__")
    national = nat_kpi(nat_o, nat_r, nat_m, nat_c)

    # ── Regions ──
    reg_o = agg_osat(osat_df,    "region")
    reg_r = agg_rating(reviews_df,"region")
    reg_m = agg_ms(ms_df,        "region")
    reg_c = agg_claims(atel_df,  "region")

    all_regions = sorted(set(mapping_df["region"].dropna().unique()))
    regions = {}
    for reg in all_regions:
        o = reg_o.get(reg, {}); r = reg_r.get(reg, {})
        ms_v = reg_m.get(reg, {}); c = reg_c.get(reg, {})
        regions[reg] = {
            "osat":   {**o,   "l4w": last_n_weeks(o.get("weeks",{}),  4, all_weeks)},
            "rating": {**r,   "l4w": last_n_weeks(r.get("weeks",{}),  4, all_weeks)},
            "ms":     ms_v,
            "claims": {**c,   "l4w": last_n_weeks(c.get("weeks",{}),  4, all_weeks)}
        }

    # ── Areas ──
    area_o = agg_osat(osat_df,    "area")
    area_r = agg_rating(reviews_df,"area")
    area_m = agg_ms(ms_df,        "area")
    area_c = agg_claims(atel_df,  "area")

    area_meta = mapping_df[["area","region","area_manager"]].drop_duplicates("area").dropna(subset=["area"])
    areas = {}
    for _, row in area_meta.iterrows():
        a = row["area"]
        o = area_o.get(a, {}); r = area_r.get(a, {})
        ms_v = area_m.get(a, {}); c = area_c.get(a, {})
        areas[a] = {
            "region":  row["region"],
            "manager": row["area_manager"],
            "osat":    {**o, "l4w": last_n_weeks(o.get("weeks",{}), 4, all_weeks)},
            "rating":  {**r, "l4w": last_n_weeks(r.get("weeks",{}), 4, all_weeks)},
            "ms":      ms_v,
            "claims":  {**c, "l4w": last_n_weeks(c.get("weeks",{}), 4, all_weeks)}
        }

    # ── Stores ──
    store_o = agg_osat(osat_df,    "store_id")
    store_r = agg_rating(reviews_df,"store_id")
    store_m = agg_ms(ms_df,        "store_id")
    store_c = agg_claims(atel_df,  "store_id")

    stores = {}
    for _, row in mapping_df.iterrows():
        sid = row["store_id"]
        o = store_o.get(sid, {}); r = store_r.get(sid, {})
        ms_v = store_m.get(sid, {}); c = store_c.get(sid, {})
        stores[sid] = {
            "name":   row["store_name"],
            "area":   row["area"],
            "region": row["region"],
            "osat":   {**o, "l4w": last_n_weeks(o.get("weeks",{}), 4, all_weeks)} if o else {},
            "rating": {**r, "l4w": last_n_weeks(r.get("weeks",{}), 4, all_weeks)} if r else {},
            "ms":     ms_v,
            "claims": {**c, "l4w": last_n_weeks(c.get("weeks",{}), 4, all_weeks)} if c else {}
        }

    # ── Top / Bottom stores by OSAT YTD ──
    scored = [(sid, d["osat"].get("ytd"), d["osat"].get("n",0))
              for sid, d in stores.items()
              if d["osat"].get("ytd") is not None and d["osat"].get("n",0) >= 30]
    scored.sort(key=lambda x: x[1], reverse=True)
    top5    = [{"id": s[0], "name": stores[s[0]]["name"], "area": stores[s[0]]["area"],
                "osat": s[1], "n": s[2]} for s in scored[:5]]
    bottom5 = [{"id": s[0], "name": stores[s[0]]["name"], "area": stores[s[0]]["area"],
                "osat": s[1], "n": s[2]} for s in scored[-5:]]

    # ── Assemble ──
    last_wk = all_weeks[-1] if all_weeks else ""
    last_mo = all_months[-1] if all_months else ""

    data = {
        "meta": {
            "generated":    date.today().isoformat(),
            "last_week":    last_wk,
            "last_month":   last_mo,
            "goals":        GOALS,
            "all_weeks":    all_weeks[-26:],   # last 26 weeks
            "all_months":   all_months,
            "all_regions":  all_regions,
            "all_areas":    sorted(areas.keys()),
        },
        "national": national,
        "regions":  regions,
        "areas":    areas,
        "stores":   stores,
        "top5":     top5,
        "bottom5":  bottom5,
    }
    return data


# ── Step 10: Git push ──────────────────────────────────────────────────────────
def git_push():
    log("Pushing to GitHub Pages…")
    try:
        os.chdir(REPO_DIR)
        subprocess.run(["git", "add", "cx_data.json", "index.html"], check=True)
        msg = f"Update dashboard data {date.today().isoformat()}"
        result = subprocess.run(["git", "diff", "--cached", "--quiet"])
        if result.returncode == 0:
            log("Nothing changed — skipping commit.")
            return
        subprocess.run(["git", "commit", "-m", msg], check=True)
        subprocess.run(["git", "push"], check=True)
        log("Done! Dashboard is live at https://drgonm5.github.io/cx-dashboard/")
    except subprocess.CalledProcessError as e:
        print(f"\n  ERROR during git push: {e}")
        print("  You may need to authenticate. Run this in a terminal:")
        print(f"  cd {REPO_DIR} && git push")


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    print("\n" + "="*60)
    print("  CX Dashboard Update Pipeline")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("="*60 + "\n")

    mapping_df = load_mapping()
    osat_df    = load_osat()
    reviews_df = load_reviews()
    ms_df      = load_ms()
    atel_df    = load_atel()

    print("\nAggregating data…")
    data = build(mapping_df, osat_df, reviews_df, ms_df, atel_df)

    out_path = os.path.join(REPO_DIR, "cx_data.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, separators=(",", ":"))

    size_kb = os.path.getsize(out_path) / 1024
    log(f"cx_data.json written ({size_kb:.0f} KB)")

    git_push()

    print("\n" + "="*60)
    print("  All done!")
    print("="*60 + "\n")


if __name__ == "__main__":
    main()
