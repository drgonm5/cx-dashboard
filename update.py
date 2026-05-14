#!/usr/bin/env python3
"""
CX Dashboard — weekly data pipeline (v12 format).
Run every Monday after adding new rows to the 2026 source files.
"""
import os, sys, json, subprocess, glob, re
from datetime import datetime

import pandas as pd
import numpy as np

# ─────────────────────── CONFIG ───────────────────────────────────────────────
DATA_DIR = r"C:\Users\ITAdmin\Desktop\Claude CX Desktop\Data"
REPO_DIR = r"C:\Users\ITAdmin\Desktop\cx-dashboard"
OUT_JSON = os.path.join(REPO_DIR, "cx_data.json")

GOALS = {"osat": 90.0, "rating": 4.3, "ms": 90.0}

OSAT_CAT_COLS = {
    "taste":        "Taste of Food",
    "temperature":  "Temperature of Food",
    "cleanliness":  "Cleanliness of Restaurant",
    "friendliness": "Friendliness of Staff",
    "speed":        "Speed of Service",
    "accuracy":     "Accuracy of Order",
    "availability": "Availability of menu items",
    "value":        "Value for price paid",
}

# ─────────────────────── HELPERS ──────────────────────────────────────────────
def safe(v, n=1):
    if v is None:
        return None
    if isinstance(v, float) and (np.isnan(v) or np.isinf(v)):
        return None
    return round(float(v), n)


def t2b(series):
    """Top 2 Box: % of non-null numeric values >= 4 on a 1-5 scale."""
    s = pd.to_numeric(series, errors="coerce").dropna()
    return round((s >= 4).sum() / len(s) * 100, 1) if len(s) else None


# ─────────────────────── STORES MAPPING ───────────────────────────────────────
def load_stores():
    df = pd.read_excel(os.path.join(DATA_DIR, "Stores mapping.xlsx"))
    df = df.rename(columns={
        "Store.StoreId":   "store_id",
        "Store.StoreName": "store_name",
        "Area.AreaName":   "area",
        "Area.AreaManager":"area_manager",
        "Region.Name":     "region",
        "Store.Open":      "open",
    })
    df = df[df["open"] == "Y"].copy()
    df["store_id"]       = df["store_id"].str.strip()
    df["store_name_lc"]  = df["store_name"].str.lower().str.strip()
    return df


# ─────────────────────── OSAT ─────────────────────────────────────────────────
def load_osat(stores):
    want = (["StoreId", "VisitDate", "Overall Satisfaction", "Visit Type"]
            + list(OSAT_CAT_COLS.values()))

    frames = []
    for yr in ("2025", "2026"):
        p = os.path.join(DATA_DIR, "OSAT", f"OSAT {yr}.csv")
        if os.path.exists(p):
            df = pd.read_csv(p, encoding="latin-1",
                             usecols=lambda c: c in want, low_memory=False)
            frames.append(df)
    osat = pd.concat(frames, ignore_index=True)

    osat["date"] = pd.to_datetime(osat["VisitDate"], format="%m/%d/%y %H:%M", errors="coerce")
    bad = osat["date"].isna()
    osat.loc[bad, "date"] = pd.to_datetime(osat.loc[bad, "VisitDate"], errors="coerce")
    osat = osat.dropna(subset=["date"])
    osat["month"] = osat["date"].dt.strftime("%Y-%m")
    osat["week"]  = osat["date"].dt.strftime("%G-W%V")

    osat["sat"] = pd.to_numeric(osat["Overall Satisfaction"], errors="coerce")
    osat = osat.dropna(subset=["sat"])
    osat["sat_ok"] = (osat["sat"] >= 4).astype(int)
    osat["dis_ok"] = (osat["sat"] <= 2).astype(int)

    for col in OSAT_CAT_COLS.values():
        if col in osat.columns:
            osat[col] = pd.to_numeric(osat[col], errors="coerce")

    smap = stores.set_index("store_id")[
        ["store_name", "area", "area_manager", "region"]
    ].to_dict("index")
    osat["store_id"] = osat["StoreId"].str.strip()
    for field in ("store_name", "area", "area_manager", "region"):
        osat[field] = osat["store_id"].map(
            lambda x, f=field: smap.get(x, {}).get(f)
        )
    return osat.dropna(subset=["store_name"])


# ─────────────────────── REVIEWS ──────────────────────────────────────────────
def load_reviews(stores):
    frames = []
    for yr in ("2025", "2026"):
        p = os.path.join(DATA_DIR, "Social Media", f"Reviews {yr}.csv")
        if os.path.exists(p):
            df = pd.read_csv(p, encoding="latin-1",
                             usecols=lambda c: c in [
                                 "Client Location ID", "Rating", "Review Date"
                             ], low_memory=False)
            frames.append(df)
    rev = pd.concat(frames, ignore_index=True)
    # strip BOM from column names
    rev.columns = [c.lstrip("﻿") for c in rev.columns]
    rev = rev.rename(columns={"Client Location ID": "store_id"})

    rev["date"] = pd.to_datetime(rev["Review Date"], format="%m/%d/%Y", errors="coerce")
    bad = rev["date"].isna()
    rev.loc[bad, "date"] = pd.to_datetime(rev.loc[bad, "Review Date"], errors="coerce")
    rev = rev.dropna(subset=["date"])
    rev["rating"] = pd.to_numeric(rev["Rating"], errors="coerce")
    rev = rev.dropna(subset=["rating"])
    rev["month"] = rev["date"].dt.strftime("%Y-%m")
    rev["week"]  = rev["date"].dt.strftime("%G-W%V")

    smap = stores.set_index("store_id")[
        ["store_name", "area", "area_manager", "region"]
    ].to_dict("index")
    rev["store_id"] = rev["store_id"].str.strip()
    for field in ("store_name", "area", "area_manager", "region"):
        rev[field] = rev["store_id"].map(
            lambda x, f=field: smap.get(x, {}).get(f)
        )
    return rev.dropna(subset=["store_name"])


# ─────────────────────── MYSTERY SHOPPER ──────────────────────────────────────
MS_SCORE_COLS = [
    "Arrival Total", "Service Total",
    "Atmosphere Total", "Overall Impressions Total", "Total",
]

def _read_ms(path):
    df = pd.read_excel(path, header=3)
    want = ["Location Name", "Date", "Survey ID"] + MS_SCORE_COLS
    return df[[c for c in want if c in df.columns]].copy()


def load_ms(stores):
    frames = []
    for yr in ("2025", "2026"):
        p = os.path.join(DATA_DIR, "Mystery Shopper", f"Mystery Shopper {yr}.xlsx")
        if os.path.exists(p):
            frames.append(_read_ms(p))
    for p in glob.glob(os.path.join(DATA_DIR, "Mystery Shopper", "Mystery Shopper April*.xlsx")):
        frames.append(_read_ms(p))

    ms = pd.concat(frames, ignore_index=True)
    if "Survey ID" in ms.columns:
        ms = ms.drop_duplicates(subset=["Survey ID"])

    ms["date"] = pd.to_datetime(ms["Date"], errors="coerce")
    ms = ms.dropna(subset=["date"])
    ms["month"] = ms["date"].dt.strftime("%Y-%m")
    ms["week"]  = ms["date"].dt.strftime("%G-W%V")

    for col in MS_SCORE_COLS:
        if col in ms.columns:
            ms[col] = pd.to_numeric(ms[col], errors="coerce")

    ms = ms.dropna(subset=["Total"])

    smap = (stores.drop_duplicates("store_name_lc")
                  .set_index("store_name_lc")[
                      ["store_name", "area", "area_manager", "region"]
                  ].to_dict("index"))
    ms["loc_lc"] = ms["Location Name"].str.lower().str.strip()
    for field in ("store_name", "area", "area_manager", "region"):
        ms[field] = ms["loc_lc"].map(
            lambda x, f=field: smap.get(x, {}).get(f)
        )
    return ms.dropna(subset=["store_name"])


# ─────────────────────── ATEL (CLAIMS) ────────────────────────────────────────
_STATE_PREFIX = re.compile(r"^[A-Za-z]{2}\s*-\s*(.+)$")


def load_atel(stores):
    frames = []
    for yr in ("2025", "2026"):
        p = os.path.join(DATA_DIR, "Atel", f"Atel {yr}.xlsx")
        if os.path.exists(p):
            df = pd.read_excel(p, usecols=lambda c: c in [
                "Hora de inicio",
                "Restaurante que Atendión (Reclamo)",
                "Tipo de Reclamo",
                "Sub categoría",
            ])
            frames.append(df)
    atel = pd.concat(frames, ignore_index=True)

    atel["date"] = pd.to_datetime(atel["Hora de inicio"], errors="coerce")
    atel = atel.dropna(subset=["date"])
    atel["month"] = atel["date"].dt.strftime("%Y-%m")
    atel["week"]  = atel["date"].dt.strftime("%G-W%V")

    def _extract(raw):
        if pd.isna(raw):
            return None
        m = _STATE_PREFIX.match(str(raw).strip())
        return m.group(1).strip() if m else None

    atel["store_part_lc"] = (
        atel["Restaurante que Atendión (Reclamo)"]
        .map(_extract)
        .str.lower()
        .str.strip()
    )

    smap = (stores.drop_duplicates("store_name_lc")
                  .set_index("store_name_lc")[
                      ["store_name", "area", "area_manager", "region"]
                  ].to_dict("index"))
    for field in ("store_name", "area", "area_manager", "region"):
        atel[field] = atel["store_part_lc"].map(
            lambda x, f=field: smap.get(x, {}).get(f) if x else None
        )
    atel["category"] = atel["Tipo de Reclamo"].fillna("Other").str.strip()
    atel["subcat"]   = atel["Sub categoría"].fillna("").str.strip()
    return atel.dropna(subset=["store_name"])


# ─────────────────────── KPI AGGREGATION ──────────────────────────────────────
def _kpi_row(osat_g, rev_g, ms_g, atel_g, key_fields: dict) -> dict:
    responses  = int(osat_g["sat"].count())
    sat_sum    = int(osat_g["sat_ok"].sum())
    dis_sum    = int(osat_g["dis_ok"].sum())
    osat_val   = safe(sat_sum / responses * 100) if responses else None
    dissat_val = safe(dis_sum / responses * 100) if responses else None

    reviews    = int(rev_g["rating"].count())
    rating_sum = float(rev_g["rating"].sum())
    rating     = safe(rating_sum / reviews, 2) if reviews else None

    evaluations = int(ms_g["Total"].count())
    ms_val      = safe(ms_g["Total"].mean()) if evaluations else None

    claims = len(atel_g)

    row = dict(key_fields)
    row.update({
        "sat_sum":    sat_sum,
        "dis_sum":    dis_sum,
        "responses":  responses,
        "osat":       osat_val,
        "dissat":     dissat_val,
        "rating_sum": round(rating_sum, 1),
        "reviews":    reviews,
        "rating":     rating,
        "ms":         ms_val,
        "evaluations":evaluations,
        "claims":     claims,
    })
    return row


def _all_times(dfs, time_key):
    s = set()
    for df in dfs:
        s.update(df[time_key].dropna().unique())
    return sorted(s)


def build_total(osat, rev, ms, atel, time_key):
    rows = []
    for t in _all_times([osat, rev, ms, atel], time_key):
        rows.append(_kpi_row(
            osat[osat[time_key] == t], rev[rev[time_key] == t],
            ms[ms[time_key] == t],     atel[atel[time_key] == t],
            {time_key: t},
        ))
    return rows


def build_by_region(osat, rev, ms, atel, time_key):
    regions = sorted(set(
        list(osat["region"].dropna().unique()) +
        list(rev["region"].dropna().unique()) +
        list(ms["region"].dropna().unique()) +
        list(atel["region"].dropna().unique())
    ))
    rows = []
    for t in _all_times([osat, rev, ms, atel], time_key):
        for rgn in regions:
            rows.append(_kpi_row(
                osat[(osat[time_key] == t) & (osat["region"] == rgn)],
                rev[(rev[time_key] == t)   & (rev["region"] == rgn)],
                ms[(ms[time_key] == t)     & (ms["region"] == rgn)],
                atel[(atel[time_key] == t) & (atel["region"] == rgn)],
                {time_key: t, "region": rgn},
            ))
    return rows


def build_by_area(osat, rev, ms, atel, time_key, stores):
    area_info = (stores[["area", "region", "area_manager"]]
                 .drop_duplicates("area")
                 .set_index("area")
                 .to_dict("index"))
    areas = sorted(set(
        list(osat["area"].dropna().unique()) +
        list(rev["area"].dropna().unique()) +
        list(ms["area"].dropna().unique()) +
        list(atel["area"].dropna().unique())
    ))
    rows = []
    for t in _all_times([osat, rev, ms, atel], time_key):
        for area in areas:
            info = area_info.get(area, {})
            rows.append(_kpi_row(
                osat[(osat[time_key] == t) & (osat["area"] == area)],
                rev[(rev[time_key] == t)   & (rev["area"] == area)],
                ms[(ms[time_key] == t)     & (ms["area"] == area)],
                atel[(atel[time_key] == t) & (atel["area"] == area)],
                {time_key: t, "area": area,
                 "region": info.get("region"),
                 "area_manager": info.get("area_manager")},
            ))
    return rows


def build_by_store(osat, rev, ms, atel, time_key, stores):
    store_info = (stores[["store_name", "area", "area_manager", "region"]]
                  .drop_duplicates("store_name")
                  .set_index("store_name")
                  .to_dict("index"))
    store_names = sorted(set(
        list(osat["store_name"].dropna().unique()) +
        list(rev["store_name"].dropna().unique()) +
        list(ms["store_name"].dropna().unique()) +
        list(atel["store_name"].dropna().unique())
    ))
    rows = []
    for t in _all_times([osat, rev, ms, atel], time_key):
        for sn in store_names:
            info = store_info.get(sn, {})
            row = _kpi_row(
                osat[(osat[time_key] == t) & (osat["store_name"] == sn)],
                rev[(rev[time_key] == t)   & (rev["store_name"] == sn)],
                ms[(ms[time_key] == t)     & (ms["store_name"] == sn)],
                atel[(atel[time_key] == t) & (atel["store_name"] == sn)],
                {time_key: t, "store_name": sn,
                 "area": info.get("area"),
                 "area_manager": info.get("area_manager"),
                 "region": info.get("region")},
            )
            # Skip completely empty rows
            if (row["responses"] + row["reviews"] + row["evaluations"] + row["claims"]) > 0:
                rows.append(row)
    return rows


# ─────────────────────── STORE CATS ───────────────────────────────────────────
def build_store_cats(osat, ms, time_key):
    osat_g = osat.groupby(["store_name", time_key], observed=True, sort=True)
    ms_g   = ms.groupby(["store_name", time_key],   observed=True, sort=True)
    keys   = set(osat_g.groups.keys()) | set(ms_g.groups.keys())

    rows = []
    for (sn, t) in sorted(keys):
        row = {"store_name": sn, time_key: t}

        if (sn, t) in osat_g.groups:
            og = osat_g.get_group((sn, t))
            for key, col in OSAT_CAT_COLS.items():
                row[key] = t2b(og[col]) if col in og.columns else None
        else:
            for key in OSAT_CAT_COLS:
                row[key] = None

        if (sn, t) in ms_g.groups:
            mg = ms_g.get_group((sn, t))
            row["ms_arrival"]    = safe(mg["Arrival Total"].mean())
            row["ms_service"]    = safe(mg["Service Total"].mean())
            row["ms_atmosphere"] = safe(mg["Atmosphere Total"].mean())
            row["ms_overall"]    = safe(mg["Overall Impressions Total"].mean())
        else:
            row["ms_arrival"] = row["ms_service"] = None
            row["ms_atmosphere"] = row["ms_overall"] = None

        rows.append(row)
    return rows


# ─────────────────────── CLAIMS DETAIL ────────────────────────────────────────
def build_store_claims_monthly(atel):
    rows = []
    for (sn, month, cat), grp in atel.groupby(
        ["store_name", "month", "category"], observed=True, sort=True
    ):
        subcats = []
        for sub, sg in grp.groupby("subcat", observed=True):
            if sub:
                subcats.append({"name": sub, "count": len(sg)})
        rows.append({
            "store_name": sn,
            "month":      month,
            "category":   cat,
            "count":      len(grp),
            "subcats":    subcats,
        })
    return rows


# ─────────────────────── META ─────────────────────────────────────────────────
def build_meta(osat, rev, ms, atel, stores):
    months = sorted(set(
        list(osat["month"].unique()) + list(rev["month"].unique()) +
        list(ms["month"].unique())   + list(atel["month"].unique())
    ))
    weeks = sorted(set(
        list(osat["week"].unique()) + list(rev["week"].unique()) +
        list(ms["week"].unique())   + list(atel["week"].unique())
    ))

    # Use last month that has OSAT data (most reliable anchor)
    osat_months = sorted(osat["month"].dropna().unique())
    osat_weeks  = sorted(osat["week"].dropna().unique())
    store_meta = [
        {"store_name": r["store_name"], "area": r["area"],
         "area_manager": r["area_manager"], "region": r["region"]}
        for _, r in stores.iterrows()
    ]
    return {
        "all_months":  months,
        "all_weeks":   weeks,
        "all_regions": sorted(stores["region"].dropna().unique().tolist()),
        "all_areas":   sorted(stores["area"].dropna().unique().tolist()),
        "all_stores":  sorted(stores["store_name"].dropna().unique().tolist()),
        "store_meta":  store_meta,
        "goals":       GOALS,
        "last_month":  osat_months[-1] if osat_months else (months[-1] if months else None),
        "last_week":   osat_weeks[-1]  if osat_weeks  else (weeks[-1]  if weeks  else None),
        "visit_labels":[],
        "time_order":  [],
        "dow_order":   ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"],
    }


# ─────────────────────── MAIN ─────────────────────────────────────────────────
def main():
    print("=" * 60)
    print(f"CX Dashboard update — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)

    print("\n[1/8] Loading stores mapping...")
    stores = load_stores()
    print(f"      {len(stores)} open stores | "
          f"{stores['region'].nunique()} regions | "
          f"{stores['area'].nunique()} areas")

    print("[2/8] Loading OSAT...")
    osat = load_osat(stores)
    print(f"      {len(osat):,} records | {osat['month'].nunique()} months")

    print("[3/8] Loading Reviews...")
    rev = load_reviews(stores)
    print(f"      {len(rev):,} records")

    print("[4/8] Loading Mystery Shopper...")
    ms = load_ms(stores)
    print(f"      {len(ms):,} records")

    print("[5/8] Loading Atel (claims)...")
    atel = load_atel(stores)
    print(f"      {len(atel):,} records")

    print("\n[6/8] Building KPI aggregations...")
    data = {
        "meta":           build_meta(osat, rev, ms, atel, stores),
        "monthly_total":  build_total(osat, rev, ms, atel, "month"),
        "weekly_total":   build_total(osat, rev, ms, atel, "week"),
        "monthly_region": build_by_region(osat, rev, ms, atel, "month"),
        "weekly_region":  build_by_region(osat, rev, ms, atel, "week"),
        "monthly_area":   build_by_area(osat, rev, ms, atel, "month", stores),
        "weekly_area":    build_by_area(osat, rev, ms, atel, "week",  stores),
        "store_monthly":  build_by_store(osat, rev, ms, atel, "month", stores),
    }

    print("[7/8] Building category & detail tables...")
    data.update({
        "store_cats":              build_store_cats(osat, ms, "month"),
        "store_cats_weekly":       build_store_cats(osat, ms, "week"),
        "store_ms_details":        [],
        "store_ms_details_weekly": [],
        "store_claims_monthly":    build_store_claims_monthly(atel),
        "store_claims_resolution": [],
        "store_reviews":           [],
        "store_reviews_monthly":   {},
        "store_channel":           [],
        "store_channel_weekly":    [],
        "store_time":              [],
        "store_time_weekly":       [],
        "store_dow":               [],
        "store_dow_weekly":        [],
        "ms_question_labels":      {},
    })

    print(f"\n[8/8] Writing {OUT_JSON}...")
    with open(OUT_JSON, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, default=str)
    kb = os.path.getsize(OUT_JSON) / 1024
    print(f"      {kb:.0f} KB written")

    mt = data["monthly_total"]
    if mt:
        last = mt[-1]
        print(f"\n  Latest month : {last.get('month')}")
        print(f"  OSAT         : {last.get('osat')}%  ({last.get('responses'):,} responses)")
        print(f"  Rating       : {last.get('rating')}  ({last.get('reviews'):,} reviews)")
        print(f"  MS           : {last.get('ms')}%  ({last.get('evaluations')} evals)")
        print(f"  Claims       : {last.get('claims'):,}")

    # Git push
    print("\nGit push...")
    os.chdir(REPO_DIR)
    subprocess.run(["git", "add", "cx_data.json", "index.html"], check=True)
    today = datetime.now().strftime("%Y-%m-%d")
    result = subprocess.run(
        ["git", "commit", "-m", f"Weekly CX update {today}"],
        capture_output=True, text=True
    )
    if "nothing to commit" in result.stdout or "nothing to commit" in result.stderr:
        print("  Nothing changed — no commit created.")
    else:
        subprocess.run(["git", "push"], check=True)
        print("  Pushed to GitHub Pages.")
    print("\nDone!")


if __name__ == "__main__":
    main()
