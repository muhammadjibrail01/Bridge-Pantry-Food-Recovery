"""Build the processed datasets used by the model, figures, and paper.

Raw inputs are large public downloads and are not stored in the repository.
Place them in data/raw/ (see data/README.md for sources and expected file
names) and run:

    python src/prepare_data.py

Each step is skipped, with a message, if its raw file is missing, so the
committed files in data/processed/ remain usable on their own.
"""
from pathlib import Path
import glob
import re
import tempfile
import warnings
import zipfile

import pandas as pd

warnings.filterwarnings("ignore")

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
PROC = ROOT / "data" / "processed"
PROC.mkdir(parents=True, exist_ok=True)

SUM_COLS = [
    "tons_surplus", "tons_supply", "us_dollars_surplus", "tons_donations", "tons_landfill",
    "surplus_total_100_year_mtco2e_footprint", "surplus_upstream_100_year_mtco2e_footprint",
    "surplus_downstream_100_year_mtco2e_footprint", "meals_wasted", "social_cost_of_carbon",
]


def find(pattern):
    hits = sorted(glob.glob(str(RAW / pattern)))
    return Path(hits[0]) if hits else None


def refed_state():
    f = find("*ReFED_US_State_Food_Surplus_Summary*.csv")
    if f is None:
        print("skip: ReFED state summary not found in data/raw/")
        return
    d = pd.read_csv(f, skiprows=3)
    il = d[(d.state == "Illinois") & (d.year == 2024)]

    by_sector = il.groupby("sector")[SUM_COLS].sum()
    by_sector.loc["Total"] = by_sector.sum()
    by_sector.round(3).to_csv(PROC / "refed_illinois_2024_by_sector.csv")

    fs = il[il.sector == "Foodservice"].groupby("sub_sector")[SUM_COLS].sum()
    fs.sort_values("tons_surplus", ascending=False).round(3).to_csv(PROC / "refed_illinois_2024_foodservice_subsectors.csv")

    lsr = il[il.sub_sector == "Limited Service Restaurants"].groupby("sub_sector_category")[SUM_COLS].sum()
    lsr.round(3).to_csv(PROC / "refed_illinois_2024_limited_service_categories.csv")

    trend = d[d.state == "Illinois"].groupby("year")[["tons_surplus", "tons_donations", "tons_landfill"]].sum()
    trend.round(3).to_csv(PROC / "refed_illinois_surplus_by_year.csv")
    print("ok: ReFED Illinois tables")


def refed_causes():
    f = find("*ReFED_US_Food_Surplus_Cause_Summary*.csv")
    if f is None:
        print("skip: ReFED cause summary not found in data/raw/")
        return
    c = pd.read_csv(f, skiprows=3)
    x = c[(c.year == 2024) & (c.sub_sector == "Limited Service Restaurants")]
    x = x.groupby(["cause_group", "cause_name"]).tons_surplus_due_to_cause.sum().reset_index()
    x["share"] = x.tons_surplus_due_to_cause / x.tons_surplus_due_to_cause.sum()
    x.sort_values("share", ascending=False).round(4).to_csv(PROC / "refed_us_2024_limited_service_causes.csv", index=False)
    print("ok: ReFED limited-service causes")


def snap_series():
    f = find("*snap-zip-fy69tocurrent*.zip")
    if f is None:
        print("skip: SNAP state history zip not found in data/raw/")
        return
    rows = []
    with tempfile.TemporaryDirectory() as tmp:
        zipfile.ZipFile(f).extractall(tmp)
        for path in sorted(glob.glob(str(Path(tmp) / "FY*.xls*"))):
            m = re.match(r"FY(\d\d)", Path(path).name)
            if not m:
                continue
            yy = int(m.group(1))
            fy = 2000 + yy if yy < 50 else 1900 + yy
            if fy < 2005:
                continue
            x = pd.ExcelFile(path)
            sheet = [s for s in x.sheet_names if s.upper().startswith("MWRO")][0]
            df = pd.read_excel(x, sheet_name=sheet, header=None)
            col0 = df[0].astype(str).str.strip()
            i = col0[col0 == "Illinois"].index[0]
            blk = df.iloc[i + 1:i + 14]
            blk = blk[pd.to_numeric(blk[2], errors="coerce").notna()]
            blk = blk[~blk[0].astype(str).str.contains("Total", case=False)]
            rows.append((fy, round(pd.to_numeric(blk[2]).mean()), len(blk)))
    s = pd.DataFrame(rows, columns=["FY", "avg_persons", "months"]).sort_values("FY")
    s.to_csv(PROC / "il_snap_series.csv", index=False)
    print("ok: Illinois SNAP series FY%d-FY%d" % (s.FY.min(), s.FY.max()))


def snap_june():
    files = {k: find(f"*snap-{k}*.csv") for k in ("persons", "households", "benefits")}
    if any(v is None for v in files.values()):
        print("skip: SNAP June state tables not found in data/raw/")
        return
    out = {}
    for k, f in files.items():
        t = pd.read_csv(f, skiprows=2)
        row = t[t.iloc[:, 0].astype(str).str.strip() == "Illinois"].iloc[0]
        num = lambda v: float(str(v).replace(",", ""))
        out[k] = (num(row.iloc[1]), num(row.iloc[3]))
    df = pd.DataFrame({
        "measure": ["persons", "households", "benefits_usd"],
        "june_2025": [out["persons"][0], out["households"][0], out["benefits"][0]],
        "june_2026": [out["persons"][1], out["households"][1], out["benefits"][1]],
    })
    df["pct_change"] = (df.june_2026 / df.june_2025 - 1).round(4)
    df.to_csv(PROC / "snap_illinois_june_2025_2026.csv", index=False)
    print("ok: SNAP Illinois June 2025 vs June 2026")


def eia_diesel():
    f = find("*psw18vwall_Data_2*.csv")
    if f is None:
        print("skip: EIA monthly diesel file not found in data/raw/")
        return
    m = pd.read_csv(f, skiprows=2).iloc[:, [0, 1, 6]]
    m.columns = ["month", "us", "midwest"]
    m = m.dropna()
    m.to_csv(PROC / "eia_diesel_monthly.csv", index=False)
    w = find("*psw18vwall_Data_1*.csv")
    if w is not None:
        wk = pd.read_csv(w, skiprows=2).iloc[:, [0, 1, 6]]
        wk.columns = ["week", "us", "midwest"]
        wk.dropna().tail(60).to_csv(PROC / "eia_diesel_weekly_recent.csv", index=False)
    print("ok: EIA diesel prices")


if __name__ == "__main__":
    refed_state()
    refed_causes()
    snap_series()
    snap_june()
    eia_diesel()
