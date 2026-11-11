"""Check that the paper and README agree with the model outputs and data.

Each check recomputes a figure from data/processed/ or outputs/model_out.json,
formats it the way the paper states it, and confirms that exact wording
appears in outputs/Bridge_Pantry_Research_Paper.docx (or README.md).
A few checks test qualitative statements, such as which assumptions matter most.

    python src/check_consistency.py      # exits with status 1 if any check fails
"""
from pathlib import Path
import html
import json
import re
import sys
import zipfile

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
PROC = ROOT / "data" / "processed"
OUT = ROOT / "outputs"

doc_xml = zipfile.ZipFile(OUT / "Bridge_Pantry_Research_Paper.docx").read("word/document.xml").decode("utf-8")
paras = [html.unescape(re.sub(r"<[^>]+>", "", p)) for p in re.findall(r"<w:p[ >].*?</w:p>", doc_xml)]
PAPER = re.sub(r"\s+", " ", " ".join(paras))
README = re.sub(r"\s+", " ", (ROOT / "README.md").read_text())

o = json.load(open(OUT / "model_out.json"))
t, s1, lean, mc = o["truck_base"], o["S1"], o["lean"], o["mc"]
npv3 = o["npv_0.03"]

pit = pd.read_csv(PROC / "chicago_pit_2015_2026.csv").set_index("year")
snap = pd.read_csv(PROC / "snap_illinois_june_2025_2026.csv").set_index("measure")
sect = pd.read_csv(PROC / "refed_illinois_2024_by_sector.csv", index_col=0)
fsub = pd.read_csv(PROC / "refed_illinois_2024_foodservice_subsectors.csv", index_col=0)
lsrc = pd.read_csv(PROC / "refed_illinois_2024_limited_service_categories.csv", index_col=0)
causes = pd.read_csv(PROC / "refed_us_2024_limited_service_causes.csv")
gcfd = pd.read_csv(PROC / "gcfd_benchmarks.csv", index_col=0)
eia = pd.read_csv(PROC / "eia_diesel_monthly.csv")
eia_wk = pd.read_csv(PROC / "eia_diesel_weekly_recent.csv")


def n0(x):
    return f"{round(x):,}"


def r100(x):
    return f"{int(round(x / 100) * 100):,}"


def money(x, d=0):
    return f"${x:,.{d}f}"


def mill(x, d=2):
    return f"${x / 1e6:.{d}f} million"


results = []


def check(section, label, expected, where="paper"):
    text = PAPER if where == "paper" else README
    ok = expected in text
    results.append((ok, section, label, expected, where))


def check_true(section, label, condition):
    results.append((bool(condition), section, label, "(logic check)", "model"))


# ---------------- Chicago homelessness (DFSS one-pager) ----------------
check("2.1", "PIT total", f"{n0(pit.loc[2026, 'total'])} people")
check("2.1", "Sheltered and unsheltered", f"{n0(pit.loc[2026, 'sheltered'])} in sheltered locations and {n0(pit.loc[2026, 'unsheltered'])} unsheltered")
check("2.1", "2024 peak", f"2024 peak of {n0(pit.loc[2024, 'total'])}")
excl = pit.unsheltered + pit.sheltered
check("2.1", "Trend excluding new arrivals",
      f"fell from {n0(excl[2015])} in 2015 to a low of {n0(excl.min())} in {int(excl.idxmin())} and has since climbed back to {n0(excl[2026])}")
check_true("2.1", "2026 count excluding new arrivals is the highest since 2015", excl.loc[2026] == excl.loc[2016:].max() and "its highest level since 2015" in PAPER)
unsh_growth = pit.loc[2026, "unsheltered"] / (pit.loc[2025, "unsheltered"] + pit.loc[2025, "unsheltered_new_arrivals"]) - 1
check_true("2.1", f"Unsheltered rose 'about a quarter' ({unsh_growth:.1%})", 0.20 <= unsh_growth <= 0.30 and "about a quarter" in PAPER)

# ---------------- SNAP (USDA FNA) ----------------
p25, p26 = snap.loc["persons", "june_2025"], snap.loc["persons", "june_2026"]
b26 = snap.loc["benefits_usd", "june_2026"]
check("ES", "SNAP enrollment drop", f"fell {abs(snap.loc['persons', 'pct_change']) * 100:.1f} percent")
check("2.2", "SNAP participants", f"{p25 / 1e6:.2f} million people in June 2025 to {p26 / 1e6:.2f} million in June 2026")
check("2.2", "SNAP drop in people", f"a drop of {n0(p25 - p26)} people")
check("2.2", "SNAP benefits", f"fell {abs(snap.loc['benefits_usd', 'pct_change']) * 100:.0f} percent to ${b26 / 1e6:.0f} million")
per_person = b26 / p26
check("2.2", "Benefit per person", f"about ${per_person:.0f} per person")
check("2.2", "Benefit per meal", f"roughly ${per_person / (30.4 * 3):.2f} per meal")
check("App A", "Benefit per meal derivation", f"${per_person:.2f} \u00f7 (30.4 days \u00d7 3 meals)")

# ---------------- Surplus food (ReFED) ----------------
fs = sect.loc["Foodservice"]
tot = sect.loc["Total"]
lsr = fsub.loc["Limited Service Restaurants"]
check("ES", "Foodservice surplus", f"{n0(fs.tons_surplus)} tons of surplus food in 2024")
check("ES", "Foodservice donations", f"{n0(fs.tons_donations)} tons, or {fs.tons_donations / fs.tons_surplus * 100:.2f} percent")
check("Abstract", "Rounded surplus", f"about {round(fs.tons_surplus, -3):,.0f} tons")
check("2.3", "Illinois total surplus", f"{tot.tons_surplus / 1e6:.2f} million tons")
check("2.3", "Meals equivalent", f"{tot.meals_wasted / 1e9:.1f} billion meals")
check("2.3", "Foodservice surplus value", f"${fs.us_dollars_surplus / 1e9:.1f} billion")
check("2.3", "Share landfilled", f"{fs.tons_landfill / fs.tons_surplus * 100:.1f} percent was landfilled")
check("2.3", "Limited-service surplus", f"{n0(lsr.tons_surplus)} tons and donated {lsr.tons_donations / lsr.tons_surplus * 100:.2f} percent")
check("2.3", "Chicken segment", f"{n0(lsrc.loc['Chicken', 'tons_surplus'])} tons")
cz = causes.groupby("cause_name").share.sum()
check("2.3", "Plate waste share", f"accounts for {cz['Plate Waste'] * 100:.0f} percent of limited-service")
check("2.3", "Overproduction shares", f"account for {cz['Overproduction'] * 100:.0f} and {cz['Catering Overproduction'] * 100:.0f} percent")
check("4", "Pounds per meal (limited service)", f"({2000 * lsrc.tons_surplus.sum() / lsrc.meals_wasted.sum():.2f} pounds per meal)")
check("7.1", "Social cost of carbon", f"${fs.social_cost_of_carbon / fs.surplus_total_100_year_mtco2e_footprint:.0f} per ton")

# ---------------- Food Depository benchmarks ----------------
g = gcfd
check("3.3", "GCFD operating expenses", f"${g.loc['total_operating_expenses', 'fy2025'] / 1e6:.1f} million")
check("3.3", "GCFD donated food", f"${g.loc['contributed_food_distributed', 'fy2025'] / 1e6:.1f} million was the value of donated food")
check("3.3", "GCFD cash expenses", f"${g.loc['cash_operating_expenses', 'fy2025'] / 1e6:.1f} million imply")
check("3.3", "GCFD cash cost per meal",
      f"${g.loc['cash_operating_expenses', 'fy2025'] / g.loc['meals_distributed', 'fy2025']:.2f} per meal")
check("3.3", "Donated share trend",
      f"{g.loc['donated_share_of_food', 'fy2023'] * 100:.0f} percent of supply in fiscal 2023 to {g.loc['donated_share_of_food', 'fy2025'] * 100:.0f} percent")
check("3.3", "Trucking trend",
      f"${g.loc['trucking_freight_storage', 'fy2023'] / 1e6:.1f} million to ${g.loc['trucking_freight_storage', 'fy2025'] / 1e6:.1f} million")
fringe = g.loc["fringe_benefits", "fy2025"] / g.loc["salaries", "fy2025"]
check("App A", "Fringe rate", f"{fringe * 100:.1f}%")

# ---------------- Diesel (EIA) ----------------
eia["yr"] = eia.month.str[-4:]
avg = eia.groupby("yr")[["us", "midwest"]].mean()
check("App A", "Base diesel = U.S. 2026 average", f"${avg.loc['2026', 'us']:.2f}")
check("App A", "Midwest 2026 average", f"Midwest ${avg.loc['2026', 'midwest']:.2f}")
check("App A", "Midwest latest week", f"${eia_wk.midwest.iloc[-1]:.2f}")
check("App A", "Midwest 2025 average", f"${avg.loc['2025', 'midwest']:.2f}")

# ---------------- Stage 1 ----------------
check("5", "Informal sources per week", f"about {n0(s1['informal_meals_week'])} meal-equivalents a week")
check("5", "Meals since January", f"about {r100(s1['since_jan_informal'])} since January 2026")
check("5", "Nando's per week", f"adds about {n0(s1['nandos_meals_week'])} per week")
check("5", "Annual meals and tons", f"about {r100(s1['meals_year'])} meals and {s1['tons_year']:.0f} short tons")
check("5", "Cash cost", f"roughly {money(round(s1['cash'], -2))} a year, or {money(s1['cash_per_meal'], 2)} per meal")
check("5", "Volunteer hours", f"about {r100(s1['vol_hrs'])} hours")

# ---------------- Stage 2 truck ----------------
check("ES", "Meals per truck", f"{n0(t['meals'])} meals a year")
check("ES", "Cash cost per meal", f"{money(t['cost_per_meal_cash'], 2)} per meal ({money(t['cost_per_meal_full'], 2)} including capital)")
check("ES", "Sponsorship coverage", f"covers about {t['self_suff'] * 100:.0f} percent of a truck")
check("Abstract", "Sponsorship coverage", f"covers about {t['self_suff'] * 100:.0f} percent of operating costs")
be = o["breakeven_sponsor"]
be_share = be / o["dedicated_equiv_month"]
check("ES", "Break-even sponsorship", f"roughly {money(round(be, -2))} per month, about {be_share * 100:.0f} percent of the market value")
check("6.2", "Break-even sponsorship", f"about {money(round(be, -2))} a month in sponsorship, roughly {be_share * 100:.0f} percent")
check("6.2", "Dedicated-equivalent value", f"roughly {money(round(o['dedicated_equiv_month'], -3))} a month")
check("6.2", "Base price share", f"about {4000 / o['dedicated_equiv_month'] * 100:.0f} percent of dedicated-equivalent value")
check("6.2", "Contributed coverage", f"covers {t['coverage'] * 100:.0f} percent and leaves an annual gap of about {money(round(-t['net_cash'], -3))}")
check("6.2", "Ten-truck cost", f"about ${10 * t['cash_cost'] / 1e6:.2f} million a year")
check("Tbl 3", "Total cash cost", money(t["cash_cost"]))
check("Tbl 4", "Break-even row", f"Break-even {money(be)} {money(12 * be + t['csr'] + t['donations'])} 100%")

# ---------------- Lean pilot ----------------
check("6.2", "Lean capital", f"about {money(round(lean['capex'], -3))}")
check("6.2", "Lean cash cost", f"about {money(round(lean['cash_cost'], -3))} for an estimated {n0(lean['meals'])} meals, or {money(lean['cost_per_meal_cash'], 2)} per meal")
check("6.2", "Lean coverage", f"covers {lean['coverage'] * 100:.0f} percent")
check("6.2", "Lean NPV and BCR", f"about ${o['lean_npv3']['npv'] / 1e6:.2f} million with a benefit-cost ratio of {o['lean_npv3']['bcr']:.2f}")
check("Tbl 8", "Lean pounds target", f"{n0(round(lean['meals'] * 1.2, -3))} lb")

# ---------------- Cost-benefit results ----------------
check("ES", "NPV at 3%", f"about {mill(npv3['npv'])} at a 3 percent real discount rate")
check("Abstract", "BCR", f"benefit-cost ratio {npv3['bcr']:.2f}")
check("7.2", "PV benefits and costs", f"{mill(npv3['pv_benefits'])} and of costs {mill(npv3['pv_costs'])}")
check_true("ES", "NPV positive at 7%", o["npv_0.07"]["npv"] > 0)
flows = [f["net"] for f in o["flows_3"]]
check_true("7.2", "Net flows negative in years 1-3, positive thereafter", all(x < 0 for x in flows[:3]) and all(x > 0 for x in flows[3:]))
check("ES", "Break-even meal value", f"at least {money(o['breakeven_meal_value'], 2)} to its recipient")
tons = t["tons"]
check("7.2", "Tons diverted", f"about {tons:.0f} short tons")
check("7.2", "Landfill emissions avoided", f"roughly {t['ghg_avoid']:.0f} tons of carbon dioxide equivalent")
check("7.2", "Diesel used and emitted", f"about {int(round(t['gal'] / 50) * 50):,} gallons of diesel, emitting about {t['ghg_truck']:.0f} tons")
with_disp = tons * (0.50 + 2.45 * 0.5) - t["ghg_truck"]
check("7.2", "Net with 50% displacement", f"about {round(with_disp, -1):.0f} tons per truck per year")

# ---------------- Sensitivity and Monte Carlo ----------------
grid = {(x["meals_night"], x["sponsor"]): x["npv3"] for x in o["grid"]}
check_true("7.3", "At 180 meals/night NPV is positive only with high sponsorship",
           grid[(180, 1500)] < 0 and grid[(180, 4000)] < 0 and grid[(180, 8000)] > 0)
check_true("7.3", "At 450 meals/night NPV is positive for all sponsorship levels", all(grid[(450, s)] > 0 for s in (1500, 4000, 8000)))
top2 = [x["param"].split(" (")[0] for x in sorted(o["tornado"], key=lambda d: -d["range"])[:2]]
check_true("7.3", f"Two largest sensitivities are meal value and meals per night ({top2})", top2 == ["Value per meal", "Meals per truck-night"])
check("7.3", "Monte Carlo median", f"median NPV is {mill(mc['p50'])}")
check("7.3", "Monte Carlo mean BCR", f"mean benefit-cost ratio is {mc['bcr_mean']:.2f}")
check("7.3", "Probability positive", f"positive NPV in {mc['prob_pos'] * 100:.0f} percent of draws")
check("Abstract", "Probability positive", f"{mc['prob_pos'] * 100:.0f} percent probability")
check("7.3", "90% interval", f"from \u2212{mill(-mc['p5'])} to {mill(mc['p95'])}")

# ---------------- Statewide and SNAP comparison ----------------
s5 = [x for x in o["statewide"] if abs(x["share"] - 0.05) < 1e-9][0]
check("8.3", "5% scenario tons", f"about {round(s5['tons'], -2):,.0f} tons a year")
check("8.3", "5% scenario meals", f"{s5['meals'] / 1e6:.1f} million meals or about {round(s5['meals_week'], -3):,.0f} meals a week")
check("8.3", "Truck-equivalents", f"about {round(s5['trucks'], -2):.0f} Bridge Pantry trucks")
check_true("8.3", "'Roughly seven times' the 100,000 meals/week target", 6.5 <= s5["meals_week"] / 100000 <= 7.5)
consumed_value = (1 - 0.15) * 2.07
check("9", "Meals per dollar", f"about {1 / t['cost_per_meal_cash']:.2f} to {1 / lean['cost_per_meal_cash']:.2f} meals")
check("9", "Food value per dollar", f"about ${consumed_value / t['cost_per_meal_cash']:.2f} to ${consumed_value / lean['cost_per_meal_cash']:.2f} of food value")

# ---------------- README headline table ----------------
check("README", "Surplus", f"{n0(fs.tons_surplus)} tons, of which {fs.tons_donations / fs.tons_surplus * 100:.2f}% donated", "readme")
check("README", "Stage 1", f"about {r100(s1['meals_year'])} meals per year at {money(s1['cash_per_meal'], 2)} per meal", "readme")
check("README", "Truck", f"{n0(t['meals'])} meals per year at {money(t['cost_per_meal_cash'], 2)} per meal ({money(t['cost_per_meal_full'], 2)} with capital)", "readme")
check("README", "Lean", f"{n0(lean['meals'])} meals per year at {money(lean['cost_per_meal_cash'], 2)} per meal; {lean['coverage'] * 100:.0f}% of costs covered", "readme")
check("README", "Sponsorship", f"{t['self_suff'] * 100:.0f}% of a truck's operating cost; break-even at about {money(round(be, -2))} per month", "readme")
check("README", "NPV", f"${npv3['npv'] / 1e6:.2f}M at 3% real (BCR {npv3['bcr']:.2f}); ${o['npv_0.07']['npv'] / 1e6:.2f}M at 7% (BCR {o['npv_0.07']['bcr']:.2f})", "readme")
check("README", "Monte Carlo", f"{mc['prob_pos'] * 100:.0f}% probability of positive NPV; median ${mc['p50'] / 1e6:.2f}M", "readme")
check("README", "Break-even meal", money(o["breakeven_meal_value"], 2), "readme")

# ---------------- Report ----------------
width = max(len(r[2]) for r in results)
fails = 0
for ok, sec, label, expected, where in results:
    if not ok:
        fails += 1
    print(f"{'PASS' if ok else 'FAIL'}  {sec:<8} {label:<{width}}  {expected if not ok else ''}")
print(f"\n{len(results) - fails} of {len(results)} checks passed.")
sys.exit(1 if fails else 0)
