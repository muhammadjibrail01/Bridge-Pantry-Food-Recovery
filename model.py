"""Financial and cost-benefit model for The Bridge Pantry.

Produces outputs/model_out.json (all figures quoted in the paper) and
outputs/mc_npv.npy (Monte Carlo draws used for Figure 7).

Parameters are listed in P as (base, low, high) where "low" and "high" are
the values that lower and raise net present value, respectively.
"""
from pathlib import Path
import copy
import json

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
PROC = ROOT / "data" / "processed"
OUT = ROOT / "outputs"
OUT.mkdir(parents=True, exist_ok=True)

# ---------------- PARAMETERS (base, low, high) ----------------
P = dict(
    # Stage 2 operations per truck
    nights=312,                          # six nights per week
    meals_night=(300, 180, 450),
    lb_per_meal=1.2,
    secondary_loss=(0.15, 0.30, 0.05),   # share of rescued meals not consumed (EPA: 4-40%)
    route_miles_day=55,
    mpg=10.0,
    reefer_gal_hr=0.6, reefer_hrs=5,
    diesel=(5.00, 6.26, 3.62),           # $/gal (EIA): U.S. Jan-Sep 2026 average; Midwest week of Oct 12, 2026; Midwest 2025 average
    shift_hrs=7,
    wage=(22.0, 26.0, 19.0),             # $/hr, driver-coordinator
    fringe=0.264,                        # GCFD FY2025 fringe benefits / salaries
    insurance=(8000, 12000, 6000),
    maint_per_mile=0.25, reefer_service=1500, led_maint=1500,
    permits=1800,
    commissary=6000,
    supplies_per_meal=(0.20, 0.25, 0.15),
    admin_rate=0.12,                     # share of direct operating cost
    sponsor_fee_rate=0.07,               # fiscal sponsor fee on contributed revenue
    # capital
    truck_cost=(95000, 110000, 80000),
    led_cost=(40000, 55000, 25000),
    wrap_cost=5000,
    life=7, salvage_rate=0.15,
    # revenue
    sponsor_month=(4000, 1500, 8000),
    csr_grants=(20000, 5000, 40000),
    donations=(10000, 5000, 20000),
    # valuation
    meal_value=(2.07, 0.90, 3.70),       # IL SNAP benefit per meal ($189.09 / (30.4 days x 3)); GCFD cash cost per meal; U.S. average meal cost
    landfill_ef=0.50,                    # t CO2e per short ton landfilled (WARM v16, food waste)
    displaced_ef=2.45,                   # t CO2e per short ton, poultry source reduction (WARM v16)
    displacement_share=(0.0, 0.0, 0.5),
    scc=(262, 0, 262),                   # $/t CO2e implied by ReFED 2024 estimates
    diesel_co2=0.01021,                  # t CO2 per gallon of diesel
    vol_hrs_per_shift=7, vol_wage=17.05, # Chicago minimum wage (July 2026) as opportunity cost
)
B = {k: (v[0] if isinstance(v, tuple) else v) for k, v in P.items()}


def truck_year(p, meals_override=None):
    meals = p["meals_night"] * p["nights"] if meals_override is None else meals_override
    miles = p["route_miles_day"] * p["nights"]
    gal = miles / p["mpg"] + p["reefer_gal_hr"] * p["reefer_hrs"] * p["nights"]
    fuel = gal * p["diesel"]
    labor = p["shift_hrs"] * p["nights"] * p["wage"] * (1 + p["fringe"])
    maint = miles * p["maint_per_mile"] + p["reefer_service"] + p["led_maint"]
    supplies = meals * p["supplies_per_meal"]
    direct = fuel + labor + p["insurance"] + maint + p["permits"] + p["commissary"] + supplies
    admin = direct * p["admin_rate"]
    sponsor = p["sponsor_month"] * 12
    contrib = sponsor + p["csr_grants"] + p["donations"]
    fee = contrib * p["sponsor_fee_rate"]
    cash_cost = direct + admin + fee
    capex = p["truck_cost"] + p["led_cost"] + p["wrap_cost"]
    annual_capital = capex * (1 - p["salvage_rate"]) / p["life"]
    tons = meals * p["lb_per_meal"] / 2000
    consumed = meals * (1 - p["secondary_loss"])
    food_val = consumed * p["meal_value"]
    ghg_avoid = tons * (p["landfill_ef"] + p["displaced_ef"] * p["displacement_share"])
    ghg_truck = gal * p["diesel_co2"]
    ghg_net = ghg_avoid - ghg_truck
    ghg_val = ghg_net * p["scc"]
    vol_cost = p["vol_hrs_per_shift"] * p["nights"] * p["vol_wage"]
    return dict(
        meals=meals, miles=miles, gal=gal, fuel=fuel, labor=labor, insurance=p["insurance"],
        maint=maint, permits=p["permits"], commissary=p["commissary"], supplies=supplies,
        direct=direct, admin=admin, fee=fee, cash_cost=cash_cost, capex=capex,
        annual_capital=annual_capital, sponsor=sponsor, csr=p["csr_grants"], donations=p["donations"],
        contrib=contrib, net_cash=contrib - cash_cost, tons=tons, consumed=consumed, food_val=food_val,
        ghg_avoid=ghg_avoid, ghg_truck=ghg_truck, ghg_net=ghg_net, ghg_val=ghg_val, vol_cost=vol_cost,
        cost_per_meal_cash=cash_cost / meals, cost_per_meal_full=(cash_cost + annual_capital) / meals,
        self_suff=sponsor / cash_cost, coverage=contrib / cash_cost,
        # Monthly sponsorship at which contributions exactly cover cash costs. The fiscal sponsor fee is charged on
        # all contributions, so contributions C must satisfy C = direct + admin + fee_rate * C.
        breakeven_sponsor_month=((direct + admin) / (1 - p["sponsor_fee_rate"]) - p["csr_grants"] - p["donations"]) / 12,
    )


# ---------------- Stage 1 baseline ----------------
# Founders' operating reports with standard conversions (8 meals per box, 1.2 lb per meal).
S1 = dict(boxes_day=7, days_week=6, meals_box=8, nandos_locs=8, nandos_meals_loc=10)
s1_meals_week = S1["boxes_day"] * S1["days_week"] * S1["meals_box"] + S1["nandos_locs"] * S1["nandos_meals_loc"]
s1_meals_year = s1_meals_week * 52
s1_miles = 35 * S1["days_week"] * 52
s1_gas = 4.15                                      # Midwest regular gasoline, Oct 12, 2026 (EIA)
s1_fuel = s1_miles / 25 * s1_gas
s1_supplies = s1_meals_year * 0.25
s1_misc = 1000
s1_cash = s1_fuel + s1_supplies + s1_misc
s1_vol_hrs = 2 * 3.5 * S1["days_week"] * 52
S1res = dict(
    meals_week=s1_meals_week, meals_year=s1_meals_year, lbs_year=s1_meals_year * 1.2,
    tons_year=s1_meals_year * 1.2 / 2000, miles=s1_miles, fuel=s1_fuel, supplies=s1_supplies,
    misc=s1_misc, cash=s1_cash, cash_per_meal=s1_cash / s1_meals_year, vol_hrs=s1_vol_hrs,
    informal_meals_week=S1["boxes_day"] * S1["days_week"] * S1["meals_box"],
    nandos_meals_week=S1["nandos_locs"] * S1["nandos_meals_loc"],
    since_jan_informal=S1["boxes_day"] * S1["days_week"] * S1["meals_box"] * 40,
)

# ---------------- Program-level 10-year NPV, incremental to Stage 1 ----------------
FLEET = [1, 2, 3, 3, 3, 3, 3, 3, 3, 3]  # trucks in service, years 1..10
TERMINAL_VALUE_SHARE = 0.45             # below straight-line book value of the replacement trucks (conservative)


def program_npv(p, r):
    ty = truck_year(p)
    pv_b = pv_c = 0
    flows = []
    capex = ty["capex"]
    prev = 0
    s1_ben = S1res["meals_year"] * (1 - p["secondary_loss"]) * p["meal_value"]
    s1_cost = S1res["cash"] + S1res["vol_hrs"] * p["vol_wage"]
    for t, n in enumerate(FLEET, start=1):
        add = max(n - prev, 0)
        cap = add * capex
        if t > p["life"]:  # replace the cohort purchased `life` years earlier
            first = t - p["life"]
            rep = max(FLEET[first - 1] - (FLEET[first - 2] if first >= 2 else 0), 0)
            cap += rep * capex * (1 - p["salvage_rate"])
        ben = n * (ty["food_val"] + ty["sponsor"] + ty["ghg_val"]) - s1_ben
        cost = n * (ty["cash_cost"] + ty["vol_cost"]) - s1_cost + cap
        if t == len(FLEET):  # conservative terminal value: 45% of purchase price per truck still in service
            ben += n * capex * TERMINAL_VALUE_SHARE
        df = (1 + r) ** (-t)
        pv_b += ben * df
        pv_c += cost * df
        flows.append(dict(t=t, trucks=n, benefits=ben, costs=cost, net=ben - cost))
        prev = n
    return dict(pv_benefits=pv_b, pv_costs=pv_c, npv=pv_b - pv_c, bcr=pv_b / pv_c, flows=flows)


def org_cashflow(p):
    ty = truck_year(p)
    rows = []
    prev = 0
    for t, n in enumerate(FLEET, start=1):
        add = max(n - prev, 0)
        prev = n
        rows.append(dict(t=t, trucks=n, revenue=n * ty["contrib"], opex=n * ty["cash_cost"],
                         capex=add * ty["capex"], net=n * ty["contrib"] - n * ty["cash_cost"] - add * ty["capex"]))
    return rows


def scen(overrides):
    p = copy.deepcopy(B)
    p.update(overrides)
    return p


out = {}
base = truck_year(B)
out["truck_base"] = base
out["S1"] = S1res
for r in (0.016, 0.03, 0.07):
    out[f"npv_{r}"] = {k: v for k, v in program_npv(B, r).items() if k != "flows"}
out["flows_3"] = program_npv(B, 0.03)["flows"]
out["org_flows"] = org_cashflow(B)

# Sponsorship scenarios
spons = {}
for name, val in [("low", 1500), ("base", 4000), ("high", 8000)]:
    p = scen(dict(sponsor_month=val))
    ty = truck_year(p)
    res = program_npv(p, 0.03)
    spons[name] = dict(sponsor_month=val, sponsor=ty["sponsor"], contrib=ty["contrib"], cash_cost=ty["cash_cost"],
                       net=ty["net_cash"], self_suff=ty["self_suff"], coverage=ty["coverage"],
                       npv3=res["npv"], bcr3=res["bcr"])
out["spons"] = spons
out["breakeven_sponsor"] = base["breakeven_sponsor_month"]
# Dedicated LED truck benchmark: $1,300/day for a 9-hour day (Boost Outdoor Media 2026), pro-rated to a 7-hour shift, 26 nights
out["dedicated_equiv_month"] = 1300 * 7 / 9 * 26

# Lean configuration: used refrigerated van, volunteer drivers, part-time paid coordinator
lean = scen(dict(truck_cost=55000, led_cost=25000, wage=22.0, shift_hrs=3.5, mpg=14, meals_night=220, insurance=6500))
out["lean"] = truck_year(lean)
out["lean_npv3"] = {k: v for k, v in program_npv(lean, 0.03).items() if k != "flows"}

# Scenario grid: meals per night x sponsorship
grid = []
for mn in (180, 300, 450):
    for sp in (1500, 4000, 8000):
        p = scen(dict(meals_night=mn, sponsor_month=sp))
        ty = truck_year(p)
        grid.append(dict(meals_night=mn, sponsor=sp, cpm=ty["cost_per_meal_cash"], coverage=ty["coverage"],
                         npv3=program_npv(p, 0.03)["npv"]))
out["grid"] = grid

# One-way sensitivity (tornado) on program NPV at 3%
base_npv = program_npv(B, 0.03)["npv"]
torn = []
labels = {
    "meal_value": "Value per meal ($0.90–$3.70)",
    "meals_night": "Meals per truck-night (180–450)",
    "sponsor_month": "Sponsorship per truck ($1,500–$8,000/mo)",
    "wage": "Driver wage ($19–$26/hr)",
    "secondary_loss": "Meals not consumed (5%–30%)",
    "diesel": "Diesel price ($3.62–$6.26/gal)",
    "insurance": "Insurance ($6k–$12k/yr)",
    "supplies_per_meal": "Packaging ($0.15–$0.25/meal)",
}
for k, lab in labels.items():
    lo = program_npv(scen({k: P[k][1]}), 0.03)["npv"]
    hi = program_npv(scen({k: P[k][2]}), 0.03)["npv"]
    torn.append(dict(param=lab, low=lo, high=hi, range=abs(hi - lo)))
cap_lo = program_npv(scen(dict(truck_cost=110000, led_cost=55000)), 0.03)["npv"]
cap_hi = program_npv(scen(dict(truck_cost=80000, led_cost=25000)), 0.03)["npv"]
torn.append(dict(param="Truck + LED capital cost ($105k–$165k)", low=cap_lo, high=cap_hi, range=abs(cap_hi - cap_lo)))
d_lo = program_npv(B, 0.07)["npv"]
d_hi = program_npv(B, 0.016)["npv"]
torn.append(dict(param="Discount rate (7%–1.6%)", low=d_lo, high=d_hi, range=abs(d_hi - d_lo)))
gh = program_npv(scen(dict(displacement_share=0.5)), 0.03)["npv"]
gl = program_npv(scen(dict(scc=0)), 0.03)["npv"]
torn.append(dict(param="Climate accounting (SCC $0 → 50% displaced production)", low=gl, high=gh, range=abs(gh - gl)))
torn.sort(key=lambda d: d["range"])
out["tornado"] = torn
out["base_npv3"] = base_npv

# Monte Carlo (10,000 draws, triangular distributions)
rng = np.random.default_rng(2026)
N = 10000
npvs, bcrs = [], []


def tri(t):
    b, l, h = t
    lo, hi = min(l, h), max(l, h)
    return rng.triangular(lo, b, hi)


for _ in range(N):
    p = copy.deepcopy(B)
    for k in ["meals_night", "secondary_loss", "diesel", "wage", "insurance", "supplies_per_meal",
              "truck_cost", "led_cost", "sponsor_month", "meal_value"]:
        p[k] = tri(P[k])
    p["displacement_share"] = rng.uniform(0, 0.5)
    res = program_npv(p, 0.03)
    npvs.append(res["npv"])
    bcrs.append(res["bcr"])
npvs = np.array(npvs)
bcrs = np.array(bcrs)
out["mc"] = dict(mean=float(npvs.mean()), p5=float(np.percentile(npvs, 5)), p50=float(np.percentile(npvs, 50)),
                 p95=float(np.percentile(npvs, 95)), prob_pos=float((npvs > 0).mean()), bcr_mean=float(bcrs.mean()),
                 bcr_p5=float(np.percentile(bcrs, 5)), bcr_p95=float(np.percentile(bcrs, 95)))
np.save(OUT / "mc_npv.npy", npvs)

# Break-even value per consumed meal (societal NPV = 0 at 3%)
lo, hi = 0.0, 10.0
for _ in range(60):
    mid = (lo + hi) / 2
    if program_npv(scen(dict(meal_value=mid)), 0.03)["npv"] > 0:
        hi = mid
    else:
        lo = mid
out["breakeven_meal_value"] = hi

# Statewide Stage 3 scenarios (ReFED, Illinois foodservice, 2024)
sectors = pd.read_csv(PROC / "refed_illinois_2024_by_sector.csv", index_col=0)
fs_surplus = round(float(sectors.loc["Foodservice", "tons_surplus"]))
fs_donated = round(float(sectors.loc["Foodservice", "tons_donations"]))
meals_per_ton = 2000 / 1.2
st = []
for share in (0.02, 0.05, 0.10):
    tons = fs_surplus * share
    st.append(dict(share=share, tons=tons, meals=tons * meals_per_ton, meals_week=tons * meals_per_ton / 52,
                   trucks=tons * meals_per_ton / base["meals"], ghg_landfill=tons * 0.5))
out["statewide"] = st
out["fs_surplus"] = fs_surplus
out["fs_donated"] = fs_donated

json.dump(out, open(OUT / "model_out.json", "w"), indent=1, default=float)

if __name__ == "__main__":
    print("Stage 1 meals/year:", S1res["meals_year"], " cash per meal: %.2f" % S1res["cash_per_meal"])
    print("Truck: meals %d, cash cost %.0f, cost/meal %.2f, coverage %.0f%%, break-even sponsorship %.0f/mo" % (
        base["meals"], base["cash_cost"], base["cost_per_meal_cash"], 100 * base["coverage"], base["breakeven_sponsor_month"]))
    for r in (0.016, 0.03, 0.07):
        v = out[f"npv_{r}"]
        print("NPV @ %.1f%%: %.0f  BCR %.2f" % (100 * r, v["npv"], v["bcr"]))
    print("Monte Carlo: P(NPV>0) = %.2f, median %.0f" % (out["mc"]["prob_pos"], out["mc"]["p50"]))
    print("Break-even meal value: %.2f" % out["breakeven_meal_value"])
