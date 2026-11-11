"""Generate Figures 1-8 for the paper into outputs/figures/.

Requires outputs/model_out.json and outputs/mc_npv.npy from src/model.py.
"""
from pathlib import Path
import json

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
PROC = ROOT / "data" / "processed"
OUT = ROOT / "outputs"
FIG = OUT / "figures"
FIG.mkdir(parents=True, exist_ok=True)

plt.rcParams.update({"font.family": "DejaVu Serif", "font.size": 10, "axes.spines.top": False,
                     "axes.spines.right": False, "axes.titleweight": "bold", "axes.titlesize": 11, "figure.dpi": 200})
NAVY = "#1f3a5f"; GREEN = "#2e6b4f"; RED = "#a23b3b"; GREY = "#9aa5b1"; LIGHT = "#d9e2ec"; GOLD = "#c8962e"
out = json.load(open(OUT / "model_out.json"))
gcfd = pd.read_csv(PROC / "gcfd_benchmarks.csv", index_col=0)

# Figure 1: four-stage model
fig, ax = plt.subplots(figsize=(7.5, 2.6)); ax.axis("off"); ax.set_xlim(0, 100); ax.set_ylim(0, 30)
stages = [("Stage 1", "Volunteer food rescue\n(Jan 2026–present)", "Informal QSR sources +\n8 Nando\u2019s locations;\npersonal vehicle", GREY),
          ("Stage 2", "Sponsored mobile\nimpact trucks", "Refrigerated truck,\nLED sponsor panels,\nnightly routes", NAVY),
          ("Stage 3", "Illinois policy\nframework", "Tiered edible-food\nrecovery; donation\nlaw alignment", GREEN),
          ("Stage 4", "Reintegration centers\nand social caf\u00e9s", "Services, training,\nemployment, housing\nnavigation", GOLD)]
for i, (s, t, d, c) in enumerate(stages):
    x = 2 + i * 24.5
    ax.add_patch(FancyBboxPatch((x, 3), 21, 24, boxstyle="round,pad=0.4,rounding_size=1.5", fc="white", ec=c, lw=2))
    ax.add_patch(FancyBboxPatch((x, 21), 21, 6, boxstyle="round,pad=0.4,rounding_size=1.5", fc=c, ec=c, lw=2))
    ax.text(x + 10.5, 24, s, ha="center", va="center", color="white", fontweight="bold", fontsize=10)
    ax.text(x + 10.5, 16.3, t, ha="center", va="center", fontsize=8.6, fontweight="bold", color="#222")
    ax.text(x + 10.5, 8.2, d, ha="center", va="center", fontsize=7.6, color="#333")
    if i < 3:
        ax.add_patch(FancyArrowPatch((x + 21.6, 15), (x + 24.2, 15), arrowstyle="-|>", mutation_scale=12, color="#555"))
plt.tight_layout(); plt.savefig(FIG / "fig1_stages.png", dpi=220, bbox_inches="tight"); plt.close()

# Figure 2: Chicago Point-in-Time count
pit = pd.read_csv(PROC / "chicago_pit_2015_2026.csv")
fig, ax = plt.subplots(figsize=(7.2, 3.3))
x = np.arange(len(pit))
ax.bar(x, pit.unsheltered, color=RED, label="Unsheltered")
ax.bar(x, pit.unsheltered_new_arrivals, bottom=pit.unsheltered, color="#d98c8c", label="Unsheltered new arrivals/newcomers")
b2 = pit.unsheltered + pit.unsheltered_new_arrivals
ax.bar(x, pit.sheltered, bottom=b2, color=NAVY, label="Sheltered")
ax.bar(x, pit.sheltered_new_arrivals, bottom=b2 + pit.sheltered, color="#8fa8c8", label="Sheltered new arrivals/newcomers")
for i, t in enumerate(pit.total):
    ax.text(i, t + 250, f"{t:,}", ha="center", fontsize=7.5)
ax.set_xticks(x); ax.set_xticklabels(pit.year, fontsize=8); ax.set_ylabel("People counted on one January night")
ax.set_ylim(0, 20500); ax.legend(fontsize=7.5, frameon=False, loc="upper left")
plt.tight_layout(); plt.savefig(FIG / "fig2_pit.png", dpi=220); plt.close()

# Figure 3: Illinois SNAP participation
s = pd.read_csv(PROC / "il_snap_series.csv")
june = pd.read_csv(PROC / "snap_illinois_june_2025_2026.csv", index_col=0)
j26 = june.loc["persons", "june_2026"] / 1e6
jchg = june.loc["persons", "pct_change"] * 100
fy26 = s[s.FY == s.FY.max()].avg_persons.iloc[0] / 1e6
fig, ax = plt.subplots(figsize=(7.2, 3.0))
ax.plot(s.FY, s.avg_persons / 1e6, color=NAVY, lw=2, marker="o", ms=3.5)
ax.annotate(f"FY2026 (Oct–Jun)\n{fy26:.2f} million", xy=(2026, fy26), xytext=(2021.3, 1.30), fontsize=8,
            arrowprops=dict(arrowstyle="->", color="#555"))
ax.annotate(f"June 2026: {j26:.2f} million\n(\u2212{abs(jchg):.1f}% vs. June 2025)", xy=(2026, j26), xytext=(2014.2, 1.30),
            fontsize=8, color=RED, arrowprops=dict(arrowstyle="->", color=RED))
ax.scatter([2026], [j26], color=RED, zorder=5, s=18)
ax.set_ylabel("Average monthly participants (millions)"); ax.set_ylim(1.0, 2.2)
plt.tight_layout(); plt.savefig(FIG / "fig3_snap.png", dpi=220); plt.close()

# Figure 4: Illinois foodservice surplus by segment
fs = pd.read_csv(PROC / "refed_illinois_2024_foodservice_subsectors.csv", index_col=0)
il = fs[["tons_surplus", "tons_donations"]].sort_values("tons_surplus").tail(8)
fig, ax = plt.subplots(figsize=(7.2, 3.2))
ax.barh(il.index, il.tons_surplus / 1000, color=GREY, label="Surplus (thousand tons)")
ax.barh(il.index, il.tons_donations / 1000, color=GREEN, label="Donated")
for i, (t, dn) in enumerate(zip(il.tons_surplus, il.tons_donations)):
    ax.text(t / 1000 + 2, i, f"{t/1000:,.1f}k tons; {dn/t*100:.2f}% donated", va="center", fontsize=7.5)
ax.set_xlabel("Thousand tons of surplus food, 2024"); ax.set_xlim(0, 260)
ax.legend(frameon=False, fontsize=8, loc="lower right"); ax.tick_params(axis="y", labelsize=8)
plt.tight_layout(); plt.savefig(FIG / "fig4_refed.png", dpi=220); plt.close()

# Figure 5: cost per meal (horizontal bars)
tb = out["truck_base"]; S1 = out["S1"]; lean = out["lean"]
gcfd_cpm = gcfd.loc["cash_operating_expenses", "fy2025"] / gcfd.loc["meals_distributed", "fy2025"]
labels = ["Stage 1 volunteer operation (cash only)", "Stage 2 lean pilot (cash)", "Stage 2 base truck (cash)",
          "Stage 2 base truck (incl. capital)", "GCFD cash cost per meal, FY2025",
          "Illinois SNAP benefit per meal, June 2026", "U.S. average meal cost, 2024"]
vals = [S1["cash_per_meal"], lean["cost_per_meal_cash"], tb["cost_per_meal_cash"], tb["cost_per_meal_full"], gcfd_cpm, 2.07, 3.70]
cols = [GREY, GREEN, NAVY, NAVY, "#6b7c8f", "#6b7c8f", "#6b7c8f"]
fig, ax = plt.subplots(figsize=(7.2, 3.1))
y = np.arange(len(vals))[::-1]
ax.barh(y, vals, color=cols, height=0.62)
for yi, v in zip(y, vals):
    ax.text(v + 0.05, yi, f"${v:.2f}", va="center", fontsize=8.5)
ax.set_yticks(y); ax.set_yticklabels(labels, fontsize=8.5); ax.set_xlabel("Dollars per meal"); ax.set_xlim(0, 4.3)
plt.tight_layout(); plt.savefig(FIG / "fig5_cpm.png", dpi=220); plt.close()

# Figure 6: tornado sensitivity
t = out["tornado"]; base = out["base_npv3"]
fig, ax = plt.subplots(figsize=(7.2, 3.9))
for i, r in enumerate(t):
    lo, hi = min(r["low"], r["high"]), max(r["low"], r["high"])
    ax.barh(i, (hi - lo) / 1e6, left=lo / 1e6, color=NAVY if hi > 0 else RED, alpha=0.85)
ax.axvline(base / 1e6, color="#333", lw=1); ax.axvline(0, color=RED, lw=0.8, ls="--")
ax.set_yticks(range(len(t))); ax.set_yticklabels([r["param"].replace("$", r"\$") for r in t], fontsize=7.5)
ax.set_xlabel("10-year societal NPV at 3% real (millions of 2026 dollars)")
ax.text(base / 1e6 + 0.05, len(t) - 0.5, r"Base case \$" + f"{base/1e6:.2f}M", fontsize=7.5)
plt.tight_layout(); plt.savefig(FIG / "fig6_tornado.png", dpi=220); plt.close()

# Figure 7: Monte Carlo distribution
npv = np.load(OUT / "mc_npv.npy"); mc = out["mc"]
fig, ax = plt.subplots(figsize=(7.2, 2.9))
ax.hist(npv / 1e6, bins=60, color=NAVY, alpha=0.85)
ax.axvline(0, color=RED, ls="--", lw=1)
ax.axvline(mc["p50"] / 1e6, color=GOLD, lw=1.5)
ax.text(mc["p50"] / 1e6 + 0.08, ax.get_ylim()[1] * 0.85,
        r"Median \$" + f"{mc['p50']/1e6:.2f}M\nP(NPV>0) = {mc['prob_pos']*100:.0f}%", fontsize=8)
ax.set_xlabel("10-year societal NPV at 3% real (millions of 2026 dollars)"); ax.set_ylabel("Simulations")
plt.tight_layout(); plt.savefig(FIG / "fig7_mc.png", dpi=220); plt.close()

# Figure 8: Stage 4 pathway
fig, ax = plt.subplots(figsize=(7.5, 1.9)); ax.axis("off"); ax.set_xlim(0, 105); ax.set_ylim(0, 12)
steps = ["Nightly meal\n(engagement)", "Center visit:\nshower, laundry,\nID, benefits", "Case management\n& counseling",
         "Paid caf\u00e9\ntraining (13–26 wks)", "Placement in\npartner restaurant", "Stable housing\n(Housing First\nreferral)"]
for i, sname in enumerate(steps):
    x = 1 + i * 17.3
    ax.add_patch(FancyBboxPatch((x, 1.5), 14.5, 9, boxstyle="round,pad=0.3,rounding_size=1.2",
                                fc=LIGHT if i % 2 == 0 else "white", ec=NAVY, lw=1.4))
    ax.text(x + 7.25, 6, sname, ha="center", va="center", fontsize=7.4)
    if i < 5:
        ax.add_patch(FancyArrowPatch((x + 14.9, 6), (x + 17.1, 6), arrowstyle="-|>", mutation_scale=10, color="#555"))
plt.tight_layout(); plt.savefig(FIG / "fig8_pathway.png", dpi=220, bbox_inches="tight"); plt.close()

print("figures written to", FIG.relative_to(ROOT))
