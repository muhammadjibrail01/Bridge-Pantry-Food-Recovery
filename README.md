# The Bridge Pantry: A Financially Sustainable Model for Reducing Food Waste and Homelessness in Chicago

Research paper, financial model, and figures evaluating whether The Bridge Pantry, a Chicago restaurant-to-street food recovery initiative, can scale into a financially sustainable social enterprise, and what Illinois policy would help it do so.

**Author:** Muhammad Jibrail, Harris School of Public Policy, The University of Chicago
**Fiscal sponsor:** Greater Chatham Initiative (501(c)(3))
**Citation style:** The Chicago Manual of Style, author-date system

## Contents

```
.
├── README.md
├── run_all.sh                 # rebuilds everything end to end
├── requirements.txt           # Python dependencies
├── package.json               # Node dependency (docx)
├── src/
│   ├── prepare_data.py        # raw public data -> data/processed/
│   ├── model.py               # Stage 1 baseline, truck budget, 10-year CBA, scenarios, Monte Carlo
│   ├── charts.py              # Figures 1-8
│   ├── build_paper.js         # paper text, tables, figures, references -> Word document
│   ├── toc.py                 # reads page numbers from the PDF for the contents page
│   └── check_consistency.py   # confirms the paper and README match the model and data
├── data/
│   ├── README.md              # sources for every dataset
│   ├── raw/                   # large public downloads (not committed)
│   └── processed/             # small tables used by the model and figures
└── outputs/
    ├── Bridge_Pantry_Research_Paper.docx
    ├── Bridge_Pantry_Research_Paper.pdf
    ├── model_out.json         # all model results used in the paper
    ├── toc_pages.json
    └── figures/               # fig1_stages.png ... fig8_pathway.png
```

## Reproducing the paper

Requirements: Python 3.10+, Node.js 18+, LibreOffice (for PDF conversion), and poppler-utils (`pdfinfo`, `pdftotext`).

```bash
pip install -r requirements.txt
npm install
./run_all.sh
```

`run_all.sh` runs the data preparation, model, figures, and paper build in order, then runs `src/check_consistency.py`, which recomputes the paper's key figures from the data and model and confirms that the paper and this README state them correctly. The committed files in `data/processed/` are sufficient on their own; the data step only rebuilds them if the raw downloads listed in `data/raw/README.md` are present. The paper is built twice so the contents page picks up correct page numbers.

To run steps individually:

```bash
python src/model.py        # prints headline results, writes outputs/model_out.json
python src/charts.py       # writes outputs/figures/
node src/build_paper.js    # writes outputs/Bridge_Pantry_Research_Paper.docx
python src/check_consistency.py   # verifies the paper against the model and data
```

## Headline results

| Measure | Value |
|---|---|
| Illinois foodservice surplus, 2024 | 454,817 tons, of which 0.24% donated |
| Current operation (Stage 1) | about 21,600 meals per year at $0.38 per meal in cash costs |
| Base-case refrigerated truck (Stage 2) | 93,600 meals per year at $1.44 per meal ($1.63 with capital) |
| Lean pilot configuration | 68,640 meals per year at $1.33 per meal; 86% of costs covered |
| Sponsorship coverage at $4,000 per month | 36% of a truck's operating cost; break-even at about $9,100 per month |
| Ten-year societal NPV, three-truck program | $0.45M at 3% real (BCR 1.11); $0.33M at 7% (BCR 1.10) |
| Monte Carlo, 10,000 draws | 78% probability of positive NPV; median $0.93M |
| Break-even value per consumed meal | $1.79 |

## Method in brief

The cost-benefit analysis follows OMB Circular A-4 (2003), reinstated in February 2025, with 3 and 7 percent real discount rates and a 1.6 percent sensitivity case from OMB Circular A-94, Appendix C (2026). It uses a ten-year horizon in constant 2026 dollars, takes the perspective of Chicago and Illinois society, and measures the program incrementally against the current Stage 1 operation.

Benefits are the value of meals consumed (after a 15 percent loss allowance), the value of sponsor advertising services at the price paid, and net greenhouse gas reductions. Grants and donations are treated as transfers. Volunteer time is costed at the Chicago minimum wage. Health and public-service savings are not monetized; the paper reports a break-even meal value instead.

Uncertainty is assessed with one-way sensitivity tests (Figure 6) and a Monte Carlo simulation (Figure 7) using triangular distributions over the ranges in Appendix A of the paper.

## Updating assumptions

All parameters are defined at the top of `src/model.py`, each as `(base, low, high)`.

- **Measured box weights or portions.** Replace the Stage 1 conversion in the `S1` dictionary (`meals_box`, `boxes_day`, `nandos_meals_loc`). The model assumes 8 meals per box and 1.2 lb per meal.
- **Vendor quotes for Stage 2.** Replace `truck_cost`, `led_cost`, `insurance`, and related items. Current values are planning estimates derived from published price ranges (Appendix B of the paper).
- **Fiscal sponsor fee.** Set `sponsor_fee_rate` to the rate in the Greater Chatham Initiative agreement (currently 7%).

After editing, run `./run_all.sh`. All figures, Tables 1 to 5 and 7, and the Appendix C cash-flow table are generated from the model outputs and data files, so they update automatically. Numbers quoted in the paper's prose are written out in `src/build_paper.js`; if a change moves any of them, the consistency check at the end of `run_all.sh` lists each statement that no longer matches so it can be revised.

## Author Contributions

This project was designed, coded, and analyzed independently by Muhammad Jibrail. The code builds the financial model, generates all figures, and assembles the paper. The writing and the policy recommendations (Section 8) represent the author's original intellectual contribution to the field of urban food policy and homelessness.

## Data sources

See `data/README.md` for the source of every dataset and the paper's reference list for full citations. Principal sources: ReFED Insights Engine and U.S. Food Waste Policy Finder; USDA Food and Nutrition Administration SNAP data; U.S. Energy Information Administration fuel prices; Chicago Department of Family and Support Services Point-in-Time count; Chicago Coalition to End Homelessness; Greater Chicago Food Depository reports and filings; U.S. EPA WARM v16 and *From Field to Bin* (2023).
