#!/usr/bin/env bash
# Rebuild everything: processed data -> model -> figures -> paper (Word + PDF).
set -euo pipefail
cd "$(dirname "$0")"

DOCX=outputs/Bridge_Pantry_Research_Paper.docx

echo "== 1. Data preparation (skips any raw file that is not present)"
python3 src/prepare_data.py

echo "== 2. Financial and cost-benefit model"
python3 src/model.py

echo "== 3. Figures"
python3 src/charts.py

echo "== 4. Paper (two passes so the contents page has correct page numbers)"
for pass in 1 2; do
  node src/build_paper.js
  soffice --headless --convert-to pdf --outdir outputs "$DOCX" > /dev/null
  python3 src/toc.py
done
node src/build_paper.js
soffice --headless --convert-to pdf --outdir outputs "$DOCX" > /dev/null

echo "== 5. Consistency check (paper and README against model and data)"
python3 src/check_consistency.py

echo "Done: outputs/Bridge_Pantry_Research_Paper.docx and .pdf"
