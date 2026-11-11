"""Find the page on which each heading appears in the rendered PDF.

Writes outputs/toc_pages.json, which src/build_paper.js uses to print page
numbers on the contents page. Requires poppler-utils (pdfinfo, pdftotext).
"""
from pathlib import Path
import json
import re
import subprocess

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs"
PDF = OUT / "Bridge_Pantry_Research_Paper.pdf"

heads = json.load(open(OUT / "headings.json"))
info = subprocess.run(["pdfinfo", str(PDF)], capture_output=True, text=True).stdout
n_pages = int(re.search(r"Pages:\s+(\d+)", info).group(1))

page_lines = []
for p in range(1, n_pages + 1):
    text = subprocess.run(["pdftotext", "-f", str(p), "-l", str(p), "-layout", str(PDF), "-"],
                          capture_output=True, text=True).stdout
    page_lines.append([re.sub(r"\s+", " ", line).strip() for line in text.splitlines()])

result, start = {}, 4  # skip the title, abstract, and contents pages
for _, heading in heads:
    key = re.sub(r"\s+", " ", heading).strip()
    found = None
    for i in range(start - 1, n_pages):
        if any(line == key or line.startswith(key) for line in page_lines[i]):
            found = i + 1
            break
    result[heading] = found
    if found:
        start = found

json.dump(result, open(OUT / "toc_pages.json", "w"), indent=1)
missing = [h for h, v in result.items() if v is None]
print(f"{n_pages} pages; {len(result) - len(missing)} headings located" + (f"; missing: {missing}" if missing else ""))
