"""Check font sizes, Statements & Declarations structure, and abbreviation definitions."""
from pathlib import Path
from docx import Document
from docx.shared import Pt
import re

SRC = Path(r"D:\Workplace\SCS\outputs\espr_submission\manuscript_espr_final.docx")
doc = Document(str(SRC))
paras = doc.paragraphs

print("=== FONT SIZE CHECK (body paragraphs) ===")
sizes = {}
for i, p in enumerate(paras):
    if p.style.name in ("Normal", "Body Text") and p.text.strip():
        for run in p.runs:
            if run.font.size:
                pt = run.font.size.pt
                sizes[pt] = sizes.get(pt, 0) + 1
print(f"  Font sizes in body text (pt: count): {sizes}")
# Check style default
for sname in ["Normal", "Body Text"]:
    try:
        style = doc.styles[sname]
        font = style.font
        print(f"  Style '{sname}': font={font.name}, size={font.size.pt if font.size else 'inherited'}pt")
    except:
        pass

print()
print("=== STATEMENTS & DECLARATIONS STRUCTURE ===")
# Find references section end
ref_start = -1
for i, p in enumerate(paras):
    if p.style.name == "Heading 1" and "reference" in p.text.lower():
        ref_start = i
        break
if ref_start > 0:
    print(f"  References start at [{ref_start}]")
    print(f"  Content after references:")
    for i in range(ref_start, len(paras)):
        p = paras[i]
        if p.text.strip():
            print(f"    [{i:03d}] <{p.style.name}> {p.text[:120]}")

print()
print("=== ABBREVIATION DEFINITIONS CHECK ===")
# Check if abbreviations are defined at first mention
abbrevs_to_check = {
    "DID": ["difference-in-differences", "difference in differences"],
    "IV": ["instrumental variable", "instrument"],
    "DEI": ["digital economy", "digital economy development index"],
    "OIU": ["optimization of industrial", "industrial upgrading", "industrial structure"],
    "GTI": ["green technolog", "green technology innovation"],
    "POLY": ["polycentric", "polycentricity"],
    "ER": ["environmental regulation"],
    "2SLS": ["two-stage least square", "two-stage least"],
    "TWFE": ["two-way fixed", "two-way fixed effect"],
    "NTL": ["nighttime light"],
    "FDI": ["foreign direct investment"],
}
full_text = "\n".join(p.text for p in paras)
for abbr, expansions in abbrevs_to_check.items():
    # Find first mention of abbr
    idx = full_text.find(abbr)
    if idx < 0:
        print(f"  {abbr}: NOT FOUND in text")
        continue
    # Check if expansion appears before or near first mention
    before = full_text[:idx + 200]
    found_expansion = False
    for exp in expansions:
        if exp.lower() in before.lower():
            found_expansion = True
            break
    status = "DEFINED" if found_expansion else "NOT DEFINED at first mention"
    print(f"  {abbr}: {status}")
    if not found_expansion:
        context = full_text[max(0, idx-100):idx+100]
        print(f"    Context: ...{context}...")

print()
print("=== TITLE CHECK ===")
title = paras[0].text
print(f"  Title: {title}")
print(f"  Contains country/region: {'YES' if 'China' in title or 'Chinese' in title else 'NO'}")
acronyms = re.findall(r"\b[A-Z]{3,}\b", title)
# Filter out common words
acronyms = [a for a in acronyms if a not in ("CO2", "AND", "THE")]
if acronyms:
    print(f"  Acronyms in title: {acronyms} (ESPR says avoid acronyms)")
else:
    print(f"  OK: No problematic acronyms in title")

print()
print("=== CITATION STYLE CHECK (in-text) ===")
# ESPR uses author-date: (Author Year) or Author (Year)
# Check for any numbered citations [1] which would be wrong
numbered = re.findall(r"\[\d+\]", full_text[:ref_start if ref_start > 0 else len(full_text)])
auth_date = re.findall(r"\([A-Z][a-z]+ et al\. \d{4}\)", full_text)
auth_date2 = re.findall(r"\([A-Z][a-z]+ and [A-Z][a-z]+ \d{4}\)", full_text)
auth_date3 = re.findall(r"\([A-Z][a-z]+ \d{4}\)", full_text)
print(f"  Numbered citations [N] in body: {len(numbered)}")
print(f"  Author-date citations found: {len(auth_date) + len(auth_date2) + len(auth_date3)}")
if numbered:
    print(f"  WARNING: ESPR uses author-date style, not numbered citations")
    print(f"    Examples: {numbered[:5]}")
else:
    print(f"  OK: Using author-date citation style")
