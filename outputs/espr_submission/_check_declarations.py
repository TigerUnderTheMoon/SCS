"""Check location of Statements & Declarations and ER abbreviation."""
from pathlib import Path
from docx import Document
import re

SRC = Path(r"D:\Workplace\SCS\outputs\espr_submission\manuscript_espr_final.docx")
doc = Document(str(SRC))
paras = doc.paragraphs

print("=== DECLARATIONS LOCATION ===")
decl_keywords = ["funding", "competing interest", "conflict of interest",
                 "author contribution", "data availability", "acknowledgment",
                 "statements and declarations", "declarations"]
for i, p in enumerate(paras):
    t = p.text.lower()
    for kw in decl_keywords:
        if kw in t and len(p.text) < 300:
            style = p.style.name
            print(f"  [{i:03d}] <{style}> {p.text[:150]}")
            break

print()
print("=== ER ABBREVIATION CONTEXT ===")
for i, p in enumerate(paras):
    if "ER" in p.text and ("environmental regulation" in p.text.lower() or "moderation" in p.text.lower()):
        print(f"  [{i:03d}] {p.text[:300]}")
        print()

print()
print("=== DOCUMENT STRUCTURE (all headings) ===")
for i, p in enumerate(paras):
    if p.style.name.startswith("Heading"):
        print(f"  [{i:03d}] <{p.style.name}> {p.text}")
