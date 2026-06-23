"""Print exact text of paragraph containing first ER mention."""
from pathlib import Path
from docx import Document

SRC = Path(r"D:\Workplace\SCS\outputs\espr_submission\manuscript_espr_final.docx")
doc = Document(str(SRC))

for i, p in enumerate(doc.paragraphs):
    if "Moderation models examine DEI, ER, and POLY" in p.text:
        print(f"[{i:03d}] <{p.style.name}>")
        print(f"FULL TEXT:")
        print(p.text)
        print()
        print(f"RUNS ({len(p.runs)}):")
        for j, r in enumerate(p.runs):
            print(f"  run[{j}]: '{r.text[:80]}' font={r.font.name} size={r.font.size}")
        break
