"""Dump Section 2 of revised manuscript for review."""
from docx import Document
doc = Document(r'D:\Workplace\SCS\outputs\espr_submission\manuscript_espr_revised.docx')
in_sec2 = False
in_2x = None
for i, p in enumerate(doc.paragraphs):
    if p.style and p.style.name == "Heading 1":
        if "Literature Review" in p.text:
            in_sec2 = True
            print(f"\n[{i}] # {p.text}")
            continue
        elif in_sec2 and p.text.startswith("3."):
            in_sec2 = False
            break
    if in_sec2:
        st = p.style.name if p.style else ""
        marker = "##" if st == "Heading 2" else "  "
        print(f"[{i}] {marker} {p.text}")