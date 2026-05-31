# Sustainability Deep Revision Evidence Manifest

Generated: 2026-05-31

## Deliverables

- DOCX: `D:\Workplace\SCS\outputs\sustainability_restructure_20260531_deep\manuscript_sustainability_deep.docx`
- PDF: `D:\Workplace\SCS\outputs\sustainability_restructure_20260531_deep\manuscript_sustainability_deep.pdf`
- QA JSON: `D:\Workplace\SCS\outputs\sustainability_restructure_20260531_deep\sustainability_deep_checks.json`

## Evidence Inputs

- Observed 2006-2021 baseline, robustness, heterogeneity, mechanism, and moderation tables under `outputs/tables/`.
- Observed-sample U-test, IV diagnostics, DEI conditional turning points, OPEN data-quality diagnostic, and turning-point distribution outputs.
- Figures under `outputs/figures/`, including the observed-sample SCCD distribution, SCCD-lnCE fit, and DEI moderation curves.

## Resolved Issues

- Main text uses observed 2006-2021 rows as the empirical sample.
- Fitted or extrapolated 2022-2024 rows are not described as observed records.
- Mechanism results are framed as associated pathways rather than decomposed indirect effects.
- IV results are downgraded to supplementary evidence and accompanied by first-stage and weak-instrument diagnostics.
- OPEN negative values are disclosed and retained rather than silently recoded.

## Remaining Author Actions

- Supply final author CRediT contribution statement.
- Supply final funding statement.
- Supply final conflict-of-interest declaration.
- Confirm data-source licenses and final repository/access wording.
- Review the newly added Sustainability references and replace any that the author team does not want to cite.

## Structural QA

- Passed: `True`
- PDF created: `True`
- DOCX renderer note: if `render_docx.py` cannot locate LibreOffice/soffice in this environment, export the DOCX through local Word COM and render the PDF with `pdftoppm` for visual QA.
- Current run: `render_docx.py` was attempted, but the LibreOffice/converter executable was unavailable.
- Visual fallback: final PDF was rendered to 17 PNG pages with `pdftoppm` and inspected through `outputs/sustainability_restructure_20260531_deep/rendered_pdf/contact_sheet.png`.
