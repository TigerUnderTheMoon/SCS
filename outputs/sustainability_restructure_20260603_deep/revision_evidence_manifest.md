# Sustainability Deep Revision Evidence Manifest

Generated: 2026-06-14
Package stamp: 2026-06-03

## Deliverables

- DOCX: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\manuscript_sustainability_deep.docx`
- PDF: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\manuscript_sustainability_deep.pdf`
- QA JSON: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\sustainability_deep_checks.json`
- Highlights: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\submission_highlights.md`
- Graphical abstract: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\graphical_abstract.png`
- Supplementary README: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\supplementary_materials\README.md`
- Reference DOI audit: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\reference_doi_audit.csv`
- Submission requirements audit: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\submission_requirements_audit.md`

## Evidence Inputs

- Observed 2006-2021 baseline, robustness, heterogeneity, mechanism, and moderation tables under `outputs/tables/`.
- Observed-sample U-test, IV diagnostics, DEI conditional turning points, OPEN data-quality diagnostic, and turning-point distribution outputs.
- Observed-sample VIF diagnostics and fixed-effect joint significance tests.
- Figures under `outputs/figures/`, including the observed-sample SCCD distribution, SCCD-lnCE fit, and DEI moderation curves.

## Resolved Issues

- Main text uses observed 2006-2021 rows as the empirical sample.
- Fitted or extrapolated 2022-2024 rows are not described as observed records.
- Mechanism results are framed as associated pathways rather than decomposed indirect effects.
- IV results are downgraded to supplementary evidence and accompanied by first-stage and weak-instrument diagnostics.
- VIF and fixed-effect joint-test diagnostics are included for reviewer-facing baseline validation.
- OPEN negative values are disclosed and retained rather than silently recoded.

## Remaining Author Actions

- Supply final title-page author metadata and CRediT contribution statement.
- Supply final funding statement.
- Supply final acknowledgment and conflict-of-interest declarations.
- Confirm data-source licenses and final repository/access wording.
- Confirm that the author team wants to retain the final literature framing and citation choices.

## Structural QA

- Passed: `True`
- Submission ready: `False`
- Submission status: `STRUCTURAL_QA_PASSED_AUTHOR_METADATA_BLOCKED`
- PDF created: `True`
- DOCX renderer note: if `render_docx.py` cannot locate LibreOffice/soffice in this environment, export the DOCX through local Word COM and render the PDF with `pdftoppm` for visual QA.
