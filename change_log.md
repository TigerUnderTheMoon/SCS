# Change Log

## 2026-06-14

### Submission-Standard Tightening Pass

- Removed remaining manuscript-facing internal process wording from `build_sustainability_observed_manuscript.py`, including `working repository`, `workflow`, `author-side items`, `reviewer-proof`, and `current outputs` language in the generated DOCX/PDF text.
- Added stricter QA fields to `sustainability_deep_checks.json`: `keywords_count`, `title_page_author_metadata_present`, `internal_process_terms_found_main_text`, `author_metadata_blockers`, `submission_ready`, and `submission_status`.
- Added `outputs/sustainability_restructure_20260603_deep/submission_requirements_audit.md` to map MDPI/Sustainability requirements to local evidence and author-only blockers.
- Updated package trackers: `submission_readiness_audit.md`, `final_submission_checklist.md`, `author_metadata_intake.md`, and `submission_cover_letter_draft.md`.
- Replaced the unconfirmed cover-letter conflict statement with author-confirmation placeholders.
- Added a graphical-abstract fallback in the generator so a valid existing `graphical_abstract.png` is retained when the current Python environment lacks `matplotlib`.

### Verification

- Ran `powershell -NoProfile -ExecutionPolicy Bypass -File .\run_stata.ps1`; `outputs/logs/00_master.log` contains `Completed SCS workflow at 14 Jun 2026 01:50:48`, and no `outputs/logs/*.log` entry contains a Stata `r()` error.
- Ran `SUSTAINABILITY_REVISION_DATE=20260603 python build_sustainability_observed_manuscript.py`; `sustainability_deep_checks.json` reports `passed: true`, `submission_ready: false`, `submission_status: STRUCTURAL_QA_PASSED_AUTHOR_METADATA_BLOCKED`, 130 abstract words, 6 keywords, 10,026 main-text words before references, zero long sentences at the audit threshold, and no main-text internal process terms.
- Rendered the final PDF to 32 PNG pages under `outputs/sustainability_restructure_20260603_deep/rendered_pdf_latest/`, rebuilt `contact_sheet.png`, and visually checked the contact sheet plus the title/abstract, back-matter, and reference pages.

### Sustainability Submission Checklist Pass

- Added `08_submission_diagnostics.do` and called it from `00_master.do` to export reviewer-facing VIF diagnostics and fixed-effect joint significance tests for the observed 2006-2021 full-control baseline sample.
- Updated `build_sustainability_observed_manuscript.py` so the generated Sustainability manuscript now uses a structured abstract with `Background`, `Methods`, `Results`, and `Conclusions`.
- Added five submission highlights, a 300 dpi graphical abstract, a categorized supplementary-materials package with `main_analysis/` and `sensitivity_analysis/`, and a supplementary README.
- Added manuscript text for VIF and fixed-effect diagnostics, plus supplementary Tables S1-S2.
- Added an AI-assisted writing declaration and removed manuscript-facing `generator` / `generated workflow` wording.
- Split long manuscript sentences flagged by the sentence-length audit.
- Added `outputs/sustainability_restructure_20260603_deep/reference_doi_audit.csv` for DOI resolver checks.

### Verification

- Ran `powershell -NoProfile -ExecutionPolicy Bypass -File .\run_stata.ps1`; `outputs/logs/00_master.log` contains `Completed SCS workflow`, and no `outputs/logs/*.log` entry contains a Stata `r()` error.
- Confirmed `outputs/tables/table08_vif_observed_2006_2021.csv` and `outputs/tables/table08_fixed_effects_joint_tests_observed_2006_2021.csv` were exported. The maximum VIF is 28.210 for `SCCD`, and city/year fixed effects are jointly significant with p-values below 0.001.
- Ran `SUSTAINABILITY_REVISION_DATE=20260603 python build_sustainability_observed_manuscript.py`; `sustainability_deep_checks.json` reports `passed: true`, structured abstract present, 5 highlights, graphical abstract present, 0 long sentences at the audit threshold, and no manuscript-facing generator trace terms.
- Verified `graphical_abstract.png` is 2400 x 1440 pixels at approximately 300 dpi.
- Rendered the regenerated PDF to 32 PNG pages under `outputs/sustainability_restructure_20260603_deep/rendered_pdf_latest/` with `pdftoppm`, rebuilt `contact_sheet.png`, and visually inspected the contact sheet plus the front page, new diagnostic tables, dense table pages, back matter, and graphical abstract.
- DOI audit status: 27 DOI links resolved directly, 5 DOI links resolved to publisher pages that returned automated-access 403 responses, and 1 reference is a web/data source without a DOI.

## 2026-06-08

### Empirical Workflow Consistency Repair

- Rewrote the moderation regressions in `06_moderation.do` to use explicit `SCCD`, `SCCD2`, moderator main effects, and interaction terms instead of redundant `##` expansions.
- Added the missing `ivreg2` dependency check to `00_master.do` and replaced placeholder weak-instrument diagnostic rows in `03_robustness_endogeneity.do` with observed-sample `ivreg2` Kleibergen-Paap and Cragg-Donald outputs.
- Confirmed by Stata `tabulate region` that the data contain only eastern, central, and western regions; `AGENTS.md` and Sustainability manuscript-generator wording now avoid unsupported northeastern heterogeneity claims.
- Clarified in `AGENTS.md` that `_2006_2024_noER` do-files are historical workflow files and that `run_stata.ps1` calls the current `00_master.do`.
- Retained `outputs/sustainability_restructure_20260603_10k_deep/` as a failed 10,000-word QA attempt rather than deleting it, because its QA JSON is useful provenance.
- Moved `apply_20260524_scs_submission_revisions.py` from the root directory to `backups/`.

### Verification

- Ran `powershell -NoProfile -ExecutionPolicy Bypass -File .\run_stata.ps1`; `outputs/logs/00_master.log` contains `Completed SCS workflow` and `outputs/logs/*.log` has no `r()` errors.
- Confirmed `outputs/tables/table4_iv_diagnostics_observed_2006_2021.csv` contains Kleibergen-Paap rk LM and Kleibergen-Paap rk Wald F rows.
- Confirmed the observed-sample abstract values match the regenerated Stata outputs at displayed precision: SCCD `8.444`, SCCD2 `-8.091`, turning point `0.522`.
- Ran `SUSTAINABILITY_REVISION_DATE=20260603 python build_sustainability_observed_manuscript.py`; `outputs/sustainability_restructure_20260603_deep/sustainability_deep_checks.json` reports `passed: true`.

## 2026-06-03

### Sustainability Literature Integration

- Revised `build_sustainability_observed_manuscript.py` to integrate the four requested smart-city/carbon-emissions studies into the Sustainability draft.
- Corrected the `Sustainability 2023, 15, 225` reference to Ma and Wu's smart-city digitalization paper and added the 2025 277-city spatial-distribution paper and the 2024 green-technology-progress paper.
- Kept the existing Zhu et al. coupling-coordination reference and used it to support the smart performance and low-carbon coordination framing.
- Strengthened the introduction, hypotheses, regional heterogeneity discussion, associated-pathway wording, and policy conclusion around the distinction between average smart-city treatment effects and this paper's nonlinear SCCD intensity result.
- Generated:
  - `outputs/sustainability_restructure_20260603_deep/manuscript_sustainability_deep.docx`
  - `outputs/sustainability_restructure_20260603_deep/manuscript_sustainability_deep.pdf`
  - `outputs/sustainability_restructure_20260603_deep/sustainability_deep_checks.json`
  - `outputs/sustainability_restructure_20260603_deep/revision_evidence_manifest.md`
  - `outputs/sustainability_restructure_20260603_deep/submission_readiness_audit.md`
  - `outputs/sustainability_restructure_20260603_deep/submission_cover_letter_draft.md`
  - `outputs/sustainability_restructure_20260603_deep/author_metadata_intake.md`
  - `outputs/sustainability_restructure_20260603_deep/final_submission_checklist.md`

### Verification

- Ran `SUSTAINABILITY_REVISION_DATE=20260603 python build_sustainability_observed_manuscript.py`.
- Confirmed `sustainability_deep_checks.json` reports `passed: true`, abstract length 153 words, references count 33, H1-H4 present, back-matter headings present, observed-main-sample wording present, and no forbidden `mediation effect` phrase.
- Rendered the final PDF to 18 PNG pages under `outputs/sustainability_restructure_20260603_deep/rendered_pdf_latest/` with `pdftoppm` and inspected the contact sheet plus dense table/reference pages for clipping or overlap.
- Added `submission_readiness_audit.md` to document current MDPI Sustainability submission readiness and the remaining author-only blockers.
- Added `submission_cover_letter_draft.md` as a fill-in cover-letter template for the final Sustainability submission packet.
- Added `author_metadata_intake.md` and `final_submission_checklist.md` to make the remaining author-side submission inputs explicit.

## 2026-05-31

### Sustainability Deep Revision

- Added observed-focused diagnostics to the main workflow:
  - `outputs/tables/table_u_test_observed_2006_2021.csv`
  - `outputs/tables/table_u_test_observed_2006_2021.txt`
  - `outputs/tables/table4_iv_diagnostics_observed_2006_2021.csv`
  - `outputs/tables/table4_iv_diagnostics_observed_2006_2021.txt`
  - `outputs/tables/dei_conditional_turning_points_observed_2006_2021.csv`
  - `outputs/tables/dei_conditional_turning_points_observed_2006_2021.txt`
  - `outputs/tables/open_data_quality_observed_2006_2021.csv`
  - `outputs/tables/open_data_quality_observed_2006_2021.txt`
  - `outputs/tables/sample_position_relative_to_turning_point_observed_2006_2021.csv`
- Updated `02_baseline_regression.do` to export observed-sample U-test evidence and sample-position diagnostics around the observed turning point.
- Updated `03_robustness_endogeneity.do` to export observed-sample IV diagnostics and OPEN data-quality diagnostics.
- Updated `06_moderation.do` to export observed-sample DEI conditional turning-point calculations.
- Rebuilt `build_sustainability_observed_manuscript.py` as the Sustainability deep-revision generator. The generated manuscript now uses observed 2006-2021 as the main sample, moves 2006-2024 fitted/extrapolated evidence to supplementary discussion, adds H1-H4, expands methods/SCCD construction, treats IV as supplementary evidence, frames OIU/GTI as associated mechanisms, and keeps DEI as the supported moderation result.
- Added `SUSTAINABILITY_REVISION_DATE` support to the generator so the output directory can be pinned when rerunning a dated revision package.
- Adjusted Table 1 layout in the generated manuscript so the SCCD indicator-system table renders cleanly.
- Generated:
  - `outputs/sustainability_restructure_20260531_deep/manuscript_sustainability_deep.docx`
  - `outputs/sustainability_restructure_20260531_deep/manuscript_sustainability_deep.pdf`
  - `outputs/sustainability_restructure_20260531_deep/sustainability_deep_checks.json`
  - `outputs/sustainability_restructure_20260531_deep/revision_evidence_manifest.md`
- Updated `issues_remaining.md` with closed items, remaining author actions, and verification notes.

### Verification

- Ran `powershell -NoProfile -ExecutionPolicy Bypass -File .\run_stata.ps1`; the master log contains the completed marker and `outputs/logs/*.log` contains no Stata `r()` errors.
- Confirmed all required observed diagnostic CSV/TXT files exist and are nonempty.
- Confirmed `sustainability_deep_checks.json` reports `passed: true`.
- Attempted `render_docx.py`; it failed because the LibreOffice/converter executable was unavailable in this environment.
- Exported the final PDF through local Word COM and rendered the PDF to 17 PNG pages with `pdftoppm` for visual inspection.

## 2026-05-15

### Files Modified

- `00_master.do`
- `01_descriptive_statistics.do`
- `02_baseline_regression.do`
- `03_robustness_endogeneity.do`
- `04_heterogeneity.do`
- `05_mediation.do`
- `06_moderation.do`
- `07_figures.do`
- `issues_remaining.md`
- `change_log.md`
- `run_summary.md`
- `run_stata.ps1`

### Backup

- Backed up the previous do-files to `D:\Workplace\SCS\backups\do_files_20260515_001723`.

### Data Path Changed

- Old workflow data path: `D:/Workplace/SCS/插值回归.dta`
- New workflow data path: `D:/Workplace/SCS/插值回归_2006_2024_拟合更新.dta`

### Workflow Changes

- Updated the main analysis period to 2006-2024.
- Kept `id` as the city identifier and `year` as the time identifier with `xtset id year`.
- Kept `SCCD2 = SCCD^2` generation only when `SCCD2` is absent.
- Treated existing `IV` as the lagged SCCD instrument and generated only `IV2 = IV^2` for the squared SCCD instrument.
- Did not recalculate or overwrite existing `IV` or `c_SCCD`.
- Added observed-only sensitivity checks using `if is_fitted == 0`.
- Revised figure output names to distinguish full-sample and observed-only outputs where applicable.

### Tables To Be Regenerated By Stata

- `table02_descriptive_statistics_full_2006_2024.rtf`
- `table02_descriptive_statistics_observed_2006_2021.rtf`
- `table03_baseline_full_2006_2024.rtf`
- `table03_baseline_observed_2006_2021.rtf`
- `table04_robustness_endogeneity_full_2006_2024.rtf`
- `table04_robustness_endogeneity_observed_2006_2021.rtf`
- `table04_first_stage_full_2006_2024.rtf`
- `table04_first_stage_observed_2006_2021.rtf`
- `table05_regional_heterogeneity_full_2006_2024.rtf`
- `table05_regional_heterogeneity_observed_2006_2021.rtf`
- `table06_mediation_full_2006_2024.rtf`
- `table06_mediation_observed_2006_2021.rtf`
- `table07_moderation_full_2006_2024.rtf`
- `table07_moderation_observed_2006_2021.rtf`

### Execution Status

- Created `run_stata.ps1` to detect Stata on Windows and run `00_master.do` in batch mode.
- Detected Stata executable: `D:\Staata18\StataMP-64.exe`.
- Executed the workflow through PowerShell. All module logs closed normally and no Stata `r()` errors were found.
- Generated logs for `00_master.do` and modules `01` through `07`.
- Regenerated descriptive, baseline, robustness/endogeneity, heterogeneity, mediation, moderation, and figure outputs under `outputs/`.
- The baseline full-control sample changed from observed-only N = 4,514 to full-sample N = 5,366.
- The baseline full-control coefficients changed from observed-only SCCD = 8.444*** and SCCD2 = -8.091*** to full-sample SCCD = 8.378*** and SCCD2 = -7.725***.
- The baseline turning point changed from 0.521838 in the observed-only sample to 0.542269 in the full sample.

### Manuscript Update Planning

- Inspected `outputs/logs/`, `outputs/tables/`, `outputs/figures/`, `run_summary.md`, `issues_remaining.md`, `change_log.md`, and `SCS0514_revised_tables_ER_IV.docx`.
- Created `manuscript_update_plan.md`.
- Confirmed that the manuscript still contains older table values and narrative claims that must be updated before revision.
- Flagged that full-sample results support the baseline inverted-U conclusion, robustness checks, IV-2SLS nonlinear relationship, regional heterogeneity, and OIU/GTI mechanism evidence.
- Flagged that full-sample moderation results support DEI moderation but do not support the current strong ER or POLY moderation claims.
- No edits were made to the Word manuscript.
