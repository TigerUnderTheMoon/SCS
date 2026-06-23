# Issues Remaining

## 2026-06-14 Submission-Standard Tightening Pass

### Closed Or Addressed In The Current Package

1. Removed remaining manuscript-facing internal process wording such as `working repository`, `workflow`, `author-side items`, `reviewer-proof`, and `current outputs` from the generated DOCX/PDF text.
2. Added a stricter manuscript QA split in `sustainability_deep_checks.json`: structural evidence QA now remains `passed: true`, while final upload readiness is reported separately as `submission_ready: false`.
3. Added `outputs/sustainability_restructure_20260603_deep/submission_requirements_audit.md` to map MDPI/Sustainability requirements to package evidence and author-side blockers.
4. Added QA checks for article type, keyword count, ethics statements, title-page author metadata, and main-text internal process terms.
5. Revised package trackers so `submission_readiness_audit.md`, `final_submission_checklist.md`, `author_metadata_intake.md`, and `submission_cover_letter_draft.md` use the same status: structural QA passed, author metadata blocked.
6. Replaced unconfirmed cover-letter conflict wording with author-confirmation placeholders rather than assuming no conflicts.

### Remaining Author Actions Before Submission

1. Add final title-page author names, affiliations, ORCID IDs if used, and corresponding-author metadata.
2. Finalize the named author list and CRediT contribution statement.
3. Provide the final funding statement with grant numbers, or confirm the formal no-external-funding statement.
4. Provide final acknowledgments text, or confirm that no acknowledgments are required.
5. Provide the final conflict-of-interest declaration.
6. Confirm the data/source-license wording for raw data, processed city-level data, code, tables, and figures.
7. Confirm the AI-use declaration wording if the author team wants journal-specific phrasing.
8. Fill and finalize the cover letter with author metadata, originality/no-under-review declaration, and any reviewer suggestions or exclusions.

### Verification Notes

1. Ran `powershell -NoProfile -ExecutionPolicy Bypass -File .\run_stata.ps1`; `outputs/logs/00_master.log` contains `Completed SCS workflow at 14 Jun 2026 01:50:48`, and no `outputs/logs/*.log` entry contains a Stata `r()` error.
2. Ran `SUSTAINABILITY_REVISION_DATE=20260603 python build_sustainability_observed_manuscript.py`; `sustainability_deep_checks.json` reports `passed: true`, `submission_ready: false`, `submission_status: STRUCTURAL_QA_PASSED_AUTHOR_METADATA_BLOCKED`, 130 abstract words, 6 keywords, 10,026 main-text words before references, 0 long sentences, and no main-text internal process terms.
3. Rendered the final PDF to 32 PNG pages under `outputs/sustainability_restructure_20260603_deep/rendered_pdf_latest/`, rebuilt `contact_sheet.png`, and visually checked the contact sheet plus the title/abstract, author-declaration, data/AI/conflict/supplementary, and reference pages.

## 2026-06-14 Sustainability Submission Checklist Pass

### Closed Or Addressed In The Current Package

1. The manuscript abstract now uses the requested Sustainability structured format: Background, Methods, Results, and Conclusions.
2. Five highlights are prepared in `outputs/sustainability_restructure_20260603_deep/submission_highlights.md`.
3. A graphical abstract is prepared at `outputs/sustainability_restructure_20260603_deep/graphical_abstract.png`; the file is 2400 x 1440 pixels at approximately 300 dpi.
4. Manuscript-facing `generator` / `generated workflow` wording has been removed by the QA gate.
5. VIF diagnostics and fixed-effect joint significance tests are exported to:
   - `outputs/tables/table08_vif_observed_2006_2021.csv`
   - `outputs/tables/table08_fixed_effects_joint_tests_observed_2006_2021.csv`
6. The manuscript now includes supplementary Tables S1-S2 for VIF and fixed-effect diagnostics.
7. The manuscript includes a declaration of generative AI and AI-assisted technologies in the writing process.
8. Supplementary materials are organized under `outputs/sustainability_restructure_20260603_deep/supplementary_materials/` with `main_analysis/`, `sensitivity_analysis/`, and a README.
9. A DOI audit is saved as `outputs/sustainability_restructure_20260603_deep/reference_doi_audit.csv`.
10. The sentence-length QA audit reports zero sentences at or above the 35-word threshold before the reference list.

### Remaining Author Actions Before Submission

1. Finalize the named author list and CRediT contribution statement. The repository still does not contain author-specific roles.
2. Provide the final funding statement with grant numbers, or confirm the formal no-external-funding statement.
3. Provide final acknowledgments text, or confirm that no acknowledgments are required.
4. Provide the final conflict-of-interest declaration.
5. Confirm the data/source-license wording for raw data, processed city-level data, code, tables, and figures.
6. Confirm the AI-use declaration wording if the author team wants journal-specific phrasing.
7. Fill and finalize the cover letter with author metadata, originality/no-under-review declaration, and any reviewer suggestions or exclusions.

### Verification Notes

1. The 2026-06-14 generator run reports `passed: true` in `sustainability_deep_checks.json`.
2. The regenerated PDF was rendered to 32 PNG pages under `outputs/sustainability_restructure_20260603_deep/rendered_pdf_latest/`; the contact sheet and selected dense pages were visually inspected without obvious clipping or overlap.
3. The DOI audit produced 27 direct resolver successes, 5 DOI redirects to publisher pages that returned automated-access 403 responses, and 1 web/data-source reference without a DOI.

## 2026-06-08 Data Consistency Notes

1. Stata `tabulate region` on `插值回归_2006_2024_拟合更新.dta` reports only `东部`, `中部`, and `西部`; no `东北部` category is present.
2. The current heterogeneity workflow should therefore remain limited to eastern, central, and western subsamples. `AGENTS.md` and the Sustainability generator wording were aligned to avoid claiming a northeastern heterogeneity estimate.
3. The `_2006_2024_noER` do-files are retained historical workflow files. `run_stata.ps1` calls `00_master.do`, not `00_master_2006_2024_noER.do`.
4. `outputs/sustainability_restructure_20260603_10k_deep/` is retained as provenance for the prior 10,000-word QA attempt; its `sustainability_deep_checks.json` recorded `passed: false` because `word_count_main_text_no_references` was 9,631.
5. The regenerated `outputs/sustainability_restructure_20260603_deep/sustainability_deep_checks.json` now records `passed: true`, with `word_count_main_text_no_references` equal to 10,000.

## 2026-06-03 Sustainability Literature-Integrated Revision Status

This section records the current implementation state after integrating the four user-specified smart-city and carbon-emissions papers into the Sustainability draft.

### Closed Or Addressed In The Current Draft

1. The Sustainability draft now cites and discusses the requested smart-city digitalization, spatial-distribution, green-technology-progress, and coupling-coordination papers.
2. The `Sustainability 2023, 15, 225` reference has been corrected to Ma and Wu's 353-city smart-city digitalization paper.
3. The introduction now distinguishes prior average smart-city treatment-effect evidence from this paper's nonlinear SCCD intensity evidence.
4. H2, H3, and H4 have been tightened around green technology progress, digital economy capacity, and regional heterogeneity while preserving the current empirical scope.
5. Discussion and conclusions now explain why average policy-adoption carbon reductions can coexist with phase-dependent SCCD effects.
6. The generated manuscript outputs are:
   - `outputs/sustainability_restructure_20260603_deep/manuscript_sustainability_deep.docx`
   - `outputs/sustainability_restructure_20260603_deep/manuscript_sustainability_deep.pdf`
   - `outputs/sustainability_restructure_20260603_deep/sustainability_deep_checks.json`
   - `outputs/sustainability_restructure_20260603_deep/revision_evidence_manifest.md`
   - `outputs/sustainability_restructure_20260603_deep/submission_readiness_audit.md`
   - `outputs/sustainability_restructure_20260603_deep/submission_cover_letter_draft.md`
   - `outputs/sustainability_restructure_20260603_deep/author_metadata_intake.md`
   - `outputs/sustainability_restructure_20260603_deep/final_submission_checklist.md`

### Remaining Author Actions Before Submission

1. Finalize the author CRediT contribution statement against the actual named author list.
2. Provide the final funding statement, including grant numbers, or confirm a formal no-external-funding statement.
3. Provide the final conflict-of-interest declaration.
4. Confirm data-source licenses and the final wording for code/data access.
5. Fill and finalize the MDPI Sustainability cover letter draft with author metadata and any reviewer suggestions or exclusions.
6. Complete the final MDPI/Sustainability submission checklist after author metadata, data-license wording, and cover letter text are finalized.

### Verification Notes

1. The 2026-06-03 generator run reports `passed: true` in `sustainability_deep_checks.json`.
2. The regenerated draft contains 33 references, H1-H4, required Sustainability back-matter headings, observed-main-sample wording, and no forbidden `mediation effect` phrase.
3. The final PDF was rendered to 18 PNG pages under `outputs/sustainability_restructure_20260603_deep/rendered_pdf_latest/`; the contact sheet and dense table/reference pages were inspected without obvious clipping or overlap.
4. `submission_readiness_audit.md` records the package as submission-prepared but blocked on author metadata, data/source-license wording, and cover letter content.
5. `submission_cover_letter_draft.md` now provides a cover-letter template, but the final corresponding author block and declarations still require author confirmation.
6. `author_metadata_intake.md` and `final_submission_checklist.md` convert the remaining blockers into fill-in fields and upload-stage checks.

## 2026-05-31 Sustainability Deep Revision Status

This section records the current implementation state for the Sustainability deep rewrite. Older notes below are retained for provenance.

### Closed Or Addressed In The Current Draft

1. The main manuscript now uses the observed 2006-2021 sample as the primary empirical sample. The 2022-2024 fitted/extrapolated rows are described only as supplementary sensitivity evidence, not as observed records.
2. The workflow now exports observed-sample diagnostics required by the deep revision plan:
   - `outputs/tables/table_u_test_observed_2006_2021.csv`
   - `outputs/tables/table_u_test_observed_2006_2021.txt`
   - `outputs/tables/table4_iv_diagnostics_observed_2006_2021.csv`
   - `outputs/tables/table4_iv_diagnostics_observed_2006_2021.txt`
   - `outputs/tables/dei_conditional_turning_points_observed_2006_2021.csv`
   - `outputs/tables/dei_conditional_turning_points_observed_2006_2021.txt`
   - `outputs/tables/open_data_quality_observed_2006_2021.csv`
   - `outputs/tables/open_data_quality_observed_2006_2021.txt`
   - `outputs/tables/sample_position_relative_to_turning_point_observed_2006_2021.csv`
3. The observed-sample baseline turning point is reported as 0.521838. The observed-sample IV turning point is reported as 0.570248 and treated as supplementary evidence.
4. Mechanism language has been revised to associated mechanisms/pathways rather than formal mediation effects.
5. DEI is treated as the main supported moderation result. ER/POLY are not written as supported mechanisms in the generated Sustainability draft.
6. OPEN data quality is explicitly disclosed for the observed sample: 72 negative nonmissing observed-sample values, equal to 1.590% of nonmissing observed OPEN records.
7. The generated manuscript outputs are:
   - `outputs/sustainability_restructure_20260531_deep/manuscript_sustainability_deep.docx`
   - `outputs/sustainability_restructure_20260531_deep/manuscript_sustainability_deep.pdf`
   - `outputs/sustainability_restructure_20260531_deep/sustainability_deep_checks.json`
   - `outputs/sustainability_restructure_20260531_deep/revision_evidence_manifest.md`

### Remaining Author Actions Before Submission

1. Finalize the author CRediT contribution statement against the actual named author list.
2. Provide the final funding statement, including grant numbers, or confirm a formal no-external-funding statement.
3. Provide the final conflict-of-interest declaration.
4. Confirm data-source licenses and the final wording for code/data access. The current analytical draft uses the agreed strategy: code and generated scripts can be public; raw/processed analytical data remain access-limited where source licenses require it.
5. Manually verify the two newly added Sustainability references and decide whether to retain or replace them.
6. Complete the final MDPI/Sustainability submission checklist after author metadata and data-license wording are finalized.

### Verification Notes

1. `run_stata.ps1` completed successfully on 2026-05-31. `outputs/logs/00_master.log` contains the completed marker and no Stata `r()` error was found in `outputs/logs/*.log`.
2. `sustainability_deep_checks.json` reports `passed: true`, abstract length 153 words, H1-H4 present, required back-matter headings present, no `mediation effect` phrase, and observed-main-sample wording present.
3. `render_docx.py` could not run in this environment because the LibreOffice/converter executable was unavailable. The PDF was exported with local Word COM and rendered to 17 PNG pages with `pdftoppm`; the contact sheet and selected pages were visually inspected without obvious clipping or overlap.

Generated after switching the workflow to `插值回归_2006_2024_拟合更新.dta` on 2026-05-15.

## Data Coverage And Panel Status

1. The new dataset covers 2006-2024. Inspection found 284 city ids, 19 years, and 5,396 total observations.
2. The expanded panel is balanced by `id year`: each city has 19 yearly observations and there are no duplicate `id year` rows.
3. The dataset contains 4,544 original observed rows (`is_fitted == 0`) and 852 fitted/extrapolated rows (`is_fitted == 1`). The fitted rows are exactly the 2022-2024 expansion and must not be described as real observed official data.
4. The manuscript must explicitly state that 2022-2024 values are fitted/extrapolated if the 2006-2024 sample is reported as the main sample.

## Missing Values

Missing counts from the new dataset:

| Variable | Missing |
| --- | ---: |
| OPEN | 15 |
| FDI | 31 |
| URG | 15 |
| IV | 284 |

All other inspected variables have zero missing values. `OPEN` and `URG` still have missing observations.

## Data-Quality Notes

1. The supplied `OPEN` series contains 118 negative observations, with a minimum of -0.7196932, despite the manuscript definition as total imports and exports divided by GDP. The submission text now flags this as a data-quality limitation and retains the values rather than recoding them ex post.

## Regression Consistency Checks

1. Stata was run through `run_stata.ps1` using `D:\Staata18\StataMP-64.exe`. All module logs closed normally and no Stata `r()` errors were found in `outputs/logs`.
2. The do-files now export paired tables for the full 2006-2024 sample and the observed-only 2006-2021 sensitivity sample.
3. The full-control baseline regression has N = 5,366 in the full 2006-2024 sample and N = 4,514 in the observed-only 2006-2021 sample. This differs from the manuscript's earlier full-control baseline N = 4,512.
4. The full-sample baseline turning point is 0.542269, while the observed-only turning point is 0.521838. The full-sample IV turning point is 0.638182, while the observed-only IV turning point is 0.570248.
5. The full-sample full-control baseline coefficients remain an inverted-U pattern: SCCD = 8.378*** and SCCD2 = -7.725***. The observed-only sensitivity results are SCCD = 8.444*** and SCCD2 = -8.091***. The signs and significance are stable, but sample sizes, coefficients, and turning points changed after adding fitted 2022-2024 observations.

## Local Stata Commands To Run

The project now includes `run_stata.ps1`, which detects the installed Stata executable and runs the workflow. Use:

```powershell
cd "D:\Workplace\SCS"
powershell -NoProfile -ExecutionPolicy Bypass -File .\run_stata.ps1
```

The detected executable in this environment is:

```text
D:\Staata18\StataMP-64.exe
```

Direct Stata command pattern:

```stata
cd "D:\Workplace\SCS"
& "D:\Staata18\StataMP-64.exe" /e do "D:\Workplace\SCS\00_master.do"
```

## Manuscript Update Planning Issues

1. `manuscript_update_plan.md` has been created from the generated outputs and the current Word manuscript. The Word manuscript has not been modified.
2. The current manuscript still reports older observed-only numerical values in several places, including the baseline turning point of 0.522 and the earlier table values.
3. If the manuscript uses the full 2006-2024 sample as the main analysis, the baseline turning point should be updated to 0.542269 and the IV turning point should be updated to 0.638182.
4. Full-sample moderation results do not support the current strong ER and POLY moderation claims. In the generated full-sample Table 7, ER interaction terms are not significant and POLY interaction terms are not significant. DEI moderation remains supported.
5. Conditional turning points for ER, DEI, and POLY at meaningful moderator levels were not generated by the current workflow, so conditional turning-point statements remain unresolved unless new outputs are generated.
6. Formal mediation indirect-effect estimates were not generated by the current workflow. The current mediation tables support mechanism consistency, but not separately quantified indirect effects.
