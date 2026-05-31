# Issues Remaining

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
