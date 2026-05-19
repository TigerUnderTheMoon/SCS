# Revision tasks

## Main task
Based on SCS0514_revised_tables_ER_IV.docx and 插值回归.dta, reconstruct a clean and reproducible Stata workflow.

## Required checks
1. Inspect the dataset structure, variable names, missing values, city count, and year range.
2. Check whether the manuscript's claimed sample period is consistent with the .dta file.
3. Generate SCCD2 if it does not exist:
   gen SCCD2 = SCCD^2
4. Set panel structure:
   xtset id year
5. Reproduce or approximate the following analyses:
   - descriptive statistics
   - baseline two-way fixed-effects regression
   - robustness tests
   - IV-2SLS using lagged SCCD and lagged SCCD2
   - regional heterogeneity
   - mediation analysis using OIU and GTI
   - moderation analysis using ER, DEI, and POLY
6. Export regression tables in a clean format.
7. Write issues_remaining.md listing any mismatch between data, manuscript text, and generated results.

## Important warning
The manuscript currently claims 2006-2024, but the available .dta may not contain 2022-2024. Do not invent missing years. Flag this clearly.