# Run Summary

## Dataset Summary

- Dataset: `插值回归_2006_2024_拟合更新.dta`
- City count: 284
- Year range: 2006-2024
- Total observations: 5,396
- Original observed observations: 4,544
- Fitted/extrapolated observations: 852
- Panel status: balanced, with 19 yearly observations per city
- Duplicate `id year` rows: 0

## Full Sample Vs Observed-Only Sensitivity

The revised workflow is set up to estimate both:

- Full sample: 2006-2024, including `is_fitted == 1` rows for 2022-2024.
- Observed-only sensitivity sample: 2006-2021, using `if is_fitted == 0`.

The workflow was executed from PowerShell with `run_stata.ps1`, using `D:\Staata18\StataMP-64.exe`. All module logs closed normally and no Stata `r()` errors were found.

Descriptive inspection shows the full sample adds 852 fitted rows. Compared with the observed-only sample, full-sample means are higher for CE, SCCD, UR, GI, ER, DEI, OIU, GTI, MCCD, and NTL, while URG is lower. OPEN remains nearly unchanged.

Baseline full-control comparison:

- Full sample, 2006-2024: SCCD = 8.378***, SCCD2 = -7.725***, N = 5,366, within R-squared = 0.641, turning point = 0.542269.
- Observed-only sample, 2006-2021: SCCD = 8.444***, SCCD2 = -8.091***, N = 4,514, within R-squared = 0.650, turning point = 0.521838.

IV comparison:

- Full sample, 2006-2024: IV turning point = 0.638182.
- Observed-only sample, 2006-2021: IV turning point = 0.570248.

The main inverted-U finding is stable in sign and significance, but the added fitted 2022-2024 rows changed sample sizes, coefficients, R-squared values, and turning points.

## Manuscript Sections And Tables To Update

- Abstract: state that the 2006-2024 sample includes fitted/extrapolated 2022-2024 values.
- Section 4.1 Data source and sample: revise the sample description and disclose `is_fitted`.
- Section 4.2 Variables and index construction: clarify treatment of fitted observations and lagged `IV`.
- Table 2: regenerate descriptive statistics for the full 2006-2024 sample and optionally report observed-only sensitivity statistics.
- Table 3: regenerate baseline regression results for full and observed-only samples.
- Table 4: regenerate robustness, endogeneity, and first-stage results for full and observed-only samples.
- Table 5: regenerate regional heterogeneity results for full and observed-only samples.
- Table 6: regenerate mediation results for full and observed-only samples.
- Table 7: regenerate moderation results for full and observed-only samples.
- Results text and conclusion: update coefficients, significance levels, turning points, sample sizes, and any claim that 2022-2024 are observed official data.

## Local Execution Commands

```powershell
cd "D:\Workplace\SCS"
powershell -NoProfile -ExecutionPolicy Bypass -File .\run_stata.ps1
```

Detected executable:

```text
D:\Staata18\StataMP-64.exe
```
