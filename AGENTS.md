# AGENTS.md

## Project

Empirical manuscript revision for *Sustainable Cities and Society*. Two-way fixed-effects panel analysis with nonlinear SCCD effects, IV-2SLS, robustness, mediation, moderation, and regional heterogeneity.

## Repository Layout

| Path | Purpose |
|------|---------|
| `00_master*.do` … `07_*.do` | Stata do-files (original + `_2006_2024_noER` variants) |
| `插值回归_2006_2024_拟合更新.dta` | Main dataset (284 cities × 19 years = 5,396 obs) |
| `outputs/tables/` | Generated `.rtf`, `.csv`, `.dta` tables |
| `outputs/figures/` | Generated `.png` figures |
| `outputs/logs/` | Stata execution logs |
| `backups/` | Snapshots of previous do-file versions |
| `rendered_*` | Manuscript page renders (PNG contact sheets) |
| `*.docx` | Word manuscript versions (never overwrite directly; use revision scripts) |
| `*_revisions.py`, `apply_*.py` | Python scripts that patch Word manuscripts from text plans |
| `run_stata.ps1` | Auto-detects Stata and runs `00_master.do` in batch mode |

## Data

- **File**: `插值回归_2006_2024_拟合更新.dta`
- **Panel**: `xtset id year` — 284 cities, 2006–2024
- **Observed vs fitted**: 4,544 observed (`is_fitted == 0`), 852 fitted/extrapolated (`is_fitted == 1` for 2022–2024)
- **Key variables**:
  - `CE` → used as `lnCE = CE` (already in log form)
  - `SCCD` / `SCCD2` (core explanatory, `SCCD2 = SCCD^2` generated if absent)
  - `OPEN`, `UR`, `URG`, `GI` (baseline controls)
  - `IV` / `IV2` (lagged SCCD instrument; `IV` has 284 missing values)
  - `DEI`, `POLY`, `OIU`, `GTI`, `MCCD`, `NTL`, `FDI`, `RD`, `EI`, `ISL`, `ISU`
  - `c_SCCD` (centered SCCD, do not recalculate)
  - `is_fitted` (flag for fitted observations)

## Rules

- Do **not** invent data, coefficients, years, p-values, figures, or references.
- Do **not** overwrite the original `.dta` file.
- Do **not** recalculate or overwrite existing `IV` or `c_SCCD`.
- If manuscript text conflicts with the data, flag the issue in `issues_remaining.md` instead of silently fixing it.
- Preserve the empirical logic: two-way fixed effects, nonlinear SCCD effect, robustness tests, IV-2SLS, mediation, moderation, and regional heterogeneity.
- Use `id` as city identifier and `year` as time identifier.
- Use `CE` as dependent variable (log carbon emissions).
- Use `SCCD` and `SCCD2` as core explanatory variables.
- Use `OPEN`, `UR`, `URG`, and `GI` as baseline controls unless the task says otherwise.
- Create new clean do-files only.

## Running the Workflow

### Prerequisites

- Stata 17+ with `estout` and `ivreg2` installed (`ssc install estout, replace`; `ssc install ivreg2, replace`)
- Windows PowerShell (for `run_stata.ps1`)

### Command

```powershell
cd "D:\Workplace\SCS"
powershell -NoProfile -ExecutionPolicy Bypass -File .\run_stata.ps1
```

The script auto-detects Stata (`StataMP-64.exe`, `StataSE-64.exe`, etc.) from registry and common install paths (`D:\Staata18`, `C:\Program Files\Stata18`, etc.).

### Direct Stata (if PowerShell fails)

```stata
cd "D:\Workplace\SCS"
& "D:\Staata18\StataMP-64.exe" /e do "D:\Workplace\SCS\00_master_2006_2024_noER.do"
```

## Workflow Files

| Do-file | Content |
|---------|---------|
| `00_master.do` / `00_master_2006_2024_noER.do` | Orchestrates 01–07; sets globals (`$ROOT`, `$DATA`, `$OUT`, `$TABLES`, `$FIGURES`, `$LOGS`) |
| `01_descriptive_statistics*.do` | Panel diagnostics, missing-value tables, descriptive stats |
| `02_baseline_regression*.do` | TWFE baseline with SCCD + SCCD2, turning-point calculation |
| `03_robustness_endogeneity*.do` | Robustness checks, IV-2SLS first-stage and second-stage |
| `04_heterogeneity*.do` | Regional heterogeneity (eastern/central/western/northeastern) |
| `05_mediation*.do` | Mediation via OIU and GTI (Sobel-Goodman / bootstrap) |
| `06_moderation*.do` | Moderation by DEI, POLY, ER |
| `07_figures*.do` | Kernel density, fitted curves, regional trends |

## Output Conventions

- Tables export as `.rtf` (for Word) and `.csv`/`.dta` (for verification)
- Figures export as `.png` to `outputs/figures/`
- Every module writes a `.log` to `outputs/logs/`
- The master log is checked for `"Completed SCS workflow"` and `r()` errors on exit

## Critical Numbers (2006–2024 Full Sample)

- Full-sample baseline N = 5,366; observed-only N = 4,514
- Baseline turning point = 0.542269; observed-only = 0.521838
- IV turning point = 0.638182; observed-only IV = 0.570248
- Full-control coefficients: SCCD = 8.378***, SCCD2 = –7.725***

## Manuscript Revision

- **Do not** manually edit `.docx` files. Use the Python revision scripts (`apply_SCS_final_revisions.py`, `_revise_manuscript_2006_2024_noER.py`, etc.) that patch the Word document from a structured text plan.
- Current manuscript versions include `revised_manuscript_2006_2024_noER.docx`, `SCS0514_revised_tables_ER_IV_2006_2024_updated.docx`, `scs_submission_ready_v1.docx`.
- See `SCS_COMPREHENSIVE_REVISION_GUIDE.md` for the full 6-phase revision plan and narrative conventions (e.g., replace "mediation effect" with "associated mechanism", "carbon effect" with "emission implications").

## Required Outputs

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