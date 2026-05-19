version 17.0
clear all
set more off

* Reproducible workflow for SCS manuscript revision.
* Main sample: 2006-2024, including fitted/extrapolated 2022-2024 rows.
* Sensitivity sample: original observed rows only, is_fitted == 0.
global ROOT "D:/Workplace/SCS"
cd "$ROOT"

global DATA "$ROOT/插值回归_2006_2024_拟合更新.dta"
global OUT "$ROOT/outputs"
global TABLES "$OUT/tables"
global FIGURES "$OUT/figures"
global LOGS "$OUT/logs"

cap mkdir "$OUT"
cap mkdir "$TABLES"
cap mkdir "$FIGURES"
cap mkdir "$LOGS"

capture log close _all
log using "$LOGS/00_master.log", name(master) replace text

di as text "Starting SCS workflow at " c(current_date) " " c(current_time)
di as text "Data file: $DATA"

capture which esttab
if _rc {
    di as text "Installing estout from SSC for table export."
    ssc install estout, replace
}

do "$ROOT/01_descriptive_statistics.do"
do "$ROOT/02_baseline_regression.do"
do "$ROOT/03_robustness_endogeneity.do"
do "$ROOT/04_heterogeneity.do"
do "$ROOT/05_mediation.do"
do "$ROOT/06_moderation.do"
do "$ROOT/07_figures.do"

di as text "Completed SCS workflow at " c(current_date) " " c(current_time)
log close master
