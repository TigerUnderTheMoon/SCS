version 17.0
clear all
set more off

* Reproducible workflow for the 2006-2024 no-ER manuscript revision.
* Main sample: 284 prefecture-level and above cities, 2006-2024.
* The raw file stores the log carbon-emissions measure as CE; this workflow
* creates lnCE = CE and uses lnCE as the regression dependent variable.
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
log using "$LOGS/master_2006_2024_noER.log", name(master_noer) replace text

di as text "Starting SCS 2006-2024 noER workflow at " c(current_date) " " c(current_time)
di as text "Data file: $DATA"

local missing_cmds
foreach cmd in esttab eststo estpost estadd ivreg2 {
    capture which `cmd'
    if _rc local missing_cmds "`missing_cmds' `cmd'"
}
if trim("`missing_cmds'") != "" {
    di as error "Required Stata commands are missing:`missing_cmds'"
    di as error "Install commands, if approved by the user:"
    di as error "ssc install estout, replace"
    di as error "ssc install ivreg2, replace"
    exit 199
}

do "$ROOT/01_descriptive_statistics_2006_2024_noER.do"
do "$ROOT/02_baseline_regression_2006_2024_noER.do"
do "$ROOT/03_robustness_endogeneity_2006_2024_noER.do"
do "$ROOT/04_heterogeneity_2006_2024_noER.do"
do "$ROOT/05_mediation_2006_2024_noER.do"
do "$ROOT/06_moderation_2006_2024_noER.do"
do "$ROOT/07_figures_2006_2024_noER.do"

di as text "Completed SCS 2006-2024 noER workflow at " c(current_date) " " c(current_time)
log close master_noer
