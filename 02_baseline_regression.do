version 17.0
set more off

if "$ROOT" == "" {
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
}

capture log close log02
log using "$LOGS/02_baseline_regression.log", name(log02) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
label variable SCCD2 "Spatial coupling coordination degree squared"
xtset id year

global BASE_CONTROLS OPEN UR URG GI

eststo clear
xtreg CE SCCD SCCD2 i.year, fe vce(cluster id)
eststo b1_full
xtreg CE SCCD SCCD2 OPEN i.year, fe vce(cluster id)
eststo b2_full
xtreg CE SCCD SCCD2 OPEN UR i.year, fe vce(cluster id)
eststo b3_full
xtreg CE SCCD SCCD2 OPEN UR URG i.year, fe vce(cluster id)
eststo b4_full
xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo b5_full

esttab b1_full b2_full b3_full b4_full b5_full using "$TABLES/table03_baseline_full_2006_2024.rtf", replace ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OPEN UR URG GI _cons) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Baseline regression: full sample, 2006-2024")

est restore b5_full
scalar tp_full = -_b[SCCD] / (2 * _b[SCCD2])

eststo clear
xtreg CE SCCD SCCD2 i.year if is_fitted == 0, fe vce(cluster id)
eststo b1_obs
xtreg CE SCCD SCCD2 OPEN i.year if is_fitted == 0, fe vce(cluster id)
eststo b2_obs
xtreg CE SCCD SCCD2 OPEN UR i.year if is_fitted == 0, fe vce(cluster id)
eststo b3_obs
xtreg CE SCCD SCCD2 OPEN UR URG i.year if is_fitted == 0, fe vce(cluster id)
eststo b4_obs
xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo b5_obs

esttab b1_obs b2_obs b3_obs b4_obs b5_obs using "$TABLES/table03_baseline_observed_2006_2021.rtf", replace ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OPEN UR URG GI _cons) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Baseline regression: observed-only sample, 2006-2021")

est restore b5_obs
scalar tp_observed = -_b[SCCD] / (2 * _b[SCCD2])

file open tp using "$TABLES/baseline_turning_points.txt", write replace
file write tp "Full-sample turning point, 2006-2024: " %9.6f (tp_full) _n
file write tp "Observed-only turning point, 2006-2021: " %9.6f (tp_observed) _n
file write tp "Formula: -_b[SCCD] / (2 * _b[SCCD2])" _n
file close tp

log close log02
