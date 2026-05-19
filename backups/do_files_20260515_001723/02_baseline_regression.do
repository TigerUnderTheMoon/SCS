version 17.0
set more off

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
eststo b1
xtreg CE SCCD SCCD2 OPEN i.year, fe vce(cluster id)
eststo b2
xtreg CE SCCD SCCD2 OPEN UR i.year, fe vce(cluster id)
eststo b3
xtreg CE SCCD SCCD2 OPEN UR URG i.year, fe vce(cluster id)
eststo b4
xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo b5

esttab b1 b2 b3 b4 b5 using "$TABLES/table03_baseline_regression.rtf", replace ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OPEN UR URG GI _cons) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Baseline two-way fixed-effects regression")

esttab b1 b2 b3 b4 b5 using "$TABLES/table03_baseline_regression.csv", replace ///
    b(%9.6f) se(%9.6f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OPEN UR URG GI _cons) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 6))

est restore b5
scalar turning_point = -_b[SCCD] / (2 * _b[SCCD2])
file open tp using "$TABLES/baseline_turning_point.txt", write replace
file write tp "Turning point from baseline model b5: " %9.6f (turning_point) _n
file write tp "Formula: -_b[SCCD] / (2 * _b[SCCD2])" _n
file close tp

log close log02
