version 17.0
set more off

capture log close log03
log using "$LOGS/03_robustness_endogeneity.log", name(log03) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
capture confirm variable MCCD2
if _rc gen double MCCD2 = MCCD^2
xtset id year

global BASE_CONTROLS OPEN UR URG GI

eststo clear

* Robustness 1: bilateral winsorization at 1%.
preserve
foreach v in CE SCCD SCCD2 OPEN UR URG GI {
    gen double w_`v' = `v'
    quietly _pctile `v' if !missing(`v'), p(1 99)
    replace w_`v' = r(r1) if w_`v' < r(r1) & !missing(w_`v')
    replace w_`v' = r(r2) if w_`v' > r(r2) & !missing(w_`v')
}
xtreg w_CE w_SCCD w_SCCD2 w_OPEN w_UR w_URG w_GI i.year, fe vce(cluster id)
eststo r1_winsor
restore

* Robustness 2: alternative dependent variable, nighttime light intensity.
xtreg NTL SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo r2_ntl

* Robustness 3: alternative explanatory variable, MCCD.
xtreg CE MCCD MCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo r3_mccd

* Endogeneity: IV-2SLS using first-order lags of SCCD and SCCD2.
xtreg SCCD L.SCCD L.SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo fs_sccd
xtreg SCCD2 L.SCCD L.SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo fs_sccd2

xtivreg CE $BASE_CONTROLS i.year (SCCD SCCD2 = L.SCCD L.SCCD2), fe vce(cluster id)
eststo iv_second

esttab r1_winsor r2_ntl r3_mccd iv_second using "$TABLES/table04_robustness_endogeneity.rtf", replace ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    drop(*.year) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Robustness and IV-2SLS second-stage results")

esttab fs_sccd fs_sccd2 using "$TABLES/table04_first_stage.rtf", replace ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    drop(*.year) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("First-stage regressions for lagged instruments")

est restore iv_second
scalar iv_turning_point = -_b[SCCD] / (2 * _b[SCCD2])
file open tp using "$TABLES/iv_turning_point.txt", write replace
file write tp "Turning point from IV-2SLS model: " %9.6f (iv_turning_point) _n
file write tp "Formula: -_b[SCCD] / (2 * _b[SCCD2])" _n
file close tp

log close log03
