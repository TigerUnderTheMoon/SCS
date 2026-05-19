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

capture log close log03
log using "$LOGS/03_robustness_endogeneity.log", name(log03) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
capture confirm variable MCCD2
if _rc gen double MCCD2 = MCCD^2
capture confirm variable IV
if _rc {
    di as error "Required IV variable is missing. IV should be the lagged SCCD variable."
    exit 111
}
capture confirm variable IV2
if _rc gen double IV2 = IV^2
label variable IV2 "Squared IV based on lagged SCCD"
xtset id year

global BASE_CONTROLS OPEN UR URG GI

* Full sample: 2006-2024, including fitted observations.
eststo clear
preserve
foreach v in CE SCCD SCCD2 OPEN UR URG GI {
    gen double w_`v' = `v'
    quietly _pctile `v' if !missing(`v'), p(1 99)
    replace w_`v' = r(r1) if w_`v' < r(r1) & !missing(w_`v')
    replace w_`v' = r(r2) if w_`v' > r(r2) & !missing(w_`v')
}
xtreg w_CE w_SCCD w_SCCD2 w_OPEN w_UR w_URG w_GI i.year, fe vce(cluster id)
eststo r1_winsor_full
restore

xtreg NTL SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo r2_ntl_full

xtreg CE MCCD MCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo r3_mccd_full

xtreg SCCD IV IV2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo fs_sccd_full
xtreg SCCD2 IV IV2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo fs_sccd2_full
xtivreg CE $BASE_CONTROLS i.year (SCCD SCCD2 = IV IV2), fe vce(cluster id)
eststo iv_second_full

esttab r1_winsor_full r2_ntl_full r3_mccd_full iv_second_full using "$TABLES/table04_robustness_endogeneity_full_2006_2024.rtf", replace ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    drop(*.year) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Robustness and IV-2SLS: full sample, 2006-2024")

esttab fs_sccd_full fs_sccd2_full using "$TABLES/table04_first_stage_full_2006_2024.rtf", replace ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    drop(*.year) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("First-stage regressions: full sample, 2006-2024")

est restore iv_second_full
scalar iv_tp_full = -_b[SCCD] / (2 * _b[SCCD2])

* Sensitivity sample: original observed data only.
eststo clear
preserve
keep if is_fitted == 0
foreach v in CE SCCD SCCD2 OPEN UR URG GI {
    gen double w_`v' = `v'
    quietly _pctile `v' if !missing(`v'), p(1 99)
    replace w_`v' = r(r1) if w_`v' < r(r1) & !missing(w_`v')
    replace w_`v' = r(r2) if w_`v' > r(r2) & !missing(w_`v')
}
xtreg w_CE w_SCCD w_SCCD2 w_OPEN w_UR w_URG w_GI i.year, fe vce(cluster id)
eststo r1_winsor_obs
restore

xtreg NTL SCCD SCCD2 $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo r2_ntl_obs

xtreg CE MCCD MCCD2 $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo r3_mccd_obs

xtreg SCCD IV IV2 $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo fs_sccd_obs
xtreg SCCD2 IV IV2 $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo fs_sccd2_obs
xtivreg CE $BASE_CONTROLS i.year (SCCD SCCD2 = IV IV2) if is_fitted == 0, fe vce(cluster id)
eststo iv_second_obs

esttab r1_winsor_obs r2_ntl_obs r3_mccd_obs iv_second_obs using "$TABLES/table04_robustness_endogeneity_observed_2006_2021.rtf", replace ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    drop(*.year) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Robustness and IV-2SLS: observed-only sample, 2006-2021")

esttab fs_sccd_obs fs_sccd2_obs using "$TABLES/table04_first_stage_observed_2006_2021.rtf", replace ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    drop(*.year) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("First-stage regressions: observed-only sample, 2006-2021")

est restore iv_second_obs
scalar iv_tp_observed = -_b[SCCD] / (2 * _b[SCCD2])

file open tp using "$TABLES/iv_turning_points.txt", write replace
file write tp "Full-sample IV turning point, 2006-2024: " %9.6f (iv_tp_full) _n
file write tp "Observed-only IV turning point, 2006-2021: " %9.6f (iv_tp_observed) _n
file write tp "Formula: -_b[SCCD] / (2 * _b[SCCD2])" _n
file close tp

log close log03
