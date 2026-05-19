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

capture log close log06
log using "$LOGS/06_moderation.log", name(log06) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
xtset id year

global BASE_CONTROLS OPEN UR URG GI

eststo clear
xtreg CE c.SCCD##c.ER c.SCCD2##c.ER $BASE_CONTROLS i.year, fe vce(cluster id)
eststo mod_er_full
xtreg CE c.SCCD##c.DEI c.SCCD2##c.DEI $BASE_CONTROLS i.year, fe vce(cluster id)
eststo mod_dei_full
xtreg CE c.SCCD##c.POLY c.SCCD2##c.POLY $BASE_CONTROLS i.year, fe vce(cluster id)
eststo mod_poly_full

esttab mod_er_full mod_dei_full mod_poly_full using "$TABLES/table07_moderation_full_2006_2024.rtf", replace ///
    mtitles("ER" "DEI" "POLY") ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    drop(*.year) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Moderation analysis: full sample, 2006-2024")

eststo clear
xtreg CE c.SCCD##c.ER c.SCCD2##c.ER $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo mod_er_obs
xtreg CE c.SCCD##c.DEI c.SCCD2##c.DEI $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo mod_dei_obs
xtreg CE c.SCCD##c.POLY c.SCCD2##c.POLY $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo mod_poly_obs

esttab mod_er_obs mod_dei_obs mod_poly_obs using "$TABLES/table07_moderation_observed_2006_2021.rtf", replace ///
    mtitles("ER" "DEI" "POLY") ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    drop(*.year) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Moderation analysis: observed-only sample, 2006-2021")

log close log06
