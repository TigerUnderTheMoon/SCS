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

capture log close log05
log using "$LOGS/05_mediation.log", name(log05) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
xtset id year

global BASE_CONTROLS OPEN UR URG GI

eststo clear
xtreg OIU SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo m_oiu_full
xtreg CE SCCD SCCD2 OIU $BASE_CONTROLS i.year, fe vce(cluster id)
eststo ce_oiu_full
xtreg GTI SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo m_gti_full
xtreg CE SCCD SCCD2 GTI $BASE_CONTROLS i.year, fe vce(cluster id)
eststo ce_gti_full

esttab m_oiu_full ce_oiu_full m_gti_full ce_gti_full using "$TABLES/table06_mediation_full_2006_2024.rtf", replace ///
    mtitles("OIU" "CE with OIU" "GTI" "CE with GTI") ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OIU GTI OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OIU GTI OPEN UR URG GI _cons) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Mediation analysis: full sample, 2006-2024")

eststo clear
xtreg OIU SCCD SCCD2 $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo m_oiu_obs
xtreg CE SCCD SCCD2 OIU $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo ce_oiu_obs
xtreg GTI SCCD SCCD2 $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo m_gti_obs
xtreg CE SCCD SCCD2 GTI $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo ce_gti_obs

esttab m_oiu_obs ce_oiu_obs m_gti_obs ce_gti_obs using "$TABLES/table06_mediation_observed_2006_2021.rtf", replace ///
    mtitles("OIU" "CE with OIU" "GTI" "CE with GTI") ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OIU GTI OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OIU GTI OPEN UR URG GI _cons) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Mediation analysis: observed-only sample, 2006-2021")

log close log05
