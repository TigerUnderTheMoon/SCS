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

capture log close log04
log using "$LOGS/04_heterogeneity.log", name(log04) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
xtset id year

global BASE_CONTROLS OPEN UR URG GI

eststo clear
xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year if region == "东部", fe vce(cluster id)
eststo east_full
xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year if region == "中部", fe vce(cluster id)
eststo central_full
xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year if region == "西部", fe vce(cluster id)
eststo west_full

esttab east_full central_full west_full using "$TABLES/table05_regional_heterogeneity_full_2006_2024.rtf", replace ///
    mtitles("Eastern China" "Central China" "Western China") ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OPEN UR URG GI _cons) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Regional heterogeneity: full sample, 2006-2024")

eststo clear
xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year if region == "东部" & is_fitted == 0, fe vce(cluster id)
eststo east_obs
xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year if region == "中部" & is_fitted == 0, fe vce(cluster id)
eststo central_obs
xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year if region == "西部" & is_fitted == 0, fe vce(cluster id)
eststo west_obs

esttab east_obs central_obs west_obs using "$TABLES/table05_regional_heterogeneity_observed_2006_2021.rtf", replace ///
    mtitles("Eastern China" "Central China" "Western China") ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OPEN UR URG GI _cons) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Regional heterogeneity: observed-only sample, 2006-2021")

log close log04
