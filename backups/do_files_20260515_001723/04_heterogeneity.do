version 17.0
set more off

capture log close log04
log using "$LOGS/04_heterogeneity.log", name(log04) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
xtset id year

global BASE_CONTROLS OPEN UR URG GI

eststo clear

xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year if region == "东部", fe vce(cluster id)
eststo east

xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year if region == "中部", fe vce(cluster id)
eststo central

xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year if region == "西部", fe vce(cluster id)
eststo west

esttab east central west using "$TABLES/table05_regional_heterogeneity.rtf", replace ///
    mtitles("Eastern China" "Central China" "Western China") ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OPEN UR URG GI _cons) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Regional heterogeneity")

log close log04
