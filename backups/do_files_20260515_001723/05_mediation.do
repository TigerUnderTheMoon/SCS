version 17.0
set more off

capture log close log05
log using "$LOGS/05_mediation.log", name(log05) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
xtset id year

global BASE_CONTROLS OPEN UR URG GI

eststo clear

* Mechanism 1: industrial upgrading.
xtreg OIU SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo m_oiu
xtreg CE SCCD SCCD2 OIU $BASE_CONTROLS i.year, fe vce(cluster id)
eststo ce_oiu

* Mechanism 2: green technological innovation.
xtreg GTI SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo m_gti
xtreg CE SCCD SCCD2 GTI $BASE_CONTROLS i.year, fe vce(cluster id)
eststo ce_gti

esttab m_oiu ce_oiu m_gti ce_gti using "$TABLES/table06_mediation.rtf", replace ///
    mtitles("OIU" "CE with OIU" "GTI" "CE with GTI") ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OIU GTI OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OIU GTI OPEN UR URG GI _cons) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Mediation analysis")

log close log05
