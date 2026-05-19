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

capture log close log05_noer
log using "$LOGS/05_mediation_2006_2024_noER.log", name(log05_noer) replace text

use "$DATA", clear
keep if inrange(year, 2006, 2024)

gen double lnCE = CE
label variable lnCE "Log carbon emissions"

capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
xtset id year

global BASE_CONTROLS OPEN UR URG GI

eststo clear
xtreg OIU SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
estadd local cityfe "Yes"
estadd local yearfe "Yes"
eststo m_oiu_noer
xtreg lnCE SCCD SCCD2 OIU $BASE_CONTROLS i.year, fe vce(cluster id)
estadd local cityfe "Yes"
estadd local yearfe "Yes"
eststo lnce_oiu_noer
xtreg GTI SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
estadd local cityfe "Yes"
estadd local yearfe "Yes"
eststo m_gti_noer
xtreg lnCE SCCD SCCD2 GTI $BASE_CONTROLS i.year, fe vce(cluster id)
estadd local cityfe "Yes"
estadd local yearfe "Yes"
eststo lnce_gti_noer

esttab m_oiu_noer lnce_oiu_noer m_gti_noer lnce_gti_noer using "$TABLES/table6_mechanism_2006_2024_noER.rtf", replace ///
    mtitles("OIU" "lnCE with OIU" "GTI" "lnCE with GTI") ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OIU GTI OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OIU GTI OPEN UR URG GI _cons) ///
    stats(cityfe yearfe N r2_w, labels("City fixed effects" "Year fixed effects" "N" "Within R-squared") fmt(%9s %9s 0 3)) ///
    label compress title("Table 6. Mechanism analysis: 2006-2024 main sample")

log close log05_noer
