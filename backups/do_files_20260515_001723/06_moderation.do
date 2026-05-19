version 17.0
set more off

capture log close log06
log using "$LOGS/06_moderation.log", name(log06) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
xtset id year

global BASE_CONTROLS OPEN UR URG GI

eststo clear

xtreg CE c.SCCD##c.ER c.SCCD2##c.ER $BASE_CONTROLS i.year, fe vce(cluster id)
eststo mod_er
quietly summarize ER if e(sample), detail
scalar er_p25 = r(p25)
scalar er_p50 = r(p50)
scalar er_p75 = r(p75)
file open ter using "$TABLES/moderation_turning_points_ER.txt", write replace
foreach q in p25 p50 p75 {
    scalar b1 = _b[SCCD] + _b[c.SCCD#c.ER] * er_`q'
    scalar b2 = _b[SCCD2] + _b[c.SCCD2#c.ER] * er_`q'
    scalar tp = -b1 / (2 * b2)
    file write ter "ER `q' turning point: " %9.6f (tp) _n
}
file close ter

xtreg CE c.SCCD##c.DEI c.SCCD2##c.DEI $BASE_CONTROLS i.year, fe vce(cluster id)
eststo mod_dei
quietly summarize DEI if e(sample), detail
scalar dei_p25 = r(p25)
scalar dei_p50 = r(p50)
scalar dei_p75 = r(p75)
file open tdei using "$TABLES/moderation_turning_points_DEI.txt", write replace
foreach q in p25 p50 p75 {
    scalar b1 = _b[SCCD] + _b[c.SCCD#c.DEI] * dei_`q'
    scalar b2 = _b[SCCD2] + _b[c.SCCD2#c.DEI] * dei_`q'
    scalar tp = -b1 / (2 * b2)
    file write tdei "DEI `q' turning point: " %9.6f (tp) _n
}
file close tdei

xtreg CE c.SCCD##c.POLY c.SCCD2##c.POLY $BASE_CONTROLS i.year, fe vce(cluster id)
eststo mod_poly
quietly summarize POLY if e(sample), detail
scalar poly_p25 = r(p25)
scalar poly_p50 = r(p50)
scalar poly_p75 = r(p75)
file open tpoly using "$TABLES/moderation_turning_points_POLY.txt", write replace
foreach q in p25 p50 p75 {
    scalar b1 = _b[SCCD] + _b[c.SCCD#c.POLY] * poly_`q'
    scalar b2 = _b[SCCD2] + _b[c.SCCD2#c.POLY] * poly_`q'
    scalar tp = -b1 / (2 * b2)
    file write tpoly "POLY `q' turning point: " %9.6f (tp) _n
}
file close tpoly

esttab mod_er mod_dei mod_poly using "$TABLES/table07_moderation.rtf", replace ///
    mtitles("ER" "DEI" "POLY") ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    drop(*.year) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Moderation analysis")

log close log06
