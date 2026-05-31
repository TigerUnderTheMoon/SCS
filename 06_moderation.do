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
capture drop SCCD_DEI SCCD2_DEI SCCD_POLY SCCD2_POLY
gen double SCCD_DEI = SCCD * DEI
gen double SCCD2_DEI = SCCD2 * DEI
gen double SCCD_POLY = SCCD * POLY
gen double SCCD2_POLY = SCCD2 * POLY
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

xtreg CE SCCD SCCD2 DEI SCCD_DEI SCCD2_DEI $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
quietly summarize DEI if e(sample), detail
local dei_low = r(p25)
local dei_mean = r(mean)
local dei_high = r(p75)
quietly summarize SCCD if e(sample), detail
local sccd_mean = r(mean)
local sccd_min = r(min)
local sccd_max = r(max)

scalar b_cons = _b[_cons]
scalar b_sccd = _b[SCCD]
scalar b_sccd2 = _b[SCCD2]
scalar b_dei = _b[DEI]
scalar b_sccd_dei = _b[SCCD_DEI]
scalar b_sccd2_dei = _b[SCCD2_DEI]
foreach v in OPEN UR URG GI {
    quietly summarize `v' if e(sample)
    scalar mean_`v' = r(mean)
    scalar b_`v' = _b[`v']
}

preserve
clear
set obs 120
gen double SCCD = `sccd_min' + (`sccd_max' - `sccd_min') * (_n - 1) / (_N - 1)
gen double y_low = b_cons + (b_sccd + b_sccd_dei * `dei_low') * SCCD + (b_sccd2 + b_sccd2_dei * `dei_low') * SCCD^2 + b_dei * `dei_low' + b_OPEN * mean_OPEN + b_UR * mean_UR + b_URG * mean_URG + b_GI * mean_GI
gen double y_high = b_cons + (b_sccd + b_sccd_dei * `dei_high') * SCCD + (b_sccd2 + b_sccd2_dei * `dei_high') * SCCD^2 + b_dei * `dei_high' + b_OPEN * mean_OPEN + b_UR * mean_UR + b_URG * mean_URG + b_GI * mean_GI

twoway ///
    (line y_low SCCD, lcolor(navy) lwidth(medthick)) ///
    (line y_high SCCD, lcolor(maroon) lwidth(medthick)), ///
    xtitle("SCCD") ytitle("Predicted lnCE") ///
    legend(order(1 "Low DEI" 2 "High DEI") pos(6) row(1)) ///
    graphregion(color(white)) plotregion(color(white))
graph export "$FIGURES/fig_moderation_dei_observed_2006_2021.png", replace width(2400)
restore

file open ctp using "$TABLES/dei_conditional_turning_points_observed_2006_2021.csv", write replace
file write ctp "dei_level,dei_value,effective_linear,effective_quadratic,turning_point,tp_se,tp_p,tp_ci_lower,tp_ci_upper,marginal_effect_at_mean_sccd,me_se,me_p,me_ci_lower,me_ci_upper,notes" _n

foreach level in low mean high {
    local d = `dei_`level''
    scalar eff_linear = _b[SCCD] + _b[SCCD_DEI] * `d'
    scalar eff_quadratic = _b[SCCD2] + _b[SCCD2_DEI] * `d'
    nlcom ///
        (turning_point: -(_b[SCCD] + _b[SCCD_DEI] * `d') / (2 * (_b[SCCD2] + _b[SCCD2_DEI] * `d'))) ///
        (marginal_effect_at_mean_sccd: _b[SCCD] + _b[SCCD_DEI] * `d' + 2 * `sccd_mean' * (_b[SCCD2] + _b[SCCD2_DEI] * `d'))
    matrix C = r(table)
    scalar tp_b = C[1,1]
    scalar tp_se = C[2,1]
    scalar tp_p = C[4,1]
    scalar tp_lb = C[5,1]
    scalar tp_ub = C[6,1]
    scalar me_b = C[1,2]
    scalar me_se = C[2,2]
    scalar me_p = C[4,2]
    scalar me_lb = C[5,2]
    scalar me_ub = C[6,2]
    file write ctp "`level'," %9.6f (`d') "," %9.6f (eff_linear) "," %9.6f (eff_quadratic) "," %9.6f (tp_b) "," %9.6f (tp_se) "," %9.6f (tp_p) "," %9.6f (tp_lb) "," %9.6f (tp_ub) "," %9.6f (me_b) "," %9.6f (me_se) "," %9.6f (me_p) "," %9.6f (me_lb) "," %9.6f (me_ub) ",DEI low=p25 mean=mean high=p75; SCCD mean=" %9.6f (`sccd_mean') _n
}
file close ctp

file open ctp_txt using "$TABLES/dei_conditional_turning_points_observed_2006_2021.txt", write replace
file write ctp_txt "DEI conditional turning points and marginal effects for observed 2006-2021 sample" _n
file write ctp_txt "DEI low/mean/high are p25, mean, and p75 among e(sample)." _n
file write ctp_txt "SCCD range among e(sample): [" %9.6f (`sccd_min') ", " %9.6f (`sccd_max') "]; SCCD mean = " %9.6f (`sccd_mean') _n
file write ctp_txt "See CSV for nlcom standard errors and confidence intervals." _n
file close ctp_txt

log close log06
