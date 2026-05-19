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

capture log close log03_noer
log using "$LOGS/03_robustness_endogeneity_2006_2024_noER.log", name(log03_noer) replace text

use "$DATA", clear
keep if inrange(year, 2006, 2024)

gen double lnCE = CE
label variable lnCE "Log carbon emissions"

capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
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

eststo clear
preserve
foreach v in lnCE SCCD SCCD2 OPEN UR URG GI {
    gen double w_`v' = `v'
    quietly _pctile `v' if !missing(`v'), p(1 99)
    replace w_`v' = r(r1) if w_`v' < r(r1) & !missing(w_`v')
    replace w_`v' = r(r2) if w_`v' > r(r2) & !missing(w_`v')
}
xtreg w_lnCE w_SCCD w_SCCD2 w_OPEN w_UR w_URG w_GI i.year, fe vce(cluster id)
estadd local cityfe "Yes"
estadd local yearfe "Yes"
eststo r1_winsor_noer
restore

xtreg NTL SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
estadd local cityfe "Yes"
estadd local yearfe "Yes"
eststo r2_ntl_noer

ivreg2 lnCE $BASE_CONTROLS i.id i.year (SCCD SCCD2 = IV IV2), cluster(id) first
estadd local cityfe "Yes"
estadd local yearfe "Yes"
estadd scalar kp_lm = e(idstat)
estadd scalar kp_lm_p = e(idp)
estadd scalar kp_wald_f = e(widstat)
estadd scalar cd_f = e(cdf)
estadd scalar overid_j = e(j)
estadd scalar overid_df = e(jdf)
eststo iv_second_noer

scalar iv_tp_noer = -_b[SCCD] / (2 * _b[SCCD2])
scalar diag_kp_lm = e(idstat)
scalar diag_kp_lm_p = e(idp)
scalar diag_kp_f = e(widstat)
scalar diag_cd_f = e(cdf)
scalar diag_j = e(j)
scalar diag_jdf = e(jdf)
gen byte iv_sample_noer = e(sample)

reg SCCD IV IV2 $BASE_CONTROLS i.id i.year if iv_sample_noer, vce(cluster id)
test IV IV2
scalar fs_f_sccd = r(F)
scalar fs_p_sccd = r(p)

reg SCCD2 IV IV2 $BASE_CONTROLS i.id i.year if iv_sample_noer, vce(cluster id)
test IV IV2
scalar fs_f_sccd2 = r(F)
scalar fs_p_sccd2 = r(p)

esttab r1_winsor_noer r2_ntl_noer iv_second_noer using "$TABLES/table4_robustness_iv_2006_2024_noER.rtf", replace ///
    mtitles("Winsorized lnCE" "Alternative DV: NTL" "Supplementary IV second stage") ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(w_SCCD w_SCCD2 w_OPEN w_UR w_URG w_GI SCCD SCCD2 OPEN UR URG GI _cons) ///
    order(w_SCCD w_SCCD2 SCCD SCCD2 w_OPEN w_UR w_URG w_GI OPEN UR URG GI _cons) ///
    coeflabels(w_SCCD "SCCD" w_SCCD2 "SCCD2" w_OPEN "OPEN" w_UR "UR" w_URG "URG" w_GI "GI") ///
    stats(cityfe yearfe N r2_w r2 kp_lm kp_lm_p kp_wald_f, labels("City fixed effects" "Year fixed effects" "N" "Within R-squared" "Centered R-squared" "KP rk LM" "KP rk LM p-value" "KP rk Wald F") fmt(%9s %9s 0 3 3 3 3 3)) ///
    label compress title("Table 4 Panel A. Robustness and supplementary IV results: 2006-2024 main sample")

file open dcsv using "$TABLES/table4_iv_diagnostics_2006_2024_noER.csv", write replace
file write dcsv "diagnostic,value,p_value,notes" _n
file write dcsv "First-stage F for SCCD," %9.6f (fs_f_sccd) "," %9.6f (fs_p_sccd) ",Excluded instruments IV and IV2" _n
file write dcsv "First-stage F for SCCD2," %9.6f (fs_f_sccd2) "," %9.6f (fs_p_sccd2) ",Excluded instruments IV and IV2" _n
file write dcsv "Kleibergen-Paap rk LM underidentification," %9.6f (diag_kp_lm) "," %9.6f (diag_kp_lm_p) ",Cluster-robust ivreg2 statistic" _n
file write dcsv "Cragg-Donald Wald F," %9.6f (diag_cd_f) ",,Weak-identification statistic" _n
file write dcsv "Kleibergen-Paap rk Wald F," %9.6f (diag_kp_f) ",,Cluster-robust weak-identification statistic" _n
file write dcsv "Overidentification test,,," "Not applicable: exactly identified with two endogenous regressors and two excluded instruments" _n
file close dcsv

file open dtxt using "$TABLES/table4_iv_diagnostics_2006_2024_noER.txt", write replace
file write dtxt "Table 4 Panel B. First-stage and weak-instrument diagnostics" _n
file write dtxt "First-stage F for SCCD: " %9.3f (fs_f_sccd) " (p=" %9.6f (fs_p_sccd) ")" _n
file write dtxt "First-stage F for SCCD2: " %9.3f (fs_f_sccd2) " (p=" %9.6f (fs_p_sccd2) ")" _n
file write dtxt "Kleibergen-Paap rk LM underidentification statistic: " %9.3f (diag_kp_lm) " (p=" %9.6f (diag_kp_lm_p) ")" _n
file write dtxt "Cragg-Donald Wald F statistic: " %9.3f (diag_cd_f) _n
file write dtxt "Kleibergen-Paap rk Wald F statistic: " %9.3f (diag_kp_f) _n
file write dtxt "Overidentification test: not applicable because the model is exactly identified." _n
file write dtxt "IV turning point: " %9.6f (iv_tp_noer) _n
file close dtxt

file open pout using "$TABLES/table_purified_sccd_robustness_2006_2024_noER.md", write replace
file write pout "# Purified SCCD robustness check" _n _n
file write pout "Status: not completed from the available data and scripts." _n _n
file write pout "Reason: the Stata data file contains the composite SCCD variable but does not contain the original entropy-weighted indicator-level variables needed to reconstruct SCCD after removing ER, patent-related, trade-related, fiscal-expenditure-related, or digital-overlap indicators. The current do-files also do not include the entropy-weighting construction code." _n _n
file write pout "Additional check: MCCD is identical to SCCD in the available data and therefore cannot be used as a purified SCCD proxy." _n _n
file write pout "Required inputs to complete this check: raw indicator-level panel, the entropy-weighting script, the mapping from each indicator to SCCD dimensions and criterion layers, and explicit rules for excluding overlap indicators." _n
file close pout

file open pcsv using "$TABLES/table_purified_sccd_robustness_2006_2024_noER.csv", write replace
file write pcsv "item,status" _n
file write pcsv "purified_sccd_regression,not completed" _n
file write pcsv "reason,raw indicator-level data and entropy-weighting script are missing" _n
file write pcsv "mccd_check,MCCD is identical to SCCD and is not a purified proxy" _n
file close pcsv

log close log03_noer
