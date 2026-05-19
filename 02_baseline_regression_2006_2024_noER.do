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

capture log close log02_noer
log using "$LOGS/02_baseline_regression_2006_2024_noER.log", name(log02_noer) replace text

use "$DATA", clear
keep if inrange(year, 2006, 2024)

gen double lnCE = CE
label variable lnCE "Log carbon emissions"

capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
label variable SCCD2 "Spatial coupling coordination degree squared"

xtset id year

global BASE_CONTROLS OPEN UR URG GI

eststo clear
xtreg lnCE SCCD SCCD2 i.year, fe vce(cluster id)
estadd local cityfe "Yes"
estadd local yearfe "Yes"
eststo b1_noer
xtreg lnCE SCCD SCCD2 OPEN i.year, fe vce(cluster id)
estadd local cityfe "Yes"
estadd local yearfe "Yes"
eststo b2_noer
xtreg lnCE SCCD SCCD2 OPEN UR i.year, fe vce(cluster id)
estadd local cityfe "Yes"
estadd local yearfe "Yes"
eststo b3_noer
xtreg lnCE SCCD SCCD2 OPEN UR URG i.year, fe vce(cluster id)
estadd local cityfe "Yes"
estadd local yearfe "Yes"
eststo b4_noer
xtreg lnCE SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
estadd local cityfe "Yes"
estadd local yearfe "Yes"
eststo b5_noer

esttab b1_noer b2_noer b3_noer b4_noer b5_noer using "$TABLES/table3_baseline_2006_2024_noER.rtf", replace ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OPEN UR URG GI _cons) ///
    stats(cityfe yearfe N r2_w, labels("City fixed effects" "Year fixed effects" "N" "Within R-squared") fmt(%9s %9s 0 3)) ///
    label compress title("Table 3. Baseline regression: 2006-2024 main sample")

est restore b5_noer
scalar tp_baseline = -_b[SCCD] / (2 * _b[SCCD2])
quietly summarize SCCD if e(sample), detail
local xmin = r(min)
local xmax = r(max)
local xmean = r(mean)
local xp25 = r(p25)
local xp75 = r(p75)

nlcom turning_point: -_b[SCCD] / (2 * _b[SCCD2])
matrix TP = r(table)
scalar tp = TP[1,1]
scalar tp_se = TP[2,1]
scalar tp_p = TP[4,1]
scalar tp_lb = TP[5,1]
scalar tp_ub = TP[6,1]

lincom SCCD + 2*`xmin'*SCCD2
scalar left_b = r(estimate)
scalar left_se = r(se)
scalar left_p = r(p)
scalar left_lb = r(lb)
scalar left_ub = r(ub)

lincom SCCD + 2*`xmax'*SCCD2
scalar right_b = r(estimate)
scalar right_se = r(se)
scalar right_p = r(p)
scalar right_lb = r(lb)
scalar right_ub = r(ub)

lincom SCCD + 2*`xmean'*SCCD2
scalar mean_b = r(estimate)
scalar mean_se = r(se)
scalar mean_p = r(p)
scalar mean_lb = r(lb)
scalar mean_ub = r(ub)

local inrange = (tp >= `xmin' & tp <= `xmax')
local endpoint_pass = (left_b > 0 & left_p < 0.05 & right_b < 0 & right_p < 0.05)

file open ucsv using "$TABLES/table_u_test_2006_2024_noER.csv", write replace
file write ucsv "test_item,estimate,se,p_value,ci_lower,ci_upper,notes" _n
file write ucsv "turning_point," %9.6f (tp) "," %9.6f (tp_se) "," %9.6f (tp_p) "," %9.6f (tp_lb) "," %9.6f (tp_ub) ",nlcom -b1/(2*b2)" _n
file write ucsv "left_endpoint_marginal_effect," %9.6f (left_b) "," %9.6f (left_se) "," %9.6f (left_p) "," %9.6f (left_lb) "," %9.6f (left_ub) ",SCCD minimum = " %9.6f (`xmin') _n
file write ucsv "mean_sccd_marginal_effect," %9.6f (mean_b) "," %9.6f (mean_se) "," %9.6f (mean_p) "," %9.6f (mean_lb) "," %9.6f (mean_ub) ",SCCD mean = " %9.6f (`xmean') _n
file write ucsv "right_endpoint_marginal_effect," %9.6f (right_b) "," %9.6f (right_se) "," %9.6f (right_p) "," %9.6f (right_lb) "," %9.6f (right_ub) ",SCCD maximum = " %9.6f (`xmax') _n
file close ucsv

file open utxt using "$TABLES/table_u_test_2006_2024_noER.txt", write replace
file write utxt "Formal inverted-U test for the full-control baseline model" _n
file write utxt "Dependent variable: lnCE" _n
file write utxt "SCCD range in e(sample): [" %9.6f (`xmin') ", " %9.6f (`xmax') "]" _n
file write utxt "SCCD mean/p25/p75: " %9.6f (`xmean') " / " %9.6f (`xp25') " / " %9.6f (`xp75') _n
file write utxt "Turning point: " %9.6f (tp) " (SE " %9.6f (tp_se) ", 95% CI [" %9.6f (tp_lb) ", " %9.6f (tp_ub) "])" _n
file write utxt "Left endpoint marginal effect: " %9.6f (left_b) " (p=" %9.6f (left_p) ")" _n
file write utxt "Right endpoint marginal effect: " %9.6f (right_b) " (p=" %9.6f (right_p) ")" _n
file write utxt "Turning point inside SCCD range: `inrange'" _n
file write utxt "Endpoint-significance inverted-U check passed at 5%: `endpoint_pass'" _n
file write utxt "This is the endpoint marginal-effect equivalent of a Lind-Mehlum inverted-U check." _n
file close utxt

file open tpout using "$TABLES/baseline_turning_point_2006_2024_noER.txt", write replace
file write tpout "Baseline full-control turning point, 2006-2024 main sample: " %9.6f (tp_baseline) _n
file write tpout "Formula: -_b[SCCD] / (2 * _b[SCCD2])" _n
file close tpout

log close log02_noer
