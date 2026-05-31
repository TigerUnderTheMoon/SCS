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

capture log close log02
log using "$LOGS/02_baseline_regression.log", name(log02) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
label variable SCCD2 "Spatial coupling coordination degree squared"
xtset id year

global BASE_CONTROLS OPEN UR URG GI

eststo clear
xtreg CE SCCD SCCD2 i.year, fe vce(cluster id)
eststo b1_full
xtreg CE SCCD SCCD2 OPEN i.year, fe vce(cluster id)
eststo b2_full
xtreg CE SCCD SCCD2 OPEN UR i.year, fe vce(cluster id)
eststo b3_full
xtreg CE SCCD SCCD2 OPEN UR URG i.year, fe vce(cluster id)
eststo b4_full
xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year, fe vce(cluster id)
eststo b5_full

esttab b1_full b2_full b3_full b4_full b5_full using "$TABLES/table03_baseline_full_2006_2024.rtf", replace ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OPEN UR URG GI _cons) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Baseline regression: full sample, 2006-2024")

est restore b5_full
scalar tp_full = -_b[SCCD] / (2 * _b[SCCD2])

eststo clear
xtreg CE SCCD SCCD2 i.year if is_fitted == 0, fe vce(cluster id)
eststo b1_obs
xtreg CE SCCD SCCD2 OPEN i.year if is_fitted == 0, fe vce(cluster id)
eststo b2_obs
xtreg CE SCCD SCCD2 OPEN UR i.year if is_fitted == 0, fe vce(cluster id)
eststo b3_obs
xtreg CE SCCD SCCD2 OPEN UR URG i.year if is_fitted == 0, fe vce(cluster id)
eststo b4_obs
xtreg CE SCCD SCCD2 $BASE_CONTROLS i.year if is_fitted == 0, fe vce(cluster id)
eststo b5_obs

esttab b1_obs b2_obs b3_obs b4_obs b5_obs using "$TABLES/table03_baseline_observed_2006_2021.rtf", replace ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    keep(SCCD SCCD2 OPEN UR URG GI _cons) ///
    order(SCCD SCCD2 OPEN UR URG GI _cons) ///
    stats(N r2_w, labels("N" "Within R-squared") fmt(0 3)) ///
    label compress title("Baseline regression: observed-only sample, 2006-2021")

est restore b5_obs
scalar tp_observed = -_b[SCCD] / (2 * _b[SCCD2])
capture drop baseline_obs_sample
gen byte baseline_obs_sample = e(sample)
quietly summarize SCCD if baseline_obs_sample, detail
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

lincom SCCD + 2*`xmean'*SCCD2
scalar mean_b = r(estimate)
scalar mean_se = r(se)
scalar mean_p = r(p)
scalar mean_lb = r(lb)
scalar mean_ub = r(ub)

lincom SCCD + 2*`xmax'*SCCD2
scalar right_b = r(estimate)
scalar right_se = r(se)
scalar right_p = r(p)
scalar right_lb = r(lb)
scalar right_ub = r(ub)

local inrange = (tp >= `xmin' & tp <= `xmax')
local endpoint_pass = (left_b > 0 & left_p < 0.05 & right_b < 0 & right_p < 0.05)

quietly count if baseline_obs_sample
scalar sample_n = r(N)
quietly count if baseline_obs_sample & SCCD <= tp_observed
scalar n_at_or_below_tp = r(N)
quietly count if baseline_obs_sample & SCCD > tp_observed
scalar n_above_tp = r(N)
quietly levelsof id if baseline_obs_sample & SCCD > tp_observed, local(cities_above_tp)
local n_cities_above_tp : word count `cities_above_tp'
scalar pct_at_or_below_tp = 100 * n_at_or_below_tp / sample_n
scalar pct_above_tp = 100 * n_above_tp / sample_n

file open ucsv using "$TABLES/table_u_test_observed_2006_2021.csv", write replace
file write ucsv "test_item,estimate,se,p_value,ci_lower,ci_upper,notes" _n
file write ucsv "turning_point," %9.6f (tp) "," %9.6f (tp_se) "," %9.6f (tp_p) "," %9.6f (tp_lb) "," %9.6f (tp_ub) ",nlcom -b1/(2*b2)" _n
file write ucsv "left_endpoint_marginal_effect," %9.6f (left_b) "," %9.6f (left_se) "," %9.6f (left_p) "," %9.6f (left_lb) "," %9.6f (left_ub) ",SCCD minimum = " %9.6f (`xmin') _n
file write ucsv "mean_sccd_marginal_effect," %9.6f (mean_b) "," %9.6f (mean_se) "," %9.6f (mean_p) "," %9.6f (mean_lb) "," %9.6f (mean_ub) ",SCCD mean = " %9.6f (`xmean') _n
file write ucsv "right_endpoint_marginal_effect," %9.6f (right_b) "," %9.6f (right_se) "," %9.6f (right_p) "," %9.6f (right_lb) "," %9.6f (right_ub) ",SCCD maximum = " %9.6f (`xmax') _n
file close ucsv

file open utxt using "$TABLES/table_u_test_observed_2006_2021.txt", write replace
file write utxt "Formal inverted-U test for the observed-sample full-control baseline model" _n
file write utxt "Dependent variable: CE, interpreted as lnCE" _n
file write utxt "SCCD range in e(sample): [" %9.6f (`xmin') ", " %9.6f (`xmax') "]" _n
file write utxt "SCCD mean/p25/p75: " %9.6f (`xmean') " / " %9.6f (`xp25') " / " %9.6f (`xp75') _n
file write utxt "Turning point: " %9.6f (tp) " (SE " %9.6f (tp_se) ", 95% CI [" %9.6f (tp_lb) ", " %9.6f (tp_ub) "])" _n
file write utxt "Left endpoint marginal effect: " %9.6f (left_b) " (p=" %9.6f (left_p) ")" _n
file write utxt "Right endpoint marginal effect: " %9.6f (right_b) " (p=" %9.6f (right_p) ")" _n
file write utxt "Turning point inside SCCD range: `inrange'" _n
file write utxt "Endpoint-significance inverted-U check passed at 5%: `endpoint_pass'" _n
file write utxt "This is the endpoint marginal-effect equivalent of a Lind-Mehlum inverted-U check." _n
file close utxt

file open spcsv using "$TABLES/sample_position_relative_to_turning_point_observed_2006_2021.csv", write replace
file write spcsv "item,value,notes" _n
file write spcsv "baseline_estimation_sample_n," %12.0f (sample_n) ",Full-control observed baseline e(sample)" _n
file write spcsv "turning_point," %9.6f (tp_observed) ",Observed baseline turning point" _n
file write spcsv "observations_at_or_below_turning_point," %12.0f (n_at_or_below_tp) ",SCCD <= turning point" _n
file write spcsv "observations_above_turning_point," %12.0f (n_above_tp) ",SCCD > turning point" _n
file write spcsv "percent_at_or_below_turning_point," %9.6f (pct_at_or_below_tp) ",Percent of e(sample)" _n
file write spcsv "percent_above_turning_point," %9.6f (pct_above_tp) ",Percent of e(sample)" _n
file write spcsv "cities_with_any_observation_above_turning_point," %12.0f (`n_cities_above_tp') ",Unique ids with SCCD > turning point in e(sample)" _n
file close spcsv

quietly count if is_fitted == 0 & !missing(OPEN)
scalar open_obs_n = r(N)
quietly count if is_fitted == 0 & OPEN < 0
scalar open_neg_n = r(N)
quietly count if is_fitted == 0 & missing(OPEN)
scalar open_missing_n = r(N)
quietly summarize OPEN if is_fitted == 0, detail
scalar open_min = r(min)
scalar open_mean = r(mean)
scalar open_p25 = r(p25)
scalar open_p50 = r(p50)
scalar open_p75 = r(p75)
scalar open_max = r(max)
scalar open_neg_pct = 100 * open_neg_n / open_obs_n

file open oqcsv using "$TABLES/open_data_quality_observed_2006_2021.csv", write replace
file write oqcsv "item,value,notes" _n
file write oqcsv "nonmissing_observed_OPEN," %12.0f (open_obs_n) ",Observed sample nonmissing OPEN" _n
file write oqcsv "missing_observed_OPEN," %12.0f (open_missing_n) ",Observed sample missing OPEN" _n
file write oqcsv "negative_observed_OPEN," %12.0f (open_neg_n) ",OPEN less than zero" _n
file write oqcsv "negative_observed_OPEN_percent," %9.6f (open_neg_pct) ",Share among nonmissing observed OPEN" _n
file write oqcsv "OPEN_min," %9.6f (open_min) ",Observed sample" _n
file write oqcsv "OPEN_mean," %9.6f (open_mean) ",Observed sample" _n
file write oqcsv "OPEN_p25," %9.6f (open_p25) ",Observed sample" _n
file write oqcsv "OPEN_median," %9.6f (open_p50) ",Observed sample" _n
file write oqcsv "OPEN_p75," %9.6f (open_p75) ",Observed sample" _n
file write oqcsv "OPEN_max," %9.6f (open_max) ",Observed sample" _n
file close oqcsv

file open oqtxt using "$TABLES/open_data_quality_observed_2006_2021.txt", write replace
file write oqtxt "OPEN data-quality diagnostic for observed 2006-2021 sample" _n
file write oqtxt "Nonmissing OPEN observations: " %12.0f (open_obs_n) _n
file write oqtxt "Missing OPEN observations: " %12.0f (open_missing_n) _n
file write oqtxt "Negative OPEN observations: " %12.0f (open_neg_n) " (" %9.3f (open_neg_pct) "% of nonmissing OPEN)" _n
file write oqtxt "OPEN range: [" %9.6f (open_min) ", " %9.6f (open_max) "]" _n
file write oqtxt "Interpretation note: negative OPEN values are retained and disclosed rather than recoded ex post because the data file is treated as authoritative input." _n
file close oqtxt

file open tp using "$TABLES/baseline_turning_points.txt", write replace
file write tp "Full-sample turning point, 2006-2024: " %9.6f (tp_full) _n
file write tp "Observed-only turning point, 2006-2021: " %9.6f (tp_observed) _n
file write tp "Formula: -_b[SCCD] / (2 * _b[SCCD2])" _n
file close tp

log close log02
