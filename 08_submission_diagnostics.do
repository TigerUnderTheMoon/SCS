version 17.0
set more off

if "$ROOT" == "" {
    global ROOT "D:/Workplace/SCS"
    cd "$ROOT"
    global DATA "$ROOT/鎻掑€煎洖褰抇2006_2024_鎷熷悎鏇存柊.dta"
    global OUT "$ROOT/outputs"
    global TABLES "$OUT/tables"
    global FIGURES "$OUT/figures"
    global LOGS "$OUT/logs"
    cap mkdir "$OUT"
    cap mkdir "$TABLES"
    cap mkdir "$FIGURES"
    cap mkdir "$LOGS"
}

capture log close log08
log using "$LOGS/08_submission_diagnostics.log", name(log08) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
xtset id year

global BASE_CONTROLS OPEN UR URG GI
local main_rhs SCCD SCCD2 $BASE_CONTROLS

xtreg CE `main_rhs' i.year if is_fitted == 0, fe vce(cluster id)
capture drop baseline_obs_sample
gen byte baseline_obs_sample = e(sample)

tempname vifpost
postfile `vifpost' str32 variable double vif tolerance using "$TABLES/table08_vif_observed_2006_2021.dta", replace
foreach v of local main_rhs {
    local others
    foreach x of local main_rhs {
        if "`x'" != "`v'" local others `others' `x'
    }
    quietly regress `v' `others' i.year if baseline_obs_sample
    scalar vif_value = 1 / (1 - e(r2))
    scalar tolerance_value = 1 / vif_value
    post `vifpost' ("`v'") (vif_value) (tolerance_value)
}
postclose `vifpost'

preserve
use "$TABLES/table08_vif_observed_2006_2021.dta", clear
sort variable
export delimited using "$TABLES/table08_vif_observed_2006_2021.csv", replace
quietly summarize vif
scalar mean_vif = r(mean)
scalar max_vif = r(max)
file open vtxt using "$TABLES/table08_vif_observed_2006_2021.txt", write replace
file write vtxt "Variance inflation factor diagnostic for observed baseline regressors" _n
file write vtxt "Sample: observed 2006-2021 full-control baseline e(sample)" _n
file write vtxt "Auxiliary regressions include year indicators and exclude city indicators." _n
file write vtxt "Mean VIF: " %9.3f (mean_vif) _n
file write vtxt "Maximum VIF: " %9.3f (max_vif) _n
file write vtxt "Interpretation note: SCCD and SCCD2 are expected to have elevated VIF values because the quadratic specification includes a level term and its square." _n
file close vtxt
restore

regress CE `main_rhs' i.id i.year if baseline_obs_sample
scalar fe_n = e(N)

testparm i.id
scalar city_fe_F = r(F)
scalar city_fe_df = r(df)
scalar city_fe_df_r = r(df_r)
scalar city_fe_p = r(p)

testparm i.year
scalar year_fe_F = r(F)
scalar year_fe_df = r(df)
scalar year_fe_df_r = r(df_r)
scalar year_fe_p = r(p)

file open fecsv using "$TABLES/table08_fixed_effects_joint_tests_observed_2006_2021.csv", write replace
file write fecsv "test,F_stat,df,df_r,p_value,N,notes" _n
file write fecsv "City fixed effects," %12.6f (city_fe_F) "," %9.0f (city_fe_df) "," %9.0f (city_fe_df_r) "," %12.6f (city_fe_p) "," %9.0f (fe_n) ",Joint test of city indicators in pooled OLS dummy-variable diagnostic" _n
file write fecsv "Year fixed effects," %12.6f (year_fe_F) "," %9.0f (year_fe_df) "," %9.0f (year_fe_df_r) "," %12.6f (year_fe_p) "," %9.0f (fe_n) ",Joint test of year indicators in pooled OLS dummy-variable diagnostic" _n
file close fecsv

file open fetxt using "$TABLES/table08_fixed_effects_joint_tests_observed_2006_2021.txt", write replace
file write fetxt "Fixed-effect joint significance tests for observed full-control baseline sample" _n
file write fetxt "Diagnostic model: pooled OLS with SCCD, SCCD2, OPEN, UR, URG, GI, city indicators, and year indicators." _n
file write fetxt "Main manuscript estimates remain two-way fixed-effects regressions with city-clustered standard errors." _n
file write fetxt "City fixed effects: F(" %9.0f (city_fe_df) ", " %9.0f (city_fe_df_r) ") = " %9.3f (city_fe_F) ", p = " %9.6f (city_fe_p) _n
file write fetxt "Year fixed effects: F(" %9.0f (year_fe_df) ", " %9.0f (year_fe_df_r) ") = " %9.3f (year_fe_F) ", p = " %9.6f (year_fe_p) _n
file close fetxt

log close log08
