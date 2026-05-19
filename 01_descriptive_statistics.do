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

capture log close log01
log using "$LOGS/01_descriptive_statistics.log", name(log01) replace text

use "$DATA", clear

capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
label variable SCCD2 "Spatial coupling coordination degree squared"

capture confirm variable is_fitted
if _rc {
    di as error "Required indicator is_fitted is missing."
    exit 111
}

xtset id year

di as text "Dataset structure"
describe

di as text "Panel diagnostics"
isid id year
xtdescribe
summarize year
tabulate year
tabulate is_fitted
tabulate region is_fitted
quietly levelsof id, local(ids)
di as result "Number of unique city ids: " wordcount("`ids'")

di as text "Missing-value diagnostics"
misstable summarize

preserve
contract year is_fitted
rename _freq observations
export delimited using "$TABLES/panel_observations_by_year_fitted_status.csv", replace
restore

preserve
contract region is_fitted
rename _freq observations
export delimited using "$TABLES/panel_observations_by_region_fitted_status.csv", replace
restore

tempname misspost
postfile `misspost' str32 variable long missing_full missing_observed using "$TABLES/missing_values_by_sample.dta", replace
foreach v of varlist _all {
    quietly count if missing(`v')
    local mf = r(N)
    quietly count if missing(`v') & is_fitted == 0
    local mo = r(N)
    post `misspost' ("`v'") (`mf') (`mo')
}
postclose `misspost'
preserve
use "$TABLES/missing_values_by_sample.dta", clear
export delimited using "$TABLES/missing_values_by_sample.csv", replace
restore

local descvars CE SCCD SCCD2 OPEN UR URG GI ER DEI POLY OIU GTI MCCD NTL FDI RD EI ISL ISU IV c_SCCD is_fitted
local keepvars
foreach v of local descvars {
    capture confirm variable `v'
    if !_rc local keepvars `keepvars' `v'
}

estpost summarize `keepvars', detail
esttab using "$TABLES/table02_descriptive_statistics_full_2006_2024.rtf", replace ///
    cells("count(fmt(0)) mean(fmt(3)) p50(fmt(3)) sd(fmt(3)) min(fmt(3)) max(fmt(3))") ///
    noobs nonumber nomtitle label title("Descriptive statistics: full sample, 2006-2024")

estpost summarize `keepvars' if is_fitted == 0, detail
esttab using "$TABLES/table02_descriptive_statistics_observed_2006_2021.rtf", replace ///
    cells("count(fmt(0)) mean(fmt(3)) p50(fmt(3)) sd(fmt(3)) min(fmt(3)) max(fmt(3))") ///
    noobs nonumber nomtitle label title("Descriptive statistics: observed sample, 2006-2021")

foreach sample in full observed {
    tempname posth
    if "`sample'" == "full" local sampleif "if !missing(id)"
    if "`sample'" == "observed" local sampleif "if is_fitted == 0"
    postfile `posth' str32 variable double N mean median sd min max using "$TABLES/table02_descriptive_statistics_`sample'.dta", replace
    foreach v of local keepvars {
        quietly summarize `v' `sampleif', detail
        post `posth' ("`v'") (r(N)) (r(mean)) (r(p50)) (r(sd)) (r(min)) (r(max))
    }
    postclose `posth'
    preserve
    use "$TABLES/table02_descriptive_statistics_`sample'.dta", clear
    export delimited using "$TABLES/table02_descriptive_statistics_`sample'.csv", replace
    restore
}

log close log01
