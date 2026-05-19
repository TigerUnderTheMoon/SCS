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

capture log close log01_noer
log using "$LOGS/01_descriptive_statistics_2006_2024_noER.log", name(log01_noer) replace text

use "$DATA", clear
keep if inrange(year, 2006, 2024)

gen double lnCE = CE
label variable lnCE "Log carbon emissions"

capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
label variable SCCD2 "Spatial coupling coordination degree squared"

xtset id year

di as text "Dataset structure"
describe

di as text "Panel diagnostics"
isid id year
xtdescribe
summarize year
tabulate year
quietly levelsof id, local(ids)
di as result "Number of unique city ids: " wordcount("`ids'")

capture confirm variable is_fitted
if !_rc {
    tabulate is_fitted
    preserve
    contract year is_fitted
    rename _freq observations
    export delimited using "$TABLES/panel_observations_by_year_2006_2024_noER.csv", replace
    restore
}

di as text "Missing-value diagnostics"
misstable summarize

tempname misspost
postfile `misspost' str32 variable long missing_count using "$TABLES/missing_values_2006_2024_noER.dta", replace
foreach v of varlist _all {
    quietly count if missing(`v')
    post `misspost' ("`v'") (r(N))
}
postclose `misspost'
preserve
use "$TABLES/missing_values_2006_2024_noER.dta", clear
export delimited using "$TABLES/missing_values_2006_2024_noER.csv", replace
restore

local descvars lnCE SCCD SCCD2 OPEN UR URG GI DEI POLY OIU GTI MCCD NTL FDI RD EI ISL ISU IV c_SCCD is_fitted
local keepvars
foreach v of local descvars {
    capture confirm variable `v'
    if !_rc local keepvars `keepvars' `v'
}

estpost summarize `keepvars', detail
esttab using "$TABLES/table2_descriptive_statistics_2006_2024_noER.rtf", replace ///
    cells("count(fmt(0)) mean(fmt(3)) p50(fmt(3)) sd(fmt(3)) min(fmt(3)) max(fmt(3))") ///
    noobs nonumber nomtitle label title("Descriptive statistics: 2006-2024 main sample")

tempname posth
postfile `posth' str32 variable double N mean median sd min max using "$TABLES/table2_descriptive_statistics_2006_2024_noER.dta", replace
foreach v of local keepvars {
    quietly summarize `v', detail
    post `posth' ("`v'") (r(N)) (r(mean)) (r(p50)) (r(sd)) (r(min)) (r(max))
}
postclose `posth'
preserve
use "$TABLES/table2_descriptive_statistics_2006_2024_noER.dta", clear
export delimited using "$TABLES/table2_descriptive_statistics_2006_2024_noER.csv", replace
restore

log close log01_noer
