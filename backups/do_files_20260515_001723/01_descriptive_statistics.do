version 17.0
set more off

capture log close log01
log using "$LOGS/01_descriptive_statistics.log", name(log01) replace text

use "$DATA", clear

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
tabulate region
quietly levelsof id, local(ids)
di as result "Number of unique city ids: " wordcount("`ids'")

di as text "Missing-value diagnostics"
misstable summarize

preserve
contract year
rename _freq observations
export delimited using "$TABLES/panel_observations_by_year.csv", replace
restore

preserve
contract region
rename _freq observations
export delimited using "$TABLES/panel_observations_by_region.csv", replace
restore

local descvars CE SCCD SCCD2 OPEN UR URG GI ER DEI POLY OIU GTI MCCD NTL FDI RD EI ISL ISU
local keepvars
foreach v of local descvars {
    capture confirm variable `v'
    if !_rc local keepvars `keepvars' `v'
}

estpost summarize `keepvars', detail
esttab using "$TABLES/table02_descriptive_statistics.rtf", replace ///
    cells("count(fmt(0)) mean(fmt(3)) p50(fmt(3)) sd(fmt(3)) min(fmt(3)) max(fmt(3))") ///
    noobs nonumber nomtitle label title("Descriptive statistics")

tempname posth
postfile `posth' str32 variable double N mean median sd min max using "$TABLES/table02_descriptive_statistics.dta", replace
foreach v of local keepvars {
    quietly summarize `v', detail
    post `posth' ("`v'") (r(N)) (r(mean)) (r(p50)) (r(sd)) (r(min)) (r(max))
}
postclose `posth'
use "$TABLES/table02_descriptive_statistics.dta", clear
export delimited using "$TABLES/table02_descriptive_statistics.csv", replace

log close log01
