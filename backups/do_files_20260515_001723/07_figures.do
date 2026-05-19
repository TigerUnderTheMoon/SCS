version 17.0
set more off

capture log close log07
log using "$LOGS/07_figures.log", name(log07) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
xtset id year

global BASE_CONTROLS OPEN UR URG GI

graph set window fontface "Arial"

histogram SCCD, percent width(0.025) ///
    title("Distribution of SCCD") ///
    xtitle("Spatial coupling coordination degree") ytitle("Percent")
graph export "$FIGURES/fig_sccd_distribution.png", replace width(2400)

twoway ///
    (scatter CE SCCD, mcolor(gs12%35) msize(tiny)) ///
    (qfit CE SCCD, lcolor(navy) lwidth(medthick)), ///
    title("Carbon emissions and SCCD") ///
    xtitle("SCCD") ytitle("CE") ///
    legend(order(2 "Quadratic fit") pos(6) row(1))
graph export "$FIGURES/fig_ce_sccd_quadratic_fit.png", replace width(2400)

preserve
collapse (mean) CE SCCD DEI GTI, by(year)
twoway ///
    (line CE year, yaxis(1) lcolor(navy) lwidth(medthick)) ///
    (line SCCD year, yaxis(2) lcolor(maroon) lpattern(dash) lwidth(medthick)), ///
    title("Annual mean CE and SCCD") ///
    xtitle("Year") ytitle("Mean CE", axis(1)) ytitle("Mean SCCD", axis(2)) ///
    legend(order(1 "CE" 2 "SCCD") pos(6) row(1))
graph export "$FIGURES/fig_yearly_mean_ce_sccd.png", replace width(2400)
restore

graph box CE, over(region) ///
    title("Carbon emissions by region") ///
    ytitle("CE")
graph export "$FIGURES/fig_ce_by_region_box.png", replace width(2400)

preserve
collapse (mean) CE SCCD, by(region year)
twoway ///
    (line CE year if region == "东部", lcolor(navy) lwidth(medthick)) ///
    (line CE year if region == "中部", lcolor(maroon) lpattern(dash) lwidth(medthick)) ///
    (line CE year if region == "西部", lcolor(forest_green) lpattern(dot) lwidth(medthick)), ///
    title("Annual mean carbon emissions by region") ///
    xtitle("Year") ytitle("Mean CE") ///
    legend(order(1 "Eastern" 2 "Central" 3 "Western") pos(6) row(1))
graph export "$FIGURES/fig_regional_mean_ce_trends.png", replace width(2400)
restore

log close log07
