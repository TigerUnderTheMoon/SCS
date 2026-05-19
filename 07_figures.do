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

capture log close log07
log using "$LOGS/07_figures.log", name(log07) replace text

use "$DATA", clear
capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
xtset id year

graph set window fontface "Arial"

histogram SCCD, percent width(0.025) ///
    title("Distribution of SCCD, full sample") ///
    subtitle("2006-2024; 2022-2024 include fitted values") ///
    xtitle("Spatial coupling coordination degree") ytitle("Percent")
graph export "$FIGURES/fig_sccd_distribution_full_2006_2024.png", replace width(2400)

histogram SCCD if is_fitted == 0, percent width(0.025) ///
    title("Distribution of SCCD, observed-only sample") ///
    subtitle("2006-2021") ///
    xtitle("Spatial coupling coordination degree") ytitle("Percent")
graph export "$FIGURES/fig_sccd_distribution_observed_2006_2021.png", replace width(2400)

twoway ///
    (scatter CE SCCD if is_fitted == 0, mcolor(gs12%35) msize(tiny)) ///
    (scatter CE SCCD if is_fitted == 1, mcolor(maroon%45) msize(tiny)) ///
    (qfit CE SCCD, lcolor(navy) lwidth(medthick)), ///
    title("Carbon emissions and SCCD") ///
    subtitle("Full sample, fitted rows highlighted") ///
    xtitle("SCCD") ytitle("CE") ///
    legend(order(1 "Observed" 2 "Fitted" 3 "Quadratic fit") pos(6) row(1))
graph export "$FIGURES/fig_ce_sccd_quadratic_fit_full_2006_2024.png", replace width(2400)

preserve
collapse (mean) CE SCCD, by(year is_fitted)
twoway ///
    (line CE year, yaxis(1) lcolor(navy) lwidth(medthick)) ///
    (line SCCD year, yaxis(2) lcolor(maroon) lpattern(dash) lwidth(medthick)), ///
    title("Annual mean CE and SCCD") ///
    subtitle("2006-2024; 2022-2024 fitted") ///
    xtitle("Year") ytitle("Mean CE", axis(1)) ytitle("Mean SCCD", axis(2)) ///
    legend(order(1 "CE" 2 "SCCD") pos(6) row(1))
graph export "$FIGURES/fig_yearly_mean_ce_sccd_full_2006_2024.png", replace width(2400)
restore

graph box CE, over(region) ///
    title("Carbon emissions by region") ///
    subtitle("Full sample, 2006-2024") ///
    ytitle("CE")
graph export "$FIGURES/fig_ce_by_region_box_full_2006_2024.png", replace width(2400)

preserve
collapse (mean) CE SCCD, by(region year)
twoway ///
    (line CE year if region == "东部", lcolor(navy) lwidth(medthick)) ///
    (line CE year if region == "中部", lcolor(maroon) lpattern(dash) lwidth(medthick)) ///
    (line CE year if region == "西部", lcolor(forest_green) lpattern(dot) lwidth(medthick)), ///
    title("Annual mean carbon emissions by region") ///
    subtitle("2006-2024; 2022-2024 fitted") ///
    xtitle("Year") ytitle("Mean CE") ///
    legend(order(1 "Eastern" 2 "Central" 3 "Western") pos(6) row(1))
graph export "$FIGURES/fig_regional_mean_ce_trends_full_2006_2024.png", replace width(2400)
restore

log close log07
