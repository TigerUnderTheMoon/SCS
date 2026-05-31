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

capture log close log07_noer
log using "$LOGS/07_figures_2006_2024_noER.log", name(log07_noer) replace text

use "$DATA", clear
keep if inrange(year, 2006, 2024)

gen double lnCE = CE
label variable lnCE "Log carbon emissions"

capture confirm variable SCCD2
if _rc gen double SCCD2 = SCCD^2
xtset id year

graph set window fontface "Arial"

histogram SCCD, percent width(0.025) ///
    xtitle("Spatial coupling coordination degree") ytitle("Percent") ///
    title("") subtitle("") ///
    graphregion(color(white)) plotregion(color(white))
graph export "$FIGURES/fig_sccd_distribution_2006_2024_noER.png", replace width(2400)

twoway ///
    (scatter lnCE SCCD, mcolor(gs12%35) msize(tiny)) ///
    (qfit lnCE SCCD, lcolor(navy) lwidth(medthick)), ///
    xtitle("SCCD") ytitle("lnCE") ///
    title("") subtitle("") ///
    legend(order(1 "City-year observations" 2 "Quadratic fit") pos(6) row(1)) ///
    graphregion(color(white)) plotregion(color(white))
graph export "$FIGURES/fig_nonlinear_sccd_lnce_2006_2024_noER.png", replace width(2400)

log close log07_noer
