# AGENTS.md

## Project
This is an empirical manuscript revision project for Sustainable Cities and Society.

## Data
The available Stata data file is 插值回归.dta.

## Rules
- Do not invent data, coefficients, years, p-values, figures, or references.
- Reconstruct clean Stata do-files from the available data and manuscript tables.
- If manuscript text conflicts with the data, flag the issue in issues_remaining.md instead of silently fixing it.
- Preserve the main empirical logic: two-way fixed effects, nonlinear SCCD effect, robustness tests, IV-2SLS, mediation analysis, moderation analysis, and regional heterogeneity.
- Use id as city identifier and year as time identifier.
- Use CE as dependent variable.
- Use SCCD and SCCD squared as core explanatory variables.
- Use OPEN, UR, URG, and GI as baseline controls unless the task says otherwise.
- Do not overwrite the original .dta file.
- Create new clean do-files only.

## Required outputs
- 00_master.do
- 01_descriptive_statistics.do
- 02_baseline_regression.do
- 03_robustness_endogeneity.do
- 04_heterogeneity.do
- 05_mediation.do
- 06_moderation.do
- 07_figures.do
- issues_remaining.md
- change_log.md