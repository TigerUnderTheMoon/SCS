# Revision report: 2006-2024 main sample, no ER

Generated: 2026-05-16

## Files created

- `revised_manuscript_2006_2024_noER.docx`
- `_revise_manuscript_2006_2024_noER.py`
- `00_master_2006_2024_noER.do`
- `01_descriptive_statistics_2006_2024_noER.do`
- `02_baseline_regression_2006_2024_noER.do`
- `03_robustness_endogeneity_2006_2024_noER.do`
- `04_heterogeneity_2006_2024_noER.do`
- `05_mediation_2006_2024_noER.do`
- `06_moderation_2006_2024_noER.do`
- `07_figures_2006_2024_noER.do`
- `outputs/logs/master_2006_2024_noER.log`

## New empirical outputs

- `outputs/tables/table2_descriptive_statistics_2006_2024_noER.rtf`
- `outputs/tables/table2_descriptive_statistics_2006_2024_noER.csv`
- `outputs/tables/table3_baseline_2006_2024_noER.rtf`
- `outputs/tables/table4_robustness_iv_2006_2024_noER.rtf`
- `outputs/tables/table4_iv_diagnostics_2006_2024_noER.csv`
- `outputs/tables/table4_iv_diagnostics_2006_2024_noER.txt`
- `outputs/tables/table5_heterogeneity_2006_2024_noER.rtf`
- `outputs/tables/table6_mechanism_2006_2024_noER.rtf`
- `outputs/tables/table7_moderation_2006_2024_noER.rtf`
- `outputs/tables/table_u_test_2006_2024_noER.csv`
- `outputs/tables/table_u_test_2006_2024_noER.txt`
- `outputs/tables/dei_conditional_turning_points_2006_2024_noER.csv`
- `outputs/tables/dei_conditional_turning_points_2006_2024_noER.txt`
- `outputs/tables/table_purified_sccd_robustness_2006_2024_noER.md`
- `outputs/tables/table_purified_sccd_robustness_2006_2024_noER.csv`
- `outputs/figures/fig_sccd_distribution_2006_2024_noER.png`
- `outputs/figures/fig_nonlinear_sccd_lnce_2006_2024_noER.png`
- `outputs/figures/fig_research_framework_2006_2024_noER.png`

## Main sample and dependent variable

- The revised workflow uses the unified 2006-2024 panel: 284 cities and 5,396 city-year rows.
- Full-control regressions use 5,366 observations because some controls have missing values.
- The raw Stata file stores the log carbon-emission measure as `CE`. The new workflow generates `lnCE = CE` and uses `lnCE` in all revised regressions.
- The available file does not include a separate raw carbon-emissions level, so the current workflow cannot verify whether the upstream transformation was `ln(carbon emissions)` or `ln(carbon emissions + 1)`. The manuscript now flags this for author confirmation.

## Results rerun successfully

- Descriptive statistics: completed.
- Baseline two-way fixed effects: completed.
- Formal inverted-U endpoint test: completed.
- Winsorized robustness and alternative dependent-variable robustness: completed.
- Supplementary IV second stage and diagnostics: completed with `ivreg2`.
- Regional heterogeneity: completed.
- Mechanism equations for OIU and GTI: completed.
- DEI moderation and POLY supplementary interaction check: completed.
- Figures for SCCD distribution and nonlinear SCCD-lnCE fit: completed.

## Key numerical changes

- Baseline full-control model: `SCCD = 8.378***`, `SCCD2 = -7.725***`, `N = 5,366`, within R-squared `0.641`.
- Baseline turning point: `0.542269`.
- Turning-point standard error and 95% CI from `nlcom`: SE `0.047030`, CI `[0.450091, 0.634446]`.
- SCCD range in the baseline estimation sample: `[0.121422, 0.822216]`; the turning point is inside the sample range.
- Left endpoint marginal effect: `6.501780`, `p < 0.001`.
- Right endpoint marginal effect: `-4.324986`, `p = 0.005196`.
- IV second-stage turning point: `0.638182`.

## DEI moderation

The revised interpretation no longer says that digital economy development brings the turning point forward.

Conditional DEI results from `dei_conditional_turning_points_2006_2024_noER.csv`:

| DEI level | DEI value | Effective linear term | Effective quadratic term | Conditional turning point | Marginal effect at mean SCCD |
|---|---:|---:|---:|---:|---:|
| Low, p25 | 0.010240 | 8.599423 | -7.492671 | 0.573856 | 3.842347 |
| Mean | 0.058702 | 7.514976 | -6.443750 | 0.583121 | 3.423857 |
| High, p75 | 0.074408 | 7.163546 | -6.103833 | 0.586807 | 3.288239 |

Interpretation: higher DEI attenuates the marginal effect and weakens the inverted-U curvature. The conditional turning point moves slightly right, so the manuscript uses the revised H4: digital economy development reshapes and attenuates the nonlinear carbon effect of multidimensional spatial coupling coordination.

## IV diagnostics and interpretation

- First-stage F for `SCCD`: `85.651`.
- First-stage F for `SCCD2`: `106.982`.
- Kleibergen-Paap rk LM underidentification statistic: `57.623`, `p < 0.001`.
- Cragg-Donald Wald F statistic: `238.010`.
- Kleibergen-Paap rk Wald F statistic: `73.061`.
- Overidentification test: not applicable because the model is exactly identified with two endogenous regressors and two excluded instruments.
- Manuscript wording was downgraded to supplementary IV evidence / mitigating endogeneity concerns, because the instruments are lagged SCCD and lagged SCCD squared.

## Purified SCCD robustness

Status: not completed.

Reason: the available Stata data file contains the composite `SCCD` variable but does not include the raw indicator-level panel or the entropy-weighting script required to reconstruct SCCD after excluding overlapping indicators. The current do-files also do not construct SCCD from source indicators. `MCCD` is identical to `SCCD` in the available data and is therefore not a valid purified proxy.

Required inputs to complete this check:

- raw indicator-level city-year panel;
- entropy-weighting script;
- indicator-to-dimension mapping;
- explicit exclusion rules for policy-term, patent-related, trade-related, fiscal-expenditure-related, and digital-overlap indicators.

## Literature added and checked

Added to the reference list:

- Hou, J., Li, W., Zhang, X., 2024. Research on the impacts of digital economy on carbon emission efficiency at China's City level. PLOS ONE 19 (9), e0308001. Verified at https://journals.plos.org/plosone/article?id=10.1371/journal.pone.0308001.
- Tan, G., Zhang, X., Xiong, S., Sun, Z., Lei, Y., Wang, H., Du, S., 2024. Assessing the impacts of urban functional form on anthropogenic carbon emissions: A case study of 31 major cities in China. Ecological Indicators 167, 112700. Verified at https://www.sciencedirect.com/science/article/pii/S1470160X24011579.
- Zhu, X., Li, D., Zhou, S., Zhu, S., Yu, L., 2024. Evaluating coupling coordination between urban smart performance and low-carbon level in China's pilot cities with mixed methods. Scientific Reports 14, 20461. Verified at https://www.nature.com/articles/s41598-024-68417-4.

## Automated checks

- No exact `ER` token remains in `revised_manuscript_2006_2024_noER.docx`.
- No `environmental regulation`, `SCCD x ER`, or `SCCD2 x ER` remains in the revised DOCX.
- No `2006-2021`, `sensitivity`, `extended 2006-2024`, or `fitted/extrapolated` remains in the revised DOCX.
- Regression tables use `lnCE` rather than `CE` as the dependent variable.
- Table 1 reports 36 indicators after deleting one policy-term frequency indicator.
- Table 7 contains only the DEI moderator and POLY supplementary check.
- POLY is described as a supplementary check, not as a supported hypothesis.
- The manuscript states that IV evidence mitigates concerns but does not fully solve endogeneity.
- The formal inverted-U check reports the turning point, confidence interval, and endpoint marginal effects.
- DOCX rendering completed with artifact-tool into `rendered_2006_2024_noER/` and was visually checked from page PNGs.

## Items requiring author review

- Confirm the upstream definition of the stored `CE` variable: `ln(carbon emissions)` or `ln(carbon emissions + 1)`.
- Provide the raw SCCD indicator-level data and entropy-weighting script if purified SCCD robustness is required for submission.
- Confirm and document how the 2022-2024 rows marked by `is_fitted` were generated.
- Review the appendix indicator table for journal formatting; it is complete but dense.
