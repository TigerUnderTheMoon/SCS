# SCS Content Revision Log

Source file: `revised_manuscript_2006_2024_noER.docx`  
Revised file: `revised_manuscript_2006_2024_noER_SCS_content_revised.docx`

## Task Completion Checklist

1. Data description in Section 4.1: completed. The sample is stated as 284 cities, 2006-2024, and 5,396 city-year observations. Statements implying 2022-2024 are fitted, predicted, uncertain, or needing documentation before submission were removed from the manuscript.
2. lnCE dependent-variable definition in Section 4.2: completed. The manuscript now states that lnCE is ln(CE_raw + 1), and that the Stata variable stored as CE is renamed lnCE without an additional log transformation.
3. Table 4 MCCD/MCCD2 robustness column: completed. Existing project results from `outputs/tables/table04_robustness_endogeneity_full_2006_2024.rtf` were used: MCCD = 8.378*** (1.388), MCCD2 = -7.725*** (1.665), N = 5366, within R-squared = 0.641. No coefficients were invented.
4. Mechanism-analysis wording downgrade: completed. Strong causal wording was replaced with mechanism-equation language in Sections 3.2, 3.3, 5.4, and the conclusion. No indirect effects or mediation confidence intervals are claimed.
5. SCS-style abstract: completed. The abstract is under 250 words, reduces internal variable-name stacking, and emphasizes sustainable urban governance and the shift from expansion-oriented integration to efficiency-oriented governance.
6. Introduction contribution paragraph: completed. The contribution text now distinguishes conceptual, empirical, and mechanism/boundary contributions in integrated prose.
7. Reference DOI and format cleanup: completed for the manuscript reference list. DOI checks used Crossref and publisher pages. No unverifiable DOI was inserted.
8. SCCD calculation formula in Section 4.2: completed. The manuscript now describes normalization, entropy weights, subsystem scores, equal-weight comprehensive development index, coupling degree, and coordination degree.
9. Table 7 DEI moderation explanation compression: completed. The interpretation now keeps the key signs, conditional turning points, marginal effects, and one concise POLY paragraph.
10. H4 theory alignment: completed. Section 3.4 now explains why DEI attenuates expansion-related carbon costs and flattens the curvature without claiming that DEI brings the turning point forward.

## Remaining Issues

- None of the 10 requested items is blocked by missing data or missing regression output.
- The MCCD/MCCD2 column was restored from existing generated results rather than by rerunning Stata in this pass.
