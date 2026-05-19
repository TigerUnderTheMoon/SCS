# Manuscript Change Log

Updated: 15 May 2026

## Files created

- `revised_manuscript.docx`
- `manuscript_change_log.md`
- `remaining_submission_checks.md`

## Main revisions

- Reframed the manuscript so the main empirical analysis is based on the observed 2006-2021 panel of 284 cities.
- Treated the extended 2006-2024 panel only as a sensitivity check because the 2022-2024 rows are fitted/extrapolated observations.
- Updated the abstract, data section, empirical strategy, results discussion, conclusions, policy implications, and limitations to reflect the observed-sample framing.
- Updated Tables 2-7 to report observed 2006-2021 results as the main results.
- Revised the discussion of the extended 2006-2024 results so they are described as sensitivity evidence, not the main sample.

## Moderator revisions

- Kept digital economy development as the only robust moderator in the revised interpretation.
- Reduced environmental regulation and polycentricity to supplementary interaction checks.
- Removed claims that environmental regulation and polycentricity are supported moderators, because the extended-sample interaction terms are not statistically significant.
- Resolved the environmental-regulation overlap problem by explaining that ER is not interpreted as an independent supported moderator because environmental regulation intensity is already included in the SCCD indicator system.

## Mechanism wording

- Renamed "mediation analysis" as "mechanism analysis" throughout the main manuscript.
- Reworded industrial upgrading and green technological innovation as mechanism channels rather than formal mediators.
- Did not add formal indirect-effect estimates, confidence intervals, or new mediation statistics.

## Table checks

- Verified Table 4 column (3) against `03_robustness_endogeneity.do` and the exported observed-sample robustness table: the alternative explanatory variable specification uses `MCCD` and `MCCD2`, not `SCCD` and `SCCD2`.
- Revised Table 4 row labels and notes to make the MCCD/MCCD2 specification explicit.
- Checked Table 1 for duplicated indicators and removed duplicated manuscript-table entries for "Employees in water conservancy, environment and public facilities management"; one ecological-response entry is retained.
- Revised the text to describe the indicator system as 37 indicators after removing duplicated manuscript-table entries.

## Placeholder and declaration handling

- Removed unresolved placeholder notes from the main manuscript.
- Removed unresolved author-specific declarations from the main manuscript.
- Moved unresolved declaration and submission items to `remaining_submission_checks.md`.

## Validation

- Regenerated `revised_manuscript.docx` from `_revise_manuscript_editorial.py`.
- Rendered the revised DOCX to `rendered_revised_manuscript`.
- Checked the rendered pages for major clipping and table separation issues.
- Confirmed by DOCX text scan that the manuscript contains no unresolved placeholder notes and no remaining main-text "mediation analysis", "mediating variables", or "mediator" wording.
