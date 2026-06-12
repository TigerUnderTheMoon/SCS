# Sustainability Submission Readiness Audit

Generated: 2026-06-03

## Package Audited

- Manuscript DOCX: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\manuscript_sustainability_deep.docx`
- Manuscript PDF: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\manuscript_sustainability_deep.pdf`
- QA JSON: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\sustainability_deep_checks.json`
- Render QA: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\rendered_pdf_latest\contact_sheet.png`
- Author intake form: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\author_metadata_intake.md`
- Final checklist: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\final_submission_checklist.md`

## Official Requirement Sources Checked

- MDPI Sustainability Instructions for Authors: https://www.mdpi.com/journal/sustainability/instructions
- MDPI Sustainability manuscript templates and article structure guidance are treated as the current journal-facing source for this audit.

## Current Status

Status: SUBMISSION-PREPARED, AUTHOR-METADATA-BLOCKED.

The manuscript package has been rebuilt for Sustainability and integrates the four requested smart-city and carbon-emissions papers. The generated QA file reports `passed: true`, 33 references, 153 abstract words, 7 tables, required back-matter headings, observed-main-sample wording, and no forbidden `mediation effect` phrase. The final PDF was rendered to 18 PNG pages and visually checked through the latest contact sheet and dense table/reference pages.

## Requirement Checklist

| Item | Evidence | Status |
| --- | --- | --- |
| Target journal changed to Sustainability | Title page and section structure in generated DOCX/PDF | OK |
| Four requested papers integrated | References 16, 17, 18, and 27 in the generated manuscript | OK |
| Abstract near MDPI length expectation | QA JSON reports 153 words | OK |
| Keywords present | DOCX keyword line contains six keywords | OK |
| Main article sections present | Introduction, Materials and Methods, Results, Discussion, Conclusions | OK |
| Back-matter headings present | Author Contributions, Funding, IRB, Consent, Data Availability, Conflicts, Supplementary Materials, References | OK |
| Observed sample claim-safe | QA JSON reports observed-main-sample wording and no unqualified 2006-2024 main-sample wording | OK |
| Fitted 2022-2024 rows not overstated | QA JSON reports no forbidden real-records phrase | OK |
| Mechanism wording claim-safe | QA JSON reports no `mediation effect` phrase | OK |
| PDF generated | QA JSON reports PDF created through Word COM | OK |
| Visual layout checked | 18 rendered PNG pages and contact sheet under `rendered_pdf_latest/` | OK |
| Author CRediT statement final | Current text says author team must finalize named roles | BLOCKED |
| Funding statement final | Current text says funding metadata were unavailable | BLOCKED |
| Conflict-of-interest declaration final | Current text says formal declaration is still required | BLOCKED |
| Data/code availability wording final | Current wording is cautious but source-license confirmation remains required | BLOCKED |
| Cover letter prepared | Draft template exists with placeholders; final version still needs author metadata | BLOCKED |

## Required Author Inputs

1. Named author list and CRediT roles.
2. Funding statement with grant numbers, or a formal no-external-funding statement.
3. Conflict-of-interest declaration.
4. Final data/source-license wording for raw data, processed city-level data, code, tables, and figures.
5. Cover letter content: author names and affiliations, final corresponding author block, brief contribution statement, originality/no-under-review declaration, suggested reviewers if the author team wants to provide them, and any exclusions.

## Next Concrete Action

After the five author inputs above are supplied in `author_metadata_intake.md`, update the generator back-matter strings, regenerate `outputs/sustainability_restructure_20260603_deep/`, rerun the QA JSON and PDF render checks, and finalize the cover letter for the MDPI submission system.
