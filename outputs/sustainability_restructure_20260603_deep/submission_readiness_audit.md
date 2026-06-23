# Sustainability Submission Readiness Audit

Generated: 2026-06-14

## Package Audited

- Manuscript DOCX: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\manuscript_sustainability_deep.docx`
- Manuscript PDF: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\manuscript_sustainability_deep.pdf`
- QA JSON: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\sustainability_deep_checks.json`
- Render QA: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\rendered_pdf_latest\contact_sheet.png`
- Highlights: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\submission_highlights.md`
- Graphical abstract: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\graphical_abstract.png`
- Supplementary README: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\supplementary_materials\README.md`
- DOI audit: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\reference_doi_audit.csv`
- MDPI requirements audit: `D:\Workplace\SCS\outputs\sustainability_restructure_20260603_deep\submission_requirements_audit.md`

## Current Status

Status: STRUCTURAL-QA-PASSED, AUTHOR-METADATA-BLOCKED.

The package now addresses the Sustainability-format checklist items that can be completed from repository evidence. The QA JSON reports `passed: true`, `submission_ready: false`, and `submission_status: STRUCTURAL_QA_PASSED_AUTHOR_METADATA_BLOCKED`. It also reports a 130-word structured abstract, 6 keywords, 33 references, 9 manuscript tables, required ethics and back-matter headings, AI-use declaration, observed-main-sample wording, no forbidden `mediation effect` phrase, no manuscript-facing generator trace terms, no main-text internal process terms, and zero long sentences at the audit threshold. The regenerated PDF was rendered to 32 PNG pages and visually checked through the contact sheet plus selected dense pages.

## Requirement Checklist

| Item | Evidence | Status |
| --- | --- | --- |
| Structured abstract | Background, Methods, Results, Conclusions in DOCX/PDF; QA `structured_abstract: true` | OK |
| Article type | `Article` appears in the title block | OK |
| Title-page author metadata | QA `title_page_author_metadata_present: false` | BLOCKED |
| Keywords | 6 keywords | OK |
| Highlights | 5 items in `submission_highlights.md` | OK |
| Graphical abstract | `graphical_abstract.png`, 2400 x 1440 px, ~300 dpi | OK |
| Ethics statements | Institutional Review Board and Informed Consent statements are present | OK |
| Author Contributions heading | Present, but final CRediT roles require author names | BLOCKED |
| Funding heading | Present, but grant/no-funding statement requires author confirmation | BLOCKED |
| Acknowledgments heading | Present, but final acknowledgement wording requires author confirmation | BLOCKED |
| Conflicts heading | Present, but formal conflict declaration requires author confirmation | BLOCKED |
| Generator/process traces removed from manuscript | QA `generator_trace_terms_found: []`; QA `internal_process_terms_found_main_text: []` | OK |
| VIF table | `table08_vif_observed_2006_2021.csv`; Table S1 in manuscript | OK |
| Fixed-effect joint tests | `table08_fixed_effects_joint_tests_observed_2006_2021.csv`; Table S2 in manuscript | OK |
| References in MDPI-style numbered list | 33 numbered references in manuscript | OK |
| DOI audit | 27 direct resolver successes, 5 DOI redirects to publisher pages with automated-access 403, 1 web/data source without DOI | OK WITH NOTE |
| AI assistance statement | Dedicated declaration section in manuscript | OK |
| Supplementary materials classified | `main_analysis/`, `sensitivity_analysis/`, and README | OK |
| Long-sentence pass | QA reports `long_sentences_ge_35_words: 0` before references | OK |
| Visual layout | 32 rendered page PNGs and contact sheet under `rendered_pdf_latest/` | OK |

## Required Author Inputs

1. Named author list and CRediT roles.
2. Title-page affiliations, ORCID IDs if used, and corresponding-author metadata.
3. Funding statement with grant numbers, or a formal no-external-funding statement.
4. Final acknowledgments text, or confirmation that no acknowledgments are required.
5. Formal conflict-of-interest declaration.
6. Final data/source-license wording for raw data, processed city-level data, code, tables, and figures.
7. Confirmation that the AI-use declaration wording is acceptable for the target journal.
8. Cover letter content: author names and affiliations, final corresponding author block, contribution statement, originality/no-under-review declaration, suggested reviewers if desired, and exclusions if any.

## Next Concrete Action

After the author inputs above are supplied in `author_metadata_intake.md`, update the back-matter strings in `build_sustainability_observed_manuscript.py`, regenerate `outputs/sustainability_restructure_20260603_deep/`, rerun the QA JSON and PDF render checks, and finalize the cover letter for the MDPI submission system.
