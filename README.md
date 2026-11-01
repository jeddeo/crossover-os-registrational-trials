# Control-arm crossover, Western enrollment, and overall survival benefit in front-line registrational oncology trials

Data, code and manuscript tables for a cross-sectional analysis of 78 phase 3 registrational cancer trials (86 comparisons) in
which the experimental drug already had an FDA-approved later-line indication for the same disease. The analysis asks whether
control-arm crossover (control patients who later received the experimental drug, or a drug of its class) or Western enrollment
(participants enrolled in North America or Europe) is associated with the overall-survival (OS) hazard ratio.

**Authors.** Jeddeo Paul, Syed Muhammad Ismail Shah, Timothy Joseph Davis, Jeremy Warner, Ali Raza Khaki, James C Dickerson.
**Manuscript.** Under review.
**Contact.** Corresponding author James C. Dickerson, <jcdicker@stanford.edu>.

## Overview

Trial selection starts from the HemOnc.org Knowledge Base (HemOncKB) and the FDA Oncology/Hematologic Malignancies Approval
Notifications. Two scripts produce the automated candidate list (137 studies); manual review then excluded 61 studies, leaving the
78 analyzed trials (the exclusions, with reasons, are in [`data/manual/manual_review_exclusions.csv`](data/manual/manual_review_exclusions.csv)).
Data abstraction and verification are manual and are documented in [`docs/MANUAL_STEPS.md`](docs/MANUAL_STEPS.md). The analysis
notebook then reads the clean analysis dataset and writes the manuscript tables and Figure 2. Figure 1 counts reflect the
screening as recorded during the study.

```
 HemOncKB export (data/raw)                      FDA approval notifications (data/raw)
          |                                                       |
 [1] scripts/01_filter_hemonckb_studies.py   automated           |
          |  7,076 studies -> 304 candidates                      |
          +-------------------------+-----------------------------+
                                    |
 [2] scripts/02_match_fda_approvals.py   automated
          |  304 -> 137 candidate studies  (data/interim/02_fda_matched_studies.csv)
          |
 [3] manual review                       MANUAL   61 excluded (data/manual/manual_review_exclusions.csv)  -> 78 trials
 [4] data abstraction                    MANUAL   (working master, private; template in data/manual/)  -> 86 rows
 [5] second review + random audit        MANUAL   (15 trials x 20 fields)
          |
 [6] scripts/03_make_analysis_dataset.py   automated: working master -> data/Crossover_Trials_Master.xlsx / .csv
          |
 [7] analysis/crossover_os_analysis.ipynb  automated: tables T1-T2, S1-S10 and Figure 2 -> outputs/
```

| Stage | Type | What happens | Count |
|---|---|---|---|
| 1 | Automated | Phase 3, randomized, registrational, started on or after 2000-01-01, front-line or adjuvant, not classical hematology | 7,076 in; 304 remain |
| 2 | Automated | Experimental or comparator drug has an FDA notification for the same condition on or before the study's end | 137 candidates |
| 3 | Manual | Review of the candidates (FDA-approval rule, OS data, front-line setting, formulation/dose/duration) | 61 excluded; 78 trials |
| 4 | Manual | Abstraction into the working master | 86 rows (78 trials) |
| 5 | Manual | Full second review and a random audit of 15 trials x 20 fields | all discrepancies reconciled |
| 6 | Automated | Working master -> values-only analysis dataset | 86 rows |
| 7 | Automated | Notebook -> manuscript tables and Figure 2 | - |

## Repository structure

```
README.md  LICENSE  DATA_LICENSE.md  CITATION.cff  requirements.txt  .python-version  run_all.sh
data/
  raw/       HemOncKB tables and the FDA notifications table, unmodified
  interim/   outputs of scripts 1 and 2 (committed so the filtering can be checked)
  manual/    manual_review_exclusions.csv, abstraction_template.xlsx
  Crossover_Trials_Master.xlsx / .csv   the analysis dataset (the notebook's only input)
  Data_Dictionary.csv
scripts/     01_filter_hemonckb_studies.py  02_match_fda_approvals.py  03_make_analysis_dataset.py
analysis/    crossover_os_analysis.ipynb   (saved with its outputs)
docs/        MANUAL_STEPS.md
outputs/     Manuscript_Tables.xlsx  Manuscript_Tables.docx  figures/Figure_2.png  figures/Figure_2.pdf
```

## Reproduce

Python 3.13.7.

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
./run_all.sh
```

`run_all.sh` runs stages 1, 2, 6 and 7. Stage 6 needs the completed working master, which is private (it holds the reviewers'
check tabs and comments) and is not in this repository; without `data/manual/working_master.xlsx` the script skips stage 6 and the
notebook runs on the committed `data/Crossover_Trials_Master.xlsx`. Each script can also be run alone:

```bash
python scripts/01_filter_hemonckb_studies.py     # reads data/raw, writes data/interim
python scripts/02_match_fda_approvals.py         # reads data/interim and data/raw, writes data/interim
python scripts/03_make_analysis_dataset.py PATH  # PATH = the completed working master
jupyter nbconvert --to notebook --execute --inplace analysis/crossover_os_analysis.ipynb
```

Every script prints its counts, and rerunning any of them gives byte-identical files. The notebook asserts the counts the
current dataset must produce (the `EXPECTED` dictionary in its configuration cell), so a change in the data stops the run.

## Outputs and the manuscript

| Output | Manuscript |
|---|---|
| `outputs/Manuscript_Tables.xlsx` tab `T1` / `T2` | Table 1 (crossover vs OS) / Table 2 (Western enrollment vs OS) |
| tab `S1` | Supplementary Table S1: comparisons with crossover rates above 100% (progressed denominator) |
| tab `S2` / `S3` | Supplementary Table S2: summary of included studies / S3: all included studies |
| tab `S4` / `S5` | S4: crossover vs log OS HR, all models / S5: between-trial heterogeneity |
| tab `S6` | S6: precision bounds (largest crossover or Western enrollment effect compatible with the data) |
| tab `S7` | S7: Western enrollment vs log OS HR, all models |
| tab `S8` | S8: all 289 stratified correlations |
| tab `S9` / `S10` | S9: one comparison per trial / S10: crossover capped at 100% |
| `outputs/Manuscript_Tables.docx` | the same twelve tables, one per page, formatted for pasting into Word |
| `outputs/figures/Figure_2.png`, `.pdf` | Figure 2 (PNG at 600 dpi; PDF is vector) |

Figure 1 (study flow) and Table 3 (minimum reporting elements) are not produced by code.

## Data sources

- **HemOnc.org Knowledge Base (HemOncKB)** - <https://hemonc.org>. The `study.table.csv`, `study_results.csv` and
  `condition.table.csv` tables from the HemOnc Dataverse release (README version 1.0, last updated 2025-06-28); files dated
  2025-09-06 on disk. The scripts read only these three tables.
- **FDA Oncology/Hematologic Malignancies Approval Notifications** -
  <https://www.fda.gov/drugs/resources-information-approved-drugs/oncology-cancerhematologic-malignancies-approval-notifications>.
  Notifications from 2006 onward, compiled into `data/raw/FDA_Notifications.xlsx` (file dated 2026-02-16). The FDA database starts
  in 2006, so earlier approvals cannot be matched.

## Licenses

Code: MIT (`LICENSE`). Abstracted dataset: CC BY 4.0. The HemOncKB tables are redistributed under HemOncKB's terms. See
[`DATA_LICENSE.md`](DATA_LICENSE.md). To cite, see `CITATION.cff`.
