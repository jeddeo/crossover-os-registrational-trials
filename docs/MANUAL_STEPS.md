# Manual steps and how to produce the analysis dataset

Two scripts produce the automated candidate list (137 studies); manual review then excluded 61 studies, leaving the 78
analyzed trials. The exclusions are listed, with reasons and NCT IDs, in `data/manual/manual_review_exclusions.csv`. Figure 1
counts reflect the screening as recorded during the study.

| Stage | Who | Input -> output | Files |
|---|---|---|---|
| 1 | script 1 | HemOncKB export -> 304 candidate studies | `data/interim/01_stage1_candidate_studies.csv` |
| 2 | script 2 | candidates matched to FDA notifications -> 137 candidate studies | `data/interim/02_fda_matched_studies.csv` |
| 3 | reviewers | manual review: 137 -> 78 trials | `data/manual/manual_review_exclusions.csv`, `data/interim/02_after_manual_review.csv` |
| 4 | abstractors | primary publications -> working master, 86 rows | `data/manual/abstraction_template.xlsx` (headers only) |
| 5 | second reviewer, non-abstracting authors | full re-review and random audit | - |
| 6 | script 3 | working master -> `data/Crossover_Trials_Master.xlsx` | `scripts/03_make_analysis_dataset.py` |
| 7 | notebook | analysis dataset -> manuscript tables and Figure 2 | `analysis/crossover_os_analysis.ipynb` |

## Stages 1-2: automated candidate list

`scripts/01_filter_hemonckb_studies.py` keeps phase 3, randomized, registrational studies that started enrollment on or after
2000-01-01, in a front-line or adjuvant setting (a study is dropped only if every one of its HemOncKB context rows is a later-line
setting), excluding classical hematology indications. A study with outcome rows in `study_results.csv` but no entry in
`study.table.csv` cannot be screened and is excluded; the script names each one (five on this export).

`scripts/02_match_fda_approvals.py` keeps a study when any drug in its experimental or comparator regimen appears in an FDA
notification for the same condition dated on or before the study's end date. The match is deliberately loose (it compares drug
names as text), so every candidate is then reviewed by hand. FDA approvals before 2006-01-01 are not in the FDA database and
cannot be matched. If a study's enrollment month or day is not listed in HemOncKB it is assumed to be 1 January of that year.

## Stage 3: manual review

Reviewers went through the 137 candidates and excluded 61 studies. The CSV lists each with its NCT ID and reason:

| Reason | n | Rule |
|---|---|---|
| Fails the FDA-approval rule | 41 | The experimental drug did not have an FDA approval in a **later-line** setting for the **same condition** before the end of front-line enrollment (for example "Cemiplimab not approved yet", "Pembrolizumab not approved for RCC until 2019"), or the matched notification was for a different drug of the regimen |
| No peer-reviewed OS data | 9 | No published, peer-reviewed overall-survival readout, or only an immature one |
| Later-line trial | 6 | HemOncKB's context field passed the automated filter, but the trial enrolled later-line patients |
| Biosimilar, formulation, dose or duration | 5 | Biosimilars, alternative formulations (for example subcutaneous vs intravenous), different dosages of the same agent, or different treatment durations |

Two of the 61 (PERSEPHONE and DESTINY-Breast03) are not among the 137 on the current export, because script 1 already drops them
(PERSEPHONE is not registrational; every context row of DESTINY-Breast03 is later-line). So 59 of the 137 candidates are removed,
leaving 78. Script 2 applies the list and writes `data/interim/02_after_manual_review.csv`. None of the 78 analyzed trials is in
the list.

## Stage 4: data abstraction

Abstractors enter one row per comparison into the `Final Study List` tab of the working master (headers in
`data/manual/abstraction_template.xlsx`), using the peer-reviewed primary trial publication, its supplement, and the latest OS
readout, in that order of priority. HemOncKB fields (regimen, comparator, context, enrollment dates, sponsor, registry ID) are
verified against the publication and supplemented. General rules:

- **Numerator and denominator from the same readout.** Counts of control patients who received a subsequent drug and the
  denominators they are divided by must come from the same publication and data cutoff.
- **Progressed denominator.** When only on-protocol progressions are reported, use them as the number of control patients who
  progressed.
- **Not reported.** Write `NR`. Never leave a reported quantity blank, and never estimate one.
- **Fractions.** Proportions are fractions between 0 and 1, not percentages.
- **Split rows.** A trial with two experimental arms (CheckMate 067, CheckMate 227, EMBARK, IMvigor130) or two qualifying
  crossover drugs (CLEAR, RENOTORCH, COMBI-AD, COMBI-v) gets two rows, one per arm or drug.
- **Drug class.** Assigned from the experimental agent that carries the qualifying later-line approval, recorded in
  `Specific Treatment`. Every value must be in the notebook's `SPECIFIC_TREATMENT_TO_CLASS` map; the notebook stops on an
  unrecognized one.

Columns the analysis uses (see `data/Data_Dictionary.csv` for all of them):

| Column | What to enter |
|---|---|
| `study`, `regimen`, `comparator`, `context`, `enrollment`, `start`, `condition_x`, `sponsor_type` | From HemOncKB, checked against the publication |
| `Specific Treatment` | The experimental agent carrying the qualifying approval (drives drug class) |
| `ITT_Control_Num` | Control-arm participants in the intention-to-treat population |
| `Total_Control_Pts_Progressed` | Control patients who progressed (on-protocol progressions if that is all that is reported) |
| `Total_Control_Pts_Receiving_Subsequent_Therapy`, `Total_Control_Pts_With_EFS_or_PFS_Event` | The other two denominators |
| `Num control pts getting exact drug post-protocol` | Control patients who later received the exact experimental drug |
| `Num control pts getting post-protocol drug in class` | Control patients who later received any drug in the experimental drug's class (judged case by case) |
| `Frac_ctrl_pts_getting_*` | Computed in the working master as numerator / denominator for the six numerator-denominator pairs; `NR` if either part is missing. A fraction above 1 is kept (Supplementary Table S1 lists them) and is handled by the notebook |
| `perc_pts_WesternEurope`, `perc_pts_US`, `perc_pts_Europe`, `perc_pts_NorthAmerica` | Fraction of enrolled participants from each region |
| `perc_pts_NorthAmerica_Europe` | Computed: North America plus Europe (Western enrollment) |
| `Combined_OS_HR` and its 95% CI limits | The overall-survival hazard ratio from the primary publication or, when later, the latest OS readout (final analysis when in doubt) |
| `Combined_PFS_HR` | The hazard ratio for progression-free, disease-free or event-free survival, whichever the trial reports |

Geographic enrollment rules:

- **Western enrollment** is North America plus Europe. The secondary metrics are North America (US, with or without Canada, if
  North America was not reported) and Europe (Western Europe if Europe was not reported).
- When a trial reports discrete countries rather than regions, these count as Western Europe: Germany, Italy, Spain, France, United
  Kingdom, Belgium, Netherlands, Austria, Portugal and Greece. No other Western European country is reported individually.
- If a trial reports only a broader grouping, use the closest available one. This applies to three trials: ECHELON-1 (the Americas,
  including Latin America), KEYNOTE-811 and KEYNOTE-859 (groupings that included Israel and Australia).

## Stage 5: verification

A second member of the study team re-reviewed every abstracted field. Non-abstracting authors then audited a random sample of 15
trials across 20 manually entered fields (300 fields): 22 discrepancies (7.3%), of which 6 were fields first recorded as missing.
One discrepancy was a mislabeled result (EFS reported where DFS was meant); on full review it was the only such case. All
discrepancies were reconciled before the analysis.

## Producing the analysis dataset

When the working master is complete, copy it to `data/manual/working_master.xlsx` (gitignored) and run

```bash
python scripts/03_make_analysis_dataset.py
```

The script reads the master's cached values (it never edits the master), writes `data/Crossover_Trials_Master.xlsx` and `.csv`,
the data dictionary and the abstraction template, and checks that its trial list equals the output of manual review
(`data/interim/02_after_manual_review.csv`). Excel errors become `NR`, cell comments are dropped, institutional proxy links become
public links, and columns that are empty in every row are removed. It prints a count for each step.

## Data version

The HemOncKB tables are the HemOnc Dataverse release with README version 1.0 (last updated 2025-06-28), downloaded 2025-09-06.
