"""Stage 1: filter the HemOncKB export to candidate front-line studies.

Keeps registrational, phase 3 studies in non-classical-hematology conditions that
started enrollment on or after 2000-01-01, and drops studies whose every
`context` row is a later-line setting.

Reads : data/raw/study_results.csv, study.table.csv, condition.table.csv
Writes: data/interim/01_stage1_candidate_studies.csv  (one row per study)
Run   : python scripts/01_filter_hemonckb_studies.py
"""
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
INTERIM = ROOT / "data" / "interim"
INTERIM.mkdir(parents=True, exist_ok=True)

study_results = pd.read_csv(RAW / "study_results.csv")
studies = pd.read_csv(RAW / "study.table.csv")
conditions = pd.read_csv(RAW / "condition.table.csv")

# A study with outcome rows but no entry in study.table.csv has no phase, start date or registration flag,
# so it cannot be screened; each one is excluded and named.
_in_table = set(studies['study'].astype(str).str.strip())
_orphans = sorted(set(study_results['study'].astype(str).str.strip()) - _in_table)
for _name in _orphans:
    print(f"Excluded: {_name} has outcome rows in study_results.csv but no entry in study.table.csv")
study_results = study_results[~study_results['study'].astype(str).str.strip().isin(_orphans)]


filtered_conditions = conditions[~conditions['condition_type'].str.contains('classical hematology', case=False, na=False)]

cancerCodes = filtered_conditions['condition_cui']

studies['start_converted'] = pd.to_datetime(studies['start'], errors='coerce', format='%Y-%m-%d')
studies['end_converted'] = pd.to_datetime(studies['end'], errors='coerce', format='%Y-%m-%d')
# Handle 'start'
mask_start = studies['start_converted'].isna()
studies.loc[mask_start, 'start_converted'] = pd.to_datetime(
    studies.loc[mask_start, 'start'].str[:4] + '-01-01', errors='coerce'
)
# Handle 'end'
mask_end = studies['end_converted'].isna()
studies.loc[mask_end, 'end_converted'] = pd.to_datetime(
    studies.loc[mask_end, 'end'].str[:4] + '-01-01', errors='coerce'
)
start_date = pd.Timestamp('2000-01-01')
studies['start'] = studies['start_converted']
studies['end'] = studies['end_converted']

# Define filter masks
mask_right_condition = studies['condition_cui'].isin(cancerCodes)
mask_registrational = studies['reg_study'].eq(True)
mask_phase3 = studies['phase'].eq('Phase 3')
mask_timeframe = studies['start'] >= start_date

# Exclude studies using only study_results['context'] text.
# A study is excluded only if all of its context rows are later-line/relapsed/refractory.
context_column = next((c for c in study_results.columns if str(c).strip().lower() == 'context'), None)
later_line_pattern = r'\b(?:second[-\s]?line|2l|third[-\s]?line|3l|relapsed|refractory|subsequent)\b'
if context_column:
    context_df = study_results[['study', context_column]].copy()
    context_df['study_norm'] = context_df['study'].astype(str).str.strip()
    context_df['context_text'] = context_df[context_column].fillna('').astype(str).str.lower().str.strip()
    context_df = context_df[context_df['study_norm'].ne('')]
    context_df['is_later_line'] = context_df['context_text'].str.contains(later_line_pattern, na=False, regex=True)

    has_non_later_context = context_df.groupby('study_norm')['is_later_line'].apply(lambda s: (~s).any())
    exclude_context_study_ids = set(has_non_later_context.index[~has_non_later_context])

    mask_context_setting = ~studies['study'].astype(str).str.strip().isin(exclude_context_study_ids)
else:
    mask_context_setting = pd.Series([True] * len(studies), index=studies.index)

# Optional upfront-indication filter (only if a compatible column exists)
upfront_column = next((c for c in ['upfront_indication', 'upfront', 'indication_line'] if c in studies.columns), None)
if upfront_column:
    upfront_vals = studies[upfront_column].astype(str).str.strip().str.lower()
    mask_upfront = upfront_vals.isin({'true', '1', 'yes', 'y', 'upfront', 'first line', 'first-line', '1l'})
else:
    mask_upfront = pd.Series([True] * len(studies), index=studies.index)

# Print exclusion counts by reason (each criterion evaluated independently)
print("\nExclusions by criterion (independent counts):")
print(
    f"- Not right condition: rows={(~mask_right_condition).sum()} | "
    f"unique studies={studies.loc[~mask_right_condition, 'study'].nunique()}"
)
print(
    f"- Non-registrational: rows={(~mask_registrational).sum()} | "
    f"unique studies={studies.loc[~mask_registrational, 'study'].nunique()}"
)
print(
    f"- Not Phase 3: rows={(~mask_phase3).sum()} | "
    f"unique studies={studies.loc[~mask_phase3, 'study'].nunique()}"
)
if upfront_column:
    print(
        f"- Not upfront indication ({upfront_column}): rows={(~mask_upfront).sum()} | "
        f"unique studies={studies.loc[~mask_upfront, 'study'].nunique()}"
    )
else:
    print("- Not upfront indication: not evaluated (no upfront column found)")
print(
    f"- Not in time frame (start < {start_date.date()} or missing): rows={(~mask_timeframe).sum()} | "
    f"unique studies={studies.loc[~mask_timeframe, 'study'].nunique()}"
)
if context_column:
    print(
        f"- Later-line context excluded (all context rows later-line in study_results.{context_column}): rows={(~mask_context_setting).sum()} | "
        f"unique studies={studies.loc[~mask_context_setting, 'study'].nunique()}"
    )
else:
    print("- Later-line context excluded: not evaluated (no context column found in study_results)")

# Waterfall view: excluded at each step after prior filters are applied
step_masks = [
    ('right condition', mask_right_condition),
    ('registrational', mask_registrational),
    ('Phase 3', mask_phase3),
    ('upfront indication', mask_upfront),
    ('time frame', mask_timeframe),
    ('context (exclude only all-later-line studies)', mask_context_setting),
]

remaining = studies.copy()
print("\nWaterfall exclusion summary:")
for step_name, step_mask in step_masks:
    before_rows = len(remaining)
    before_unique = remaining['study'].nunique()
    remaining = remaining[step_mask.loc[remaining.index]]
    after_rows = len(remaining)
    after_unique = remaining['study'].nunique()
    excluded_rows = before_rows - after_rows
    excluded_unique = before_unique - after_unique
    print(
        f"- Excluded at {step_name}: rows={excluded_rows}, unique={excluded_unique} | "
        f"Remaining: rows={after_rows}, unique={after_unique}"
    )

# Define Studies of Interest: Registrational, Phase 3, after 2000, non-heme conditions
# Keep behavior aligned with previous logic (upfront/context are only applied if recognized columns exist).
selected_studies = studies[
    mask_registrational &
    mask_phase3 &
    mask_timeframe &
    mask_right_condition &
    mask_upfront &
    mask_context_setting
]

print(f"Number of studies: {studies['study'].nunique()}")

merged_studies = study_results.merge(
    selected_studies,
    left_on='study',
    right_on='study',
    how='left'
)

merged_studies = merged_studies[merged_studies['reg_study'].notna()]
print(f"Number of unique studies in merged_studies: {merged_studies['study'].nunique()}")

# One row per study: the first row of the study's results in the export.
unique_by_study = merged_studies.drop_duplicates().drop_duplicates(subset=['study'], keep='first')
unique_by_study.to_csv(INTERIM / "01_stage1_candidate_studies.csv", index=False, lineterminator="\n")
print(f"Wrote {len(unique_by_study)} studies to data/interim/01_stage1_candidate_studies.csv")
