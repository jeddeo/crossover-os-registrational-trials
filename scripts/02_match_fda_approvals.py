"""Stage 2: match candidate studies to prior FDA approvals in a later line.

Keeps studies whose experimental or comparator drug had an FDA notification, for the
same condition, dated on or before the study's end date.  The FDA Oncology/Hematologic
Malignancies Approval Notifications database starts in 2006, so earlier approvals
cannot be matched.

The automated match is deliberately loose.  Manual review (eligibility, and confirmation that the matched notification
is a prior later-line approval of the experimental drug) then excludes the studies listed in
data/manual/manual_review_exclusions.csv.

Reads : data/interim/01_stage1_candidate_studies.csv, data/raw/FDA_Notifications.xlsx,
        data/manual/manual_review_exclusions.csv
Writes: data/interim/02_fda_matched_studies.csv      (automated candidates)
        data/interim/02_after_manual_review.csv      (the analyzed trials)
Run   : python scripts/02_match_fda_approvals.py
"""
import re
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
INTERIM = ROOT / "data" / "interim"
MANUAL = ROOT / "data" / "manual"

unique_studies_df = pd.read_csv(INTERIM / "01_stage1_candidate_studies.csv", low_memory=False)

# Load the All_FDA_Notifications_06_to_26 tab from FDA_Notifications.xlsx
fda_notifications_df = pd.read_excel(ROOT / "data" / "raw" / "FDA_Notifications.xlsx",
                                     sheet_name='All_FDA_Notifications_06_to_26')

print(f"Loaded {len(unique_studies_df)} unique studies")
print(f"Loaded {len(fda_notifications_df)} FDA notifications")

# Best-effort expansions for common oncology regimen abbreviations found in the dataset.
REGIMEN_EXPANSIONS = {
    'NALIRIFOX': ['liposomal irinotecan', 'irinotecan', 'fluorouracil', 'leucovorin', 'oxaliplatin'],
    'CAPEOX': ['capecitabine', 'oxaliplatin'],
    'ISARVD': ['isatuximab', 'lenalidomide', 'bortezomib', 'dexamethasone'],
    'DARARVD': ['daratumumab', 'lenalidomide', 'bortezomib', 'dexamethasone'],
    'RVD': ['lenalidomide', 'bortezomib', 'dexamethasone'],
    'ABCP': ['atezolizumab', 'bevacizumab', 'carboplatin', 'paclitaxel'],
    'DARAVTD': ['daratumumab', 'bortezomib', 'thalidomide', 'dexamethasone'],
    'VTD': ['bortezomib', 'thalidomide', 'dexamethasone'],
    'VRCAP': ['bortezomib', 'rituximab', 'cyclophosphamide', 'doxorubicin', 'prednisone'],
    'RCHOP': ['rituximab', 'cyclophosphamide', 'doxorubicin', 'vincristine', 'prednisone'],
    'CHOP': ['cyclophosphamide', 'doxorubicin', 'vincristine', 'prednisone'],
    'RCVP': ['rituximab', 'cyclophosphamide', 'vincristine', 'prednisone'],
    'CVP': ['cyclophosphamide', 'vincristine', 'prednisone'],
    'BVCHP': ['brentuximab vedotin', 'cyclophosphamide', 'doxorubicin', 'prednisone'],
    'ABVD': ['doxorubicin', 'bleomycin', 'vinblastine', 'dacarbazine'],
    'AAVD': ['brentuximab vedotin', 'doxorubicin', 'vinblastine', 'dacarbazine'],
    'ABVEPC': ['doxorubicin', 'bleomycin', 'vincristine', 'etoposide', 'prednisone', 'cyclophosphamide'],
    'COPADM': ['cyclophosphamide', 'vincristine', 'prednisone', 'doxorubicin', 'methotrexate'],
    'RCOPADM': ['rituximab', 'cyclophosphamide', 'vincristine', 'prednisone', 'doxorubicin', 'methotrexate'],
    'FOLFIRI': ['irinotecan', 'fluorouracil', 'leucovorin'],
    'FOLFOX4': ['oxaliplatin', 'fluorouracil', 'leucovorin'],
    'TCH': ['docetaxel', 'carboplatin', 'trastuzumab'],
    'THP': ['docetaxel', 'trastuzumab', 'pertuzumab'],
    'TH': ['docetaxel', 'trastuzumab'],
    'ACTH': ['doxorubicin', 'cyclophosphamide', 'paclitaxel', 'trastuzumab'],
    'ACTHP': ['doxorubicin', 'cyclophosphamide', 'paclitaxel', 'trastuzumab', 'pertuzumab'],
    'VMP': ['bortezomib', 'melphalan', 'prednisone'],
    'MPT': ['melphalan', 'prednisone', 'thalidomide'],
    'VDC': ['vincristine', 'doxorubicin', 'cyclophosphamide'],
    'CPX351': ['cytarabine', 'daunorubicin'],
    'ADE': ['cytarabine', 'daunorubicin', 'etoposide'],
    'ATRA': ['all-trans retinoic acid', 'tretinoin'],
    'FCR': ['fludarabine', 'cyclophosphamide', 'rituximab'],
    'IFL': ['irinotecan', 'fluorouracil', 'leucovorin'],
    'ECF': ['epirubicin', 'cisplatin', 'fluorouracil'],
    'ECX': ['epirubicin', 'cisplatin', 'capecitabine'],
    'EP': ['etoposide', 'cisplatin'],
    'GC': ['gemcitabine', 'cisplatin'],
    'CF': ['cisplatin', 'fluorouracil'],
    'FC': ['fludarabine', 'cyclophosphamide'],
    'CE': ['carboplatin', 'etoposide'],
    'BR': ['bendamustine', 'rituximab'],
    'MP': ['melphalan', 'prednisone'],
    'TD': ['thalidomide', 'dexamethasone'],
}

# Treat these condition variants as equivalent during condition filtering.
CONDITION_ALIAS_GROUPS = [
    [
        'rcc',
        'ccrcc',
        'renal cell carcinoma',
        'clear cell renal cell carcinoma',
        'metastatic renal cell carcinoma',
        'advanced renal cell carcinoma',
        'renal cancer',
        'kidney cancer',
    ],
    [
        'adenocarcinoma of the pancreas',
        'pancreatic cancer',
        'pancreatic adenocarcinoma',
        'adenocarcinoma pancreas',
        'adenocarcinoma of pancreas',
        'pancreas adenocarcinoma',
    ],
    [
        'non-small cell lung cancer',
        'non small cell lung cancer',
        'nsclc',
    ],
    [
        'gastric cancer',
        'gastric adenocarcinoma',
        'adenocarcinoma of the stomach',
        'stomach adenocarcinoma',
        'gastroesophageal junction adenocarcinoma',
        'gastro-oesophageal junction adenocarcinoma',
        'gastric or gastroesophageal junction adenocarcinoma',
    ],
    [
        'esophageal cancer',
        'oesophageal cancer',
        'esophageal adenocarcinoma',
        'oesophageal adenocarcinoma',
        'gastroesophageal junction adenocarcinoma',
        'gastro-oesophageal junction adenocarcinoma',
        'gastric or gastroesophageal junction adenocarcinoma',
    ],
    [
        'waldenstrom',
        'waldenstroem',
        "waldenstrom's",
        "waldenstroem's",
        'waldenstrom macroglobulinemia',
        'waldenstroem macroglobulinemia',
        "waldenstrom's macroglobulinemia",
        "waldenstroem's macroglobulinemia",
    ],
]


def normalize_regimen_token(token):
    return re.sub(r'[^A-Z0-9]', '', str(token).upper())


def normalize_condition_text(text_value):
    return re.sub(r'[^a-z0-9]+', ' ', str(text_value).lower()).strip()


def text_has_alias(normalized_text, alias):
    alias_norm = normalize_condition_text(alias)
    if not alias_norm:
        return False
    # Token-safe containment for both single- and multi-word aliases
    return f' {alias_norm} ' in f' {normalized_text} '


def condition_matches(study_condition, fda_webpage):
    study_norm = normalize_condition_text(study_condition)
    webpage_norm = normalize_condition_text(fda_webpage)

    if not study_norm:
        return True

    matched_groups = [
        group
        for group in CONDITION_ALIAS_GROUPS
        if any(text_has_alias(study_norm, alias) for alias in group)
    ]

    if matched_groups:
        return any(
            any(text_has_alias(webpage_norm, alias) for alias in group)
            for group in matched_groups
        )

    return study_norm in webpage_norm


# Function to extract individual drugs from a free-text field
def extract_drugs_from_text(text_value):
    if pd.isna(text_value) or text_value == '':
        return []

    words = str(text_value).split()
    excluded_words = {'monotherapy', 'combination', 'and', 'or', 'with', 'plus', 'vs', 'versus'}

    drugs = []
    seen = set()

    for word in words:
        clean_word = word.strip(' ,;:.()')
        if not clean_word:
            continue
        if clean_word.lower() in excluded_words:
            continue

        # Keep original parsed token
        token_key = clean_word.lower()
        if token_key not in seen:
            seen.add(token_key)
            drugs.append(clean_word)

        # Add expanded full drug names for known regimen shorthand
        abbr_key = normalize_regimen_token(clean_word)
        for expanded in REGIMEN_EXPANSIONS.get(abbr_key, []):
            exp_key = expanded.lower()
            if exp_key not in seen:
                seen.add(exp_key)
                drugs.append(expanded)

    return drugs


# Parse study drugs from both regimen and comparator columns
def extract_study_drugs(study_row):
    regimen_drugs = extract_drugs_from_text(study_row.get('regimen', ''))
    comparator_drugs = extract_drugs_from_text(study_row.get('comparator', ''))

    # Preserve order while removing duplicates
    combined = []
    seen = set()
    for drug in regimen_drugs:
        key = drug.lower()
        if key not in seen:
            seen.add(key)
            combined.append(drug)

    for drug in comparator_drugs:
        key = drug.lower()
        if key not in seen:
            seen.add(key)
            combined.append(drug)

    return combined


# Function to check if any drug from study matches any drug in FDA notification (on or before end date)
def find_matching_notifications_before_end(study_drugs, fda_df, study_end_date, study_condition):
    matching_rows = []
    for _, fda_row in fda_df.iterrows():
        fda_drugs_str = str(fda_row['Drug']).lower() if pd.notna(fda_row['Drug']) else ''
        fda_webpage_str = str(fda_row['Webpage']).lower() if pd.notna(fda_row['Webpage']) else ''

        fda_date = pd.to_datetime(fda_row['Date'], errors='coerce')

        if pd.isna(fda_date) or pd.isna(study_end_date):
            continue
        # Check if FDA date is on or before study end date
        if fda_date > study_end_date:
            continue

        if not condition_matches(study_condition, fda_webpage_str):
            continue

        for study_drug in study_drugs:
            if study_drug.lower() in fda_drugs_str:
                matching_rows.append(fda_row)
                break
    return matching_rows


# Process for on or before end date
results_before_end = []
for idx, study_row in unique_studies_df.iterrows():
    study_drugs = extract_study_drugs(study_row)
    study_end_date = pd.to_datetime(study_row['end'], errors='coerce')
    study_condition = study_row.get('condition_x', None)

    matching_notifications = find_matching_notifications_before_end(
        study_drugs,
        fda_notifications_df,
        study_end_date,
        study_condition,
    )

    if matching_notifications:
        study_data = study_row.to_dict()
        for i, notification in enumerate(matching_notifications):
            study_data[f'FDA_Notice_{i+1}'] = notification['Webpage']
            study_data[f'FDA_Drug_{i+1}'] = notification['Drug']
            study_data[f'FDA_Indication_{i+1}'] = notification['Indication']
            study_data[f'FDA_Date_{i+1}'] = notification['Date']
        results_before_end.append(study_data)

final_df_before_end = pd.DataFrame(results_before_end)

# The export spells one study "PSMAfore"; the abstraction master has always used "PSMA Fore".
STUDY_NAME_MAP = {"PSMAfore": "PSMA Fore"}
final_df_before_end["study"] = final_df_before_end["study"].replace(STUDY_NAME_MAP)
final_df_before_end.to_csv(INTERIM / "02_fda_matched_studies.csv", index=False, lineterminator="\n")


def drop_listed(frame, csv_name, label):
    """Remove the studies named in a manual exclusion list.  Returns the remaining studies
    and the listed studies that were not in `frame`."""
    listed = pd.read_csv(MANUAL / csv_name)
    present = frame["study"].isin(listed["study"])
    absent = sorted(set(listed["study"]) - set(frame["study"]))
    kept = frame[~present].reset_index(drop=True)
    print(f"{label}: {len(listed)} listed, {int(present.sum())} present and removed -> {len(kept)} remain")
    return kept, absent


print(f"\nAutomated FDA match: {len(final_df_before_end)} candidate studies")

# Manual review: eligibility and the check that the matched FDA notification is a prior later-line approval of the
# experimental drug for the same condition.  Every excluded study is listed, with its reason, in the CSV.
listed = pd.read_csv(MANUAL / "manual_review_exclusions.csv")
present = final_df_before_end["study"].isin(listed["study"])
absent = sorted(set(listed["study"]) - set(final_df_before_end["study"]))
# Two listed studies (PERSEPHONE, DESTINY-Breast03) are not FDA-matched candidates on this export: script 1 already
# drops them (not registrational; every context row later-line).
assert not set(absent) & set(unique_studies_df["study"]), f"Listed studies that are candidates but not matched: {absent}"
analyzed = final_df_before_end[~present].reset_index(drop=True)
print(f"Manual review: {len(listed)} studies listed, {int(present.sum())} among the candidates and removed "
      f"(already excluded by script 1: {absent})")
print(f"Studies remaining after manual review: {len(analyzed)}")
analyzed.to_csv(INTERIM / "02_after_manual_review.csv", index=False, lineterminator="\n")
