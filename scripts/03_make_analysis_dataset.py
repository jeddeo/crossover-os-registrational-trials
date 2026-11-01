"""Stage 6: turn the completed working master into the clean, values-only analysis dataset.

The working master is private: its VLOOKUP columns pull from reviewer check tabs, and it
carries reviewer names and cell comments.  This script copies the *cached values* of its
`Final Study List` tab (it never edits the master, and never deletes tabs from a workbook
that still contains formulas) and writes the public record:

    data/Crossover_Trials_Master.xlsx   one sheet, values only
    data/Crossover_Trials_Master.csv    the same content
    data/Data_Dictionary.csv            one row per column
    data/manual/abstraction_template.xlsx   the headers, with no data rows

Run   : python scripts/03_make_analysis_dataset.py [path/to/working_master.xlsx]
Default input: data/manual/working_master.xlsx (gitignored; never committed).
"""
import argparse
import csv
import datetime as dt
import json
import re
import shutil
import sys
import zipfile
from collections import Counter
from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
SHEET = "Final Study List"
N_ROWS = 86            # one row per comparison: 78 trials, 8 of them contribute two rows
N_COMMENTS = 20
N_PROXY_LINKS = 8
CREATOR = "Crossover study team"
FIXED_DATE = (2026, 1, 1, 0, 0, 0)

ERROR_TOKENS = {"#VALUE!", "#DIV/0!", "#N/A", "#REF!"}
# The only columns allowed to hold a formula error: the Western-enrollment sum and the
# crossover numerators, denominators and fractions that depend on a "NR" input.
ERROR_COLUMNS = {
    "perc_pts_Americas+Europe",
    "Total_Control_Pts_Receiving_Subsequent_Therapy",
    "Total_Control_Pts_With_EFS_or_PFS_Event",
    "Total_Control_Pts_Progressed",
    "Num control pts getting post-protocol drug in class",
    "Num control pts getting exact drug post-protocol",
    "Frac_ctrl_pts_getting_post-protocol_drug_in_class_progressed_subsequent_therapy",
    "Frac_ctrl_pts_getting_exact_drug_post-protocol_progressed_subsequent_therapy",
    "Frac_ctrl_pts_getting_post-protocol_drug_in_class_event",
    "Frac_ctrl_pts_getting_exact_drug_post-protocol_event",
    "Frac_ctrl_pts_getting_post-protocol_drug_in_class_progressed",
    "Frac_ctrl_pts_getting_exact_drug_post-protocol_progressed",
}
SINGLE_ERROR_COLUMNS = {
    "Total_Control_Pts_Receiving_Subsequent_Therapy",
    "Total_Control_Pts_With_EFS_or_PFS_Event",
    "Total_Control_Pts_Progressed",
    "Num control pts getting post-protocol drug in class",
    "Num control pts getting exact drug post-protocol",
}
RENAMES = {"perc_pts_Americas+Europe": "perc_pts_NorthAmerica_Europe"}
PROXY_SUFFIX = ".laneproxy.stanford.edu"


def public_url(url):
    """Undo an EZproxy rewrite: `www-nejm-org.laneproxy.stanford.edu/...` -> `www.nejm.org/...`.
    The proxy turns dots into hyphens (and a real hyphen into a double hyphen)."""
    m = re.match(r"^(https?://)([^/]+)(.*)$", url, re.S)
    scheme, host, rest = m.groups()
    if not host.endswith(PROXY_SUFFIX):
        return url
    host = host[: -len(PROXY_SUFFIX)].replace("--", "\0").replace("-", ".").replace("\0", "-")
    return f"{scheme}{host}{rest}"


def freeze_xlsx(path, when=FIXED_DATE):
    """Rewrite an .xlsx with fixed zip-member timestamps and a fixed core.xml date, so that
    rerunning the script gives byte-identical files."""
    stamp = "%04d-%02d-%02dT%02d:%02d:%02dZ" % when
    tmp = f"{path}.tmp"
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in sorted(src.infolist(), key=lambda i: i.filename):
            payload = src.read(item.filename)
            if item.filename == "docProps/core.xml":
                payload = re.sub(rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms:)",
                                 rb"\g<1>" + stamp.encode() + rb"\g<2>", payload)
            info = zipfile.ZipInfo(item.filename, date_time=when)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = item.external_attr
            dst.writestr(info, payload)
    shutil.move(tmp, path)


def csv_value(v):
    """Deterministic text for a cell: ISO dates, shortest round-trip floats, empty for None."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, dt.datetime):
        return v.date().isoformat() if v.time() == dt.time(0) else v.isoformat(sep=" ")
    if isinstance(v, float):
        return repr(int(v)) if v.is_integer() and abs(v) < 1e15 else repr(v)
    return str(v)


def write_workbook(path, header, rows, formats):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append(header)
    for row in rows:
        ws.append(row)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    for j, fmt in enumerate(formats, start=1):
        if fmt != "General":
            for i in range(2, len(rows) + 2):
                ws.cell(row=i, column=j).number_format = fmt
    wb.properties.creator = CREATOR
    wb.properties.lastModifiedBy = CREATOR
    wb.properties.created = wb.properties.modified = dt.datetime(*FIXED_DATE)
    wb.save(path)
    freeze_xlsx(path)


# ── Data dictionary text ─────────────────────────────────────────────────────────
# Descriptions for columns copied from the HemOnc.org Knowledge Base follow its published
# data dictionary (data.dictionary.xlsx in the HemOncKB Dataverse release).  Proportions are
# fractions 0-1, and "NR" means not reported.
CROSS_DEF = {
    "exact_drug": "control patients who subsequently received the exact experimental drug",
    "in_class": "control patients who subsequently received any drug in the experimental drug's class",
}
DENOM_DEF = {
    "subsequent_therapy": "control patients who received any subsequent systemic therapy (primary definition)",
    "progressed": "control patients who progressed",
    "event": "control patients with a PFS, DFS or EFS event",
}
DESCRIPTIONS = {
    "study": ("Trial name, as in HemOncKB (the export's \"PSMAfore\" is written \"PSMA Fore\")", "text", ""),
    "Link 1": ("URL of the first source publication used for abstraction", "text", "public URL"),
    "Link 2": ("URL of a second source publication used for abstraction, when one was needed", "text", "public URL"),
    "Location of post-protocol treatment information": (
        "Where in the sources the post-protocol (subsequent) treatment counts were found", "text", ""),
    "Specific Treatment": (
        "Experimental agent carrying the qualifying later-line FDA approval; drug class is assigned from it",
        "text", ""),
    "perc_pts_WesternEurope": ("Proportion of enrolled participants from Western Europe", "proportion", "fraction 0-1; NR = not reported"),
    "perc_pts_US": ("Proportion of enrolled participants from the United States", "proportion", "fraction 0-1; NR = not reported"),
    "perc_pts_Europe": ("Proportion of enrolled participants from Europe (Western Europe if Europe was not reported)", "proportion", "fraction 0-1; NR = not reported"),
    "perc_pts_NorthAmerica": ("Proportion of enrolled participants from North America (US, with or without Canada, if North America was not reported)", "proportion", "fraction 0-1; NR = not reported"),
    "perc_pts_NorthAmerica_Europe": ("Western enrollment: proportion of enrolled participants from North America plus Europe", "proportion", "fraction 0-1; NR = not reported or not computable"),
    "ITT_Control_Num": ("Number of control-arm participants in the intention-to-treat population", "integer", "participants"),
    "Total_Control_Pts_Receiving_Subsequent_Therapy": ("Denominator: control patients who received any subsequent systemic therapy", "integer", "patients; NR = not reported"),
    "Total_Control_Pts_With_EFS_or_PFS_Event": ("Denominator: control patients with a PFS, DFS or EFS event", "integer", "patients; NR = not reported"),
    "Total_Control_Pts_Progressed": ("Denominator: control patients who progressed", "integer", "patients; NR = not reported"),
    "Num control pts getting post-protocol drug in class": ("Numerator, in-class definition: control patients who subsequently received any drug in the experimental drug's class", "integer", "patients; NR = not reported"),
    "Num control pts getting exact drug post-protocol": ("Numerator, exact-drug definition: control patients who subsequently received the exact experimental drug", "integer", "patients; NR = not reported"),
    "Combined_OS_HR": ("Overall survival hazard ratio, experimental vs control", "number", "hazard ratio"),
    "Combined_OS_HR_CI_low": ("Lower limit of the 95% confidence interval of the overall survival hazard ratio", "number", "hazard ratio; NR = not reported"),
    "Combined_OS_HR_CI_high": ("Upper limit of the 95% confidence interval of the overall survival hazard ratio", "number", "hazard ratio; NR = not reported"),
    "Combined_PFS_HR": ("Hazard ratio for the progression-type endpoint (PFS, DFS or EFS), experimental vs control", "number", "hazard ratio; NR = not reported"),
    "Combined_PFS_HR_CI_low": ("Lower limit of the 95% confidence interval of the progression-type hazard ratio", "number", "hazard ratio; NR = not reported"),
    "Combined_PFS_HR_CI_high": ("Upper limit of the 95% confidence interval of the progression-type hazard ratio", "number", "hazard ratio; NR = not reported"),
    "regimen": ("Experimental regimen (HemOncKB)", "text", ""),
    "comparator": ("Comparator regimen (HemOncKB)", "text", ""),
    "efficacy": ("Efficacy summary string taken directly from HemOnc.org (HemOncKB)", "text", "HTML line breaks as <br/>"),
    "trial_id": ("Clinical trial registry identifier (HemOncKB)", "text", "NCT number"),
    "enrollment": ("Years of enrollment as reported (HemOncKB)", "text", "YYYY-MM to YYYY-MM"),
    "phase": ("Phase of the trial (HemOncKB)", "text", ""),
    "study_design": ("Study design, derived from arm-level information (HemOncKB)", "text", ""),
    "sponsor": ("Study sponsor(s) (HemOncKB)", "text", ""),
    "context": ("Treatment context in which the study-regimen pair takes place (HemOncKB)", "text", "for example Adjuvant therapy"),
    "r_modifier": ("Additional specification of the experimental regimen, such as cycles or dose (HemOncKB)", "text", "none = no modifier"),
    "condition_x": ("Condition studied, from the results table of the HemOncKB export (suffix _x: study_results)", "text", ""),
    "condition_cui_x": ("Concept identifier of the condition, results table (HemOncKB)", "number", "OMOP concept id"),
    "biomarker_x": ("Biomarker-defined subtype of the condition, results table (HemOncKB)", "text", ""),
    "c_modifier": ("Additional specification of the comparator, such as cycles or dose (HemOncKB)", "text", "none = no modifier"),
    "comparator_code": ("Code recording why a comparator is incompletely instantiated as a regimen (HemOncKB)", "text", "none = complete"),
    "error": ("Whether HemOncKB flagged a parsing error for the row (HemOncKB)", "boolean", "0/1"),
    "date_added_x": ("Date the results row was added to HemOncKB (suffix _x: study_results)", "date", "YYYY-MM-DD"),
    "registry": ("Clinical trial registry (HemOncKB)", "text", ""),
    "condition_y": ("Condition studied, from the study table of the HemOncKB export (suffix _y: study.table)", "text", ""),
    "condition_cui_y": ("Concept identifier of the condition, study table (HemOncKB)", "number", "OMOP concept id"),
    "biomarker_y": ("Biomarker-defined subtype of the condition, study table (HemOncKB)", "text", ""),
    "study_design_imputed": ("Whether the study design was imputed from partial information (HemOncKB)", "boolean", "0/1"),
    "sact": ("Whether any arm uses systemic anticancer therapy (HemOncKB)", "boolean", "0/1"),
    "protocol": ("Whether the study involves multi-phasic protocols in any arm (HemOncKB)", "boolean", "0/1"),
    "reg_study": ("Whether the study was cited by a regulatory authority in a labeling decision (HemOncKB)", "boolean", "0/1"),
    "unreg_study": ("Whether the study was cited by a regulatory authority in a withdrawal decision (HemOncKB)", "boolean", "0/1"),
    "start": ("Start of enrollment (HemOncKB; a year alone is read as 1 January)", "date", "YYYY-MM-DD"),
    "end": ("End of enrollment (HemOncKB; a year alone is read as 1 January)", "date", "YYYY-MM-DD"),
    "study_group": ("Study group(s) named in the publication (HemOncKB)", "text", ""),
    "sponsor_type": ("Type of sponsor (HemOncKB)", "text", ""),
    "pubs_in_hemonc": ("Number of publications on this study, condition and biomarker in HemOncKB", "integer", "publications"),
    "date_added_y": ("Date the study was added to HemOncKB (suffix _y: study.table)", "date", "YYYY-MM-DD"),
    "start_converted": ("Start of enrollment parsed to a date by script 1 (a year alone becomes 1 January)", "date", "YYYY-MM-DD"),
    "end_converted": ("End of enrollment parsed to a date by script 1 (a year alone becomes 1 January)", "date", "YYYY-MM-DD"),
}
# Columns the code and Methods do not define well enough to describe; the dictionary says TODO.
# They are used only by exploratory analyses that are not part of this release.
ENDPOINT_NAMES = {"DFS": "disease-free survival", "EFS": "event-free survival", "IDFS": "invasive disease-free survival",
                  "MFS": "metastasis-free survival", "PCR_RATE": "pathologic complete response rate",
                  "RFS": "relapse-free survival", "RPFS": "radiographic PFS", "TTP": "time to progression",
                  "OS": "overall survival", "PFS": "progression-free survival"}


def describe(col, dtype):
    """(description, type, units_or_coding) for a column; description is "TODO" when unknown."""
    if col in DESCRIPTIONS:
        return DESCRIPTIONS[col]
    m = re.match(r"^Frac_ctrl_pts_getting_(post-protocol_drug_in_class|exact_drug_post-protocol)_(progressed_subsequent_therapy|event|progressed)$", col)
    if m:
        num = "in_class" if "in_class" in m.group(1) else "exact_drug"
        den = {"progressed_subsequent_therapy": "subsequent_therapy", "event": "event", "progressed": "progressed"}[m.group(2)]
        return (f"Crossover fraction: {CROSS_DEF[num]}, divided by {DENOM_DEF[den]}",
                "proportion", "fraction 0-1 (can exceed 1 when the numerator exceeds the denominator); NR = not reported")
    m = re.match(r"^FDA_(Notice|Drug|Indication|Date)_(\d+)$", col)
    if m:
        what = {"Notice": "title of the FDA approval notification", "Drug": "drug named in the notification",
                "Indication": "indication named in the notification", "Date": "date of the notification"}[m.group(1)]
        return (f"Matched FDA approval notification {m.group(2)} (found by script 2): {what}", "date" if m.group(1) == "Date" else "text",
                "YYYY-MM-DD" if m.group(1) == "Date" else "")
    m = re.match(r"^Median_(\w+?)_(intervention|ctrl)$", col)
    if m and m.group(1) in ENDPOINT_NAMES:
        arm = "experimental" if m.group(2) == "intervention" else "control"
        return (f"Median {ENDPOINT_NAMES[m.group(1)]} in the {arm} arm, from the HemOncKB efficacy string",
                "number", "months, or fraction 0-1 for rate endpoints")
    m = re.match(r"^(\w+?)_HR(_CI_low|_CI_high)?$", col)
    if m and m.group(1) in ENDPOINT_NAMES:
        part = {None: "Hazard ratio", "_CI_low": "Lower 95% confidence limit of the hazard ratio",
                "_CI_high": "Upper 95% confidence limit of the hazard ratio"}[m.group(2)]
        return (f"{part} for {ENDPOINT_NAMES[m.group(1)]}, experimental vs control, from the HemOncKB efficacy string",
                "number", "hazard ratio")
    return ("TODO", dtype, "")


def infer_type(values):
    vals = [v for v in values if v not in (None, "")]
    kinds = {("date" if isinstance(v, dt.datetime) else "boolean" if isinstance(v, bool)
              else "number" if isinstance(v, (int, float)) else "text") for v in vals}
    if kinds == {"number"}:
        return "integer" if all(float(v).is_integer() for v in vals) else "number"
    if kinds <= {"number", "text"} and "number" in kinds:
        return "number"          # numeric with "NR" entries
    return kinds.pop() if len(kinds) == 1 else "text"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("master", nargs="?", default=str(ROOT / "data" / "manual" / "working_master.xlsx"),
                    help="path to the completed working master (default: data/manual/working_master.xlsx)")
    ap.add_argument("--notebook", default=str(ROOT / "analysis" / "crossover_os_analysis.ipynb"),
                    help="notebook searched for the columns it uses (Data_Dictionary.used_in_analysis)")
    args = ap.parse_args()
    out_dir = ROOT / "data"

    path = Path(args.master)
    if not path.exists():
        sys.exit(f"Working master not found: {path}\n"
                 "Copy the completed working master there (it is private and gitignored), or pass its path.")
    wb_v = openpyxl.load_workbook(path, data_only=True)
    wb_f = openpyxl.load_workbook(path)
    ws_v, ws_f = wb_v[SHEET], wb_f[SHEET]

    header = [c.value for c in ws_v[1]]
    data_rows = [r for r in range(2, ws_v.max_row + 1) if ws_v.cell(r, 1).value not in (None, "")]
    assert data_rows == list(range(2, 2 + len(data_rows))), "data rows are not contiguous"
    assert len(data_rows) == N_ROWS, f"expected {N_ROWS} data rows, found {len(data_rows)}"
    n_cols = len(header)
    print(f"Read {len(data_rows)} data rows x {n_cols} columns from {path.name}")

    # A few lookups legitimately return an empty string, which Excel caches as no value.  A
    # master that was never recalculated would have no cached value for nearly every formula.
    formulas = [(r, c) for r in data_rows for c in range(1, n_cols + 1)
                if isinstance(ws_f.cell(r, c).value, str) and ws_f.cell(r, c).value.startswith("=")]
    blank = [f"{get_column_letter(c)}{r}" for r, c in formulas if ws_v.cell(r, c).value is None]
    assert len(blank) <= 0.02 * len(formulas), \
        f"{len(blank)} of {len(formulas)} formulas have no cached value; open and save the master in Excel first"
    print(f"Formula cells: {len(formulas)}; {len(blank)} evaluate to an empty string and are written blank")

    # Values, with Excel errors turned into NR.
    rows, error_counts = [], Counter()
    for r in data_rows:
        row = []
        for c in range(1, n_cols + 1):
            v = ws_v.cell(r, c).value
            if isinstance(v, str) and v.strip() in ERROR_TOKENS:
                col = header[c - 1]
                assert col in ERROR_COLUMNS, f"Excel error {v} in unexpected column {col!r} (row {r})"
                error_counts[col] += 1
                v = "NR"
            row.append(v)
        rows.append(row)
    over = [c for c in SINGLE_ERROR_COLUMNS if error_counts[c] > 1]
    assert not over, f"expected at most one error cell in {over}"
    print(f"Excel errors -> NR: {sum(error_counts.values())} cells")
    for col, n in sorted(error_counts.items(), key=lambda t: header.index(t[0])):
        print(f"    {n:3d}  {col}")

    n_comments = sum(1 for row in ws_f.iter_rows() for c in row if c.comment)
    print(f"Cell comments dropped: {n_comments}")
    assert n_comments == N_COMMENTS, f"expected {N_COMMENTS} comments, found {n_comments}"

    # Public links.
    pairs = []
    for name in ("Link 1", "Link 2"):
        j = header.index(name)
        for row in rows:
            if isinstance(row[j], str) and PROXY_SUFFIX in row[j]:
                new = public_url(row[j])
                pairs.append((row[j], new))
                row[j] = new
    print(f"Proxy links rewritten: {len(pairs)}")
    for old, new in pairs:
        print(f"    {old}\n -> {new}")
    assert len(pairs) == N_PROXY_LINKS, f"expected {N_PROXY_LINKS} proxy links, found {len(pairs)}"

    # Empty columns.
    keep = [j for j in range(n_cols) if any(row[j] not in (None, "") for row in rows)]
    dropped = [(get_column_letter(j + 1), header[j]) for j in range(n_cols) if j not in keep]
    print(f"Empty columns dropped: {len(dropped)} ({sum(1 for _, h in dropped if h is None)} unnamed)")
    print("    " + ", ".join(f"{a}:{b or '(unnamed)'}" for a, b in dropped))
    new_header = [RENAMES.get(header[j], header[j]) for j in keep]
    assert len(set(new_header)) == len(new_header), "duplicate column names"
    new_rows = [[row[j] for j in keep] for row in rows]
    formats = [ws_v.cell(2, j + 1).number_format for j in keep]
    print(f"Renamed: {RENAMES}")

    # The trial list must be the one the pipeline produced.
    interim = out_dir / "interim" / "02_after_manual_review.csv"
    if interim.exists():
        with open(interim, newline="", encoding="utf-8") as f:
            pipeline = {r["study"].strip() for r in csv.DictReader(f)}
        trials = {str(r[0]).strip() for r in new_rows}
        assert trials == pipeline, (f"trial list differs from data/interim/02_after_manual_review.csv: "
                                    f"only in master {sorted(trials - pipeline)}, only in pipeline {sorted(pipeline - trials)}")
        print(f"Trial list matches the pipeline output ({len(trials)} trials)")

    # Outputs.
    write_workbook(out_dir / "Crossover_Trials_Master.xlsx", new_header, new_rows, formats)
    with open(out_dir / "Crossover_Trials_Master.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(new_header)
        w.writerows([[csv_value(v) for v in row] for row in new_rows])
    (out_dir / "manual").mkdir(exist_ok=True)
    write_workbook(out_dir / "manual" / "abstraction_template.xlsx", new_header, [], formats)

    nb_path = Path(args.notebook)
    if nb_path.exists():
        nb = json.loads(nb_path.read_text(encoding="utf-8"))
        source = "\n".join("".join(c["source"]) for c in nb["cells"] if c["cell_type"] == "code")
    else:
        print(f"Notebook not found ({nb_path}); used_in_analysis is set to No for every column")
        source = ""
    todo = []
    with open(out_dir / "Data_Dictionary.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["column", "description", "type", "units_or_coding", "used_in_analysis"])
        for j, col in enumerate(new_header):
            dtype = infer_type([row[j] for row in new_rows])
            desc, typ, units = describe(col, dtype)
            if desc == "TODO":
                todo.append(col)
            used = bool(re.search(r"""["']""" + re.escape(col) + r"""["']""", source))
            w.writerow([col, desc, typ, units, "Yes" if used else "No"])
    print(f"Wrote {len(new_header)} columns x {len(new_rows)} rows; Data_Dictionary.csv has {len(todo)} TODO rows")
    for col in todo:
        print(f"    TODO: {col}")


if __name__ == "__main__":
    main()
