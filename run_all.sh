#!/usr/bin/env bash
# Runs stages 1, 2, 6 and 7 end to end, from the repository root.
#
#   python -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt
#   ./run_all.sh
#
# Stage 6 needs the completed working master (data/manual/working_master.xlsx), which is private and
# not in the repository.  Without it the committed data/Crossover_Trials_Master.xlsx is used as is.
set -euo pipefail
cd "$(dirname "$0")"
PYTHON="${PYTHON:-python}"

echo "== Stage 1: filter the HemOncKB export"
"$PYTHON" scripts/01_filter_hemonckb_studies.py
echo
echo "== Stage 2: match FDA approvals; apply the manual review lists"
"$PYTHON" scripts/02_match_fda_approvals.py
echo
if [ -f data/manual/working_master.xlsx ]; then
    echo "== Stage 6: build the analysis dataset from the working master"
    "$PYTHON" scripts/03_make_analysis_dataset.py
else
    echo "== Stage 6 skipped: data/manual/working_master.xlsx is not present; using the committed analysis dataset"
fi
echo
echo "== Stage 7: run the analysis notebook (tables and Figure 2)"
"$PYTHON" -m jupyter nbconvert --to notebook --execute --inplace \
    --ExecutePreprocessor.timeout=900 analysis/crossover_os_analysis.ipynb
# Per-cell run timestamps and the local kernel name would make every run differ; drop them.
"$PYTHON" - <<'PY'
import json
from pathlib import Path
p = Path("analysis/crossover_os_analysis.ipynb")
nb = json.loads(p.read_text(encoding="utf-8"))
for cell in nb["cells"]:
    cell["metadata"].pop("execution", None)
nb["metadata"]["kernelspec"] = {"display_name": "Python 3 (ipykernel)", "language": "python", "name": "python3"}
p.write_text(json.dumps(nb, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
PY
echo
echo "Done. Outputs are in outputs/."
