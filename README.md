# 2026 tomato harvest analysis

Yield and Brix analysis for the Sapkota Lab's 2026 tomato trials.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
jupyter lab
```

## Analyses

| Batch | File |
| --- | --- |
| 01 | [Workbook cleaning](batches/01_workbook_cleaning/clean_harvest_workbook.py) |
| 02 | [High Tunnel](batches/02_high_tunnel/high_tunnel.ipynb) |
| 03 | [Brix](batches/03_brix/brix.ipynb) |
| 04 | [Robinson](batches/04_robinson/robinson.ipynb) |
| 05 | [South Farm](batches/05_south_farm/south_farm.ipynb) |
| 06 | [Princeton](batches/06_princeton/princeton.ipynb) |

Run in batch order. If the cleaned workbook is already available, start at 02. Brix reads the High Tunnel summary; Robinson, South Farm, and Princeton use earlier batch outputs.

Input files are listed in [data/README.md](data/README.md). CSVs and figures are saved alongside each notebook.

## Review with Dr. Sapkota

| Batch | Email | Response |
| --- | --- | --- |
| 01 — Workbook cleaning | Sent | Acknowledged |
| 02 — High Tunnel | Sent | N/A |
| 03 — Brix | Sent | Acknowledged |
| 04 — Robinson | Not sent | — |
| 05 — South Farm | Not sent | — |
| 06 — Princeton | Sent | Acknowledged |

## Checks

```bash
python scripts/run_notebooks.py
python scripts/check_data.py
python -m unittest discover -s tests
```

For one batch: `python scripts/run_notebooks.py 06_princeton`.

[review/](review/) holds missing-value counts, total reconciliations, and flagged readings. The check script exits with code 1 if a structural or reconciliation check fails.

[Analysis notes](methods.md) · [Output fields](data_dictionary.md) · [Project tracker](https://github.com/Sapkota-Lab/Project-Tracker/issues/12)
