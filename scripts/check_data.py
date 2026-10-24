from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SITES = [
    {
        "name": "High Tunnel", "folder": "02_high_tunnel", "key": "Genotypes", "unit": "g",
        "master": "harvest_master.csv", "summary": "genotype_summary.csv", "date": "Date of harvest",
        "weights": [("Total Harvest Weight", "total_weight"), ("Total Marketable Weight", "total_marketable_weight")],
        "coverage": "n_harvests_with_weight",
    },
    {
        "name": "Robinson", "folder": "04_robinson", "key": "Genotypes", "unit": "g",
        "master": "robinson_harvest_master.csv", "summary": "robinson_genotype_summary.csv", "date": "Date of harvest",
        "weights": [("Total Harvest Weight", "total_weight"), ("Total Marketable Weight", "total_marketable_weight")],
        "coverage": None,
    },
    {
        "name": "South Farm", "folder": "05_south_farm", "key": "Genotypes", "unit": "g",
        "master": "south_farm_harvest_master.csv", "summary": "south_farm_genotype_summary.csv", "date": "Date of harvest",
        "weights": [("Total Harvest Weight (g)", "total_weight_g"), ("Marketable Weight (g)", "total_marketable_weight_g")],
        "coverage": "n_harvests_with_weight",
    },
    {
        "name": "Princeton", "folder": "06_princeton", "key": "Entry", "unit": "lb",
        "master": "princeton_harvest_master.csv", "summary": "princeton_genotype_summary.csv", "date": "Date",
        "weights": [("Marketable lb", "Marketable lb"), ("Cull lb", "Cull lb")],
        "coverage": "Harvests with marketable weight",
    },
]


def compare_totals(master, summary, key, source_column, summary_column):
    expected = master.groupby(key, dropna=False)[source_column].sum(min_count=1)
    actual = summary.set_index(key)[summary_column]
    comparison = pd.concat([expected.rename("from_master"), actual.rename("from_summary")], axis=1)
    comparison["difference"] = comparison["from_summary"] - comparison["from_master"]
    comparison["matches"] = np.isclose(
        comparison["from_master"], comparison["from_summary"], rtol=0, atol=0.000001, equal_nan=True
    )
    comparison.index.name = "entry"
    return comparison.reset_index()


def main():
    review = ROOT / "review"
    review.mkdir(exist_ok=True)
    checks, overview, missing, differences = [], [], [], []

    def record(site, check, count, note=""):
        checks.append({"site": site, "check": check, "items_to_review": int(count), "note": note})

    for site in SITES:
        folder = ROOT / "batches" / site["folder"]
        master = pd.read_csv(folder / site["master"])
        summary = pd.read_csv(folder / site["summary"])
        key = site["key"]
        weight = site["weights"][0][0]
        dates = pd.to_datetime(master[site["date"]], errors="coerce")
        overview.append({
            "site": site["name"], "rows": len(master), "entries": master[key].nunique(),
            "harvest_dates": dates.nunique(), "weight_unit": site["unit"],
            "weight_measure": weight, "rows_without_weight": int(master[weight].isna().sum()),
        })
        record(site["name"], "Missing entry ID", master[key].isna().sum())
        record(site["name"], "Missing or invalid harvest date", dates.isna().sum())
        record(site["name"], "Exact duplicate master rows", master.duplicated(keep=False).sum(), "Repeated entries alone are not assumed to be duplicates.")
        duplicate_keys = summary[key].duplicated(keep=False).sum()
        record(site["name"], "Repeated summary entry IDs", duplicate_keys)
        master_keys, summary_keys = set(master[key].dropna()), set(summary[key].dropna())
        record(site["name"], "Entry IDs missing from either table", len(master_keys ^ summary_keys))
        for column in master:
            missing.append({
                "site": site["name"], "column": column, "rows": len(master),
                "missing": int(master[column].isna().sum()),
            })
        for source_column, summary_column in site["weights"]:
            record(site["name"], "Negative " + source_column, master[source_column].lt(0).sum())
            if duplicate_keys:
                record(site["name"], "Cannot reconcile " + summary_column, duplicate_keys, "Summary entry IDs must be unique first.")
                continue
            comparison = compare_totals(master, summary, key, source_column, summary_column)
            failed = comparison.loc[~comparison["matches"]].copy()
            failed.insert(0, "site", site["name"])
            failed.insert(1, "measure", summary_column)
            failed["unit"] = site["unit"]
            differences.extend(failed.to_dict("records"))
            record(site["name"], "Master vs summary: " + summary_column, len(failed), "Missing measurements are kept distinct from zero.")
        if site["coverage"] and not duplicate_keys:
            expected = master.loc[master[weight].notna()].groupby(key)["Harvest"].nunique()
            actual = summary.set_index(key)[site["coverage"]]
            coverage = pd.concat([expected.rename("expected"), actual.rename("actual")], axis=1)
            coverage["expected"] = coverage["expected"].fillna(0)
            record(site["name"], "Measured harvest coverage", coverage["expected"].ne(coverage["actual"]).sum())

    pd.DataFrame(overview).to_csv(review / "site_overview.csv", index=False)
    pd.DataFrame(missing).to_csv(review / "missing_values.csv", index=False)
    pd.DataFrame(checks).to_csv(review / "checks.csv", index=False)
    pd.DataFrame(differences, columns=["site", "measure", "entry", "from_master", "from_summary", "difference", "matches", "unit"]).to_csv(review / "reconciliation_differences.csv", index=False)
    brix = pd.read_csv(ROOT / "batches/03_brix/sg_brix_master.csv")
    brix.loc[brix["QC_flag"].notna()].to_csv(review / "brix_flags.csv", index=False)
    flagged = [row for row in checks if row["items_to_review"]]
    lines = ["# Data checks", "", f"Checks run: {len(checks)}. Checks needing review: {len(flagged)}.", ""]
    if flagged:
        lines += ["| Site | Check | Items |", "| --- | --- | --- |"]
        lines += [f"| {row['site']} | {row['check']} | {row['items_to_review']} |" for row in flagged]
    else:
        lines.append("No differences found in the CSV checks.")
    lines += ["", f"Flagged Brix readings: {int(brix['QC_flag'].notna().sum())}. See `brix_flags.csv`.", "", "- `missing_values.csv`: missing cells by column.", "- `reconciliation_differences.csv`: master-to-summary differences.", "- `site_overview.csv`: entry counts, harvest coverage, and units."]
    (review / "README.md").write_text("\n".join(lines) + "\n")
    print(f"Checked {len(checks)} conditions across {len(SITES)} sites.")
    print(f"Conditions needing review: {len(flagged)}")
    return bool(flagged)


if __name__ == "__main__":
    raise SystemExit(main())
