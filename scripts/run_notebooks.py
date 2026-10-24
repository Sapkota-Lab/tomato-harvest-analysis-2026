import argparse
from pathlib import Path

import nbformat
from nbclient import NotebookClient

ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("batch", nargs="?", help="Batch folder, such as 02_high_tunnel. Omit to run all five notebooks.")
    args = parser.parse_args()
    notebooks = sorted((ROOT / "batches").glob("*/*.ipynb"))
    if args.batch:
        notebooks = [path for path in notebooks if path.parent.name == args.batch]
        if not notebooks:
            parser.error("No notebook found for that batch.")
    for path in notebooks:
        print("Running", path.relative_to(ROOT), flush=True)
        notebook = nbformat.read(path, as_version=4)
        client = NotebookClient(
            notebook, timeout=180, kernel_name="python3",
            resources={"metadata": {"path": str(ROOT)}},
        )
        client.execute()
        nbformat.write(notebook, path)
        print("Saved", path.name, flush=True)


if __name__ == "__main__":
    main()
