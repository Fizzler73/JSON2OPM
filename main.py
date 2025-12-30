from pathlib import Path
import json

from json2opm.loader import load_json
from json2opm.mapper import map_pxm_json_to_opm

BASE = Path(r"E:\OneDrive\Projects\PythonProjects\JSON2OPM")

INPUT_DIR = BASE / "input_json"
REF_DIR = BASE / "reference"
OUT_DIR = BASE / "output_opm"


def main():
    OUT_DIR.mkdir(exist_ok=True)

    for src_path in INPUT_DIR.glob("*.json"):
        src_json = load_json(src_path)

        opm_json = map_pxm_json_to_opm(src_json)

        out_path = OUT_DIR / (src_path.stem + ".opm")

        with out_path.open("w", encoding="utf-8") as f:
            json.dump(opm_json, f, indent=2)

        print(f"Converted → {out_path.name}")


if __name__ == "__main__":
    main()
