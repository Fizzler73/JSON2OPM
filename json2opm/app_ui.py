import json
import csv
from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter import scrolledtext

from json2opm.loader import load_json
from json2opm.mapper import map_pxm_json_to_opm


# Stores last-used folders next to your app (project root)
SETTINGS_FILE = Path(__file__).resolve().parent.parent / "settings.json"


def _load_settings() -> dict:
    if SETTINGS_FILE.exists():
        try:
            with SETTINGS_FILE.open("r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_settings(data: dict) -> None:
    try:
        with SETTINGS_FILE.open("w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


class JSON2OPMApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("JSON2OPM Converter")
        self.geometry("900x600")

        self.settings = _load_settings()
        self.input_dir: Path | None = None
        self.output_dir: Path | None = None

        # Length delta threshold (raw units)
        self.length_delta_var = tk.StringVar()

        # Last-run mismatch rows for CSV export
        self.mismatch_rows: list[dict] = []
        self.last_compare_stats: dict = {}

        self._build_ui()
        self._restore_last_paths()
        self._restore_length_threshold()

    def _build_ui(self):
        top = tk.Frame(self)
        top.pack(fill="x", padx=10, pady=10)

        # Input folder
        tk.Button(top, text="Select Input Folder", command=self.choose_input).grid(row=0, column=0, sticky="w")
        self.input_label = tk.Label(top, text="(none)", anchor="w")
        self.input_label.grid(row=0, column=1, sticky="we", padx=(10, 0))

        # Output folder
        tk.Button(top, text="Select Output Folder", command=self.choose_output).grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.output_label = tk.Label(top, text="(none)", anchor="w")
        self.output_label.grid(row=1, column=1, sticky="we", padx=(10, 0), pady=(6, 0))

        # Length threshold field
        thresh_frame = tk.Frame(top)
        thresh_frame.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))

        tk.Label(thresh_frame, text="Length Δ threshold (raw units):").pack(side="left")
        self.length_entry = tk.Entry(thresh_frame, width=10, textvariable=self.length_delta_var)
        self.length_entry.pack(side="left", padx=(8, 0))
        tk.Label(thresh_frame, text="(example: 0.25)").pack(side="left", padx=(8, 0))

        top.columnconfigure(1, weight=1)

        # Progress bar
        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(0, 6))

        # Status label
        self.status_label = tk.Label(self, text="Ready.", anchor="w")
        self.status_label.pack(fill="x", padx=10)

        # Buttons row
        btns = tk.Frame(self)
        btns.pack(fill="x", padx=10, pady=(6, 0))

        self.convert_btn = tk.Button(btns, text="Convert JSON → OPM", command=self.convert)
        self.convert_btn.pack(side="left")

        self.export_btn = tk.Button(btns, text="Export Mismatch CSV", command=self.export_mismatches_csv, state="disabled")
        self.export_btn.pack(side="left", padx=(10, 0))

        # Log
        self.log = scrolledtext.ScrolledText(self, height=22)
        self.log.pack(fill="both", expand=True, padx=10, pady=10)
        self.log.tag_configure("error", foreground="red", font=("Consolas", 10, "bold"))

    def choose_input(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_dir = Path(folder)
            self.input_label.config(text=str(self.input_dir))

    def choose_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir = Path(folder)
            self.output_label.config(text=str(self.output_dir))

    def _persist_paths(self):
        self.settings["last_input_dir"] = str(self.input_dir) if self.input_dir else ""
        self.settings["last_output_dir"] = str(self.output_dir) if self.output_dir else ""
        _save_settings(self.settings)

    def _restore_last_paths(self):
        last_in = self.settings.get("last_input_dir")
        last_out = self.settings.get("last_output_dir")
        if last_in and Path(last_in).exists():
            self.input_dir = Path(last_in)
            self.input_label.config(text=str(self.input_dir))
        if last_out and Path(last_out).exists():
            self.output_dir = Path(last_out)
            self.output_label.config(text=str(self.output_dir))

    def _restore_length_threshold(self):
        v = self.settings.get("length_delta_threshold", 0.25)
        try:
            v = float(v)
        except Exception:
            v = 0.25
        self.length_delta_var.set(str(v))

    def _get_length_threshold(self) -> float:
        try:
            v = float(self.length_delta_var.get().strip())
            return max(0.0, v)
        except Exception:
            return 0.25

    def _persist_length_threshold(self):
        self.settings["length_delta_threshold"] = self._get_length_threshold()
        _save_settings(self.settings)

    def log_line(self, text: str = ""):
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.update_idletasks()

    def log_error_banner(self, text: str):
        self.log.insert(tk.END, "\n" + "=" * 30 + "\n", "error")
        self.log.insert(tk.END, text + "\n", "error")
        self.log.insert(tk.END, "=" * 30 + "\n", "error")
        self.log.see(tk.END)
        self.update_idletasks()

    def _set_status(self, text: str):
        self.status_label.config(text=text)
        self.update_idletasks()

    # ----------------------------
    # Export mismatches
    # ----------------------------

    def export_mismatches_csv(self):
        if not self.mismatch_rows:
            messagebox.showinfo("No mismatches", "No mismatch rows available to export.")
            return

        default_name = "az_mismatches.csv"
        out_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=default_name,
            title="Save mismatch CSV",
        )
        if not out_path:
            return

        fields = [
            "pair_key",
            "severity",
            "issue_type",
            "expected_polarity",
            "a_polarity",
            "z_polarity",
            "a_wavelengths_nm",
            "z_wavelengths_nm",
            "a_length",
            "z_length",
            "length_delta",
            "length_threshold",
        ]

        try:
            with open(out_path, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=fields)
                w.writeheader()
                for row in self.mismatch_rows:
                    # ensure all keys exist
                    w.writerow({k: row.get(k, "") for k in fields})

            self.log_line()
            self.log_line(f"Mismatch CSV exported: {out_path}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))

    # ----------------------------
    # A/Z comparison helpers
    # ----------------------------

    def _extract_az_pair_key(self, stem: str):
        """Return (pair_key, side_letter) for stems like '..._A01_...' or '..._Z01_...'."""
        m = re.match(r"^(?P<prefix>.+)_(?P<side>[AZ])(?P<snum>\d{2})_(?P<rest>.+)$", stem)
        if not m:
            return None, None
        pair_key = f"{m.group('prefix')}_{m.group('snum')}_{m.group('rest')}"
        return pair_key, m.group("side")

    def _get_actual_polarity(self, doc: dict) -> str | None:
        try:
            return (
                doc.get("Measurement", {})
                .get("OpmResultData", {})
                .get("Connectors", {})
                .get("ActualConnectors", {})
                .get("PolarityType")
            )
        except Exception:
            return None

    def _get_expected_polarity(self, doc: dict) -> str | None:
        try:
            return (
                doc.get("Measurement", {})
                .get("OpmResultData", {})
                .get("Connectors", {})
                .get("ExpectedConnectors", {})
                .get("PolarityType")
            )
        except Exception:
            return None

    def _get_wavelengths_nm(self, doc: dict) -> list[float]:
        meas = doc.get("Measurement", {}).get("OpmResultData", {}).get("Measurements", [])
        out = set()
        for m in meas if isinstance(meas, list) else []:
            w = (m or {}).get("Wavelength", {})
            nm = w.get("Nominal")
            if isinstance(nm, (int, float)):
                out.add(float(nm))
            elif isinstance(nm, str):
                try:
                    out.add(float(nm))
                except Exception:
                    pass
        return sorted(out)

    def _get_length_info(self, doc: dict) -> tuple[str | None, float | None]:
        """
        Returns (source, value)

        source:
          - "LengthInMeters" if FiberLength.LengthInMeters was used
          - "LengthInfo.Length" if FiberLength.LengthInfo.Length was used
          - None if no FiberLength found anywhere

        value:
          - float if parseable, else None

        Notes:
          - No unit conversion: OPM doesn’t encode ft vs m consistently.
          - Status is ignored: only null/missing matters (per your rule).
        """
        meas = doc.get("Measurement", {}).get("OpmResultData", {}).get("Measurements", [])
        if not isinstance(meas, list):
            return None, None

        for m in meas:
            fl = (m or {}).get("FiberLength")
            if not isinstance(fl, dict):
                continue

            source = None
            raw = None

            if "LengthInMeters" in fl:
                source = "LengthInMeters"
                raw = fl.get("LengthInMeters")
            else:
                li = fl.get("LengthInfo")
                if isinstance(li, dict) and "Length" in li:
                    source = "LengthInfo.Length"
                    raw = li.get("Length")

            if isinstance(raw, str):
                try:
                    raw = float(raw)
                except Exception:
                    raw = None

            value = float(raw) if isinstance(raw, (int, float)) else None
            return source, value

        return None, None

    def _fmt(self, v) -> str:
        if v is None:
            return "(missing)"
        if isinstance(v, float):
            return f"{v:.3f}".rstrip("0").rstrip(".")
        return str(v)

    def _log_compare_summary(self, stats: dict):
        self.log_line("A/Z Summary")
        self.log_line("-----------")
        self.log_line(f"Pairs checked: {stats.get('pairs_checked', 0)}")
        self.log_line(f"Mismatches: {stats.get('mismatch_blocks', 0)}")
        self.log_line(f"Polarity mismatches: {stats.get('polarity_mismatches', 0)}")
        self.log_line(f"Wavelength mismatches: {stats.get('wavelength_mismatches', 0)}")
        self.log_line(f"Length missing: {stats.get('length_missing', 0)}")
        self.log_line(f"Length mismatches (>threshold): {stats.get('length_mismatch', 0)}")
        self.log_line()

    def _log_az_differences(self, produced_opm_paths: list[Path]):
        """Pair A/Z outputs and log mismatches in polarity, wavelength, or length."""
        # Reset last-run mismatch export data
        self.mismatch_rows = []
        self.export_btn.config(state="disabled")

        pairs: dict[str, dict[str, Path]] = {}
        for p in produced_opm_paths:
            key, side = self._extract_az_pair_key(p.stem)
            if not key or side not in ("A", "Z"):
                continue
            pairs.setdefault(key, {})[side] = p

        length_threshold = self._get_length_threshold()

        stats = {
            "pairs_checked": 0,
            "mismatch_blocks": 0,
            "polarity_mismatches": 0,
            "wavelength_mismatches": 0,
            "length_missing": 0,
            "length_mismatch": 0,
        }

        # Precompute all mismatches, then print summary first
        blocks: list[tuple[str, str, list[str]]] = []  # (severity, key, lines)

        for key, sides in sorted(pairs.items()):
            if "A" not in sides or "Z" not in sides:
                continue

            stats["pairs_checked"] += 1

            try:
                a_doc = load_json(sides["A"])
                z_doc = load_json(sides["Z"])

                # Polarity
                expected_pol = self._get_expected_polarity(a_doc) or self._get_expected_polarity(z_doc)
                a_pol = self._get_actual_polarity(a_doc)
                z_pol = self._get_actual_polarity(z_doc)

                polarity_mismatch = (a_pol != z_pol)

                # Wavelengths
                a_wl = self._get_wavelengths_nm(a_doc)
                z_wl = self._get_wavelengths_nm(z_doc)
                wavelength_mismatch = (a_wl != z_wl)

                # Length
                a_len_src, a_len_val = self._get_length_info(a_doc)
                z_len_src, z_len_val = self._get_length_info(z_doc)
                a_has_len = isinstance(a_len_val, (int, float))
                z_has_len = isinstance(z_len_val, (int, float))

                length_missing = (not a_has_len) or (not z_has_len)

                length_mismatch = False
                length_delta = None
                if a_has_len and z_has_len:
                    length_delta = abs(float(a_len_val) - float(z_len_val))
                    length_mismatch = (length_delta > length_threshold)

                # If no issues, skip
                if not (polarity_mismatch or wavelength_mismatch or length_missing or length_mismatch):
                    continue

                # Determine severity for the block
                severity = "ERROR" if polarity_mismatch else "WARNING"

                lines: list[str] = []

                # Polarity
                if polarity_mismatch:
                    stats["polarity_mismatches"] += 1
                    lines.append("  Polarity")
                    lines.append(f"    Expected: {self._fmt(expected_pol)}")
                    lines.append(f"    A: {self._fmt(a_pol)}")
                    lines.append(f"    Z: {self._fmt(z_pol)}")
                    lines.append("")

                    self.mismatch_rows.append({
                        "pair_key": key,
                        "severity": severity,
                        "issue_type": "Polarity",
                        "expected_polarity": expected_pol,
                        "a_polarity": a_pol,
                        "z_polarity": z_pol,
                        "a_wavelengths_nm": "",
                        "z_wavelengths_nm": "",
                        "a_length": "",
                        "z_length": "",
                        "length_delta": "",
                        "length_threshold": length_threshold,
                    })

                # Wavelengths
                if wavelength_mismatch:
                    stats["wavelength_mismatches"] += 1
                    lines.append("  Wavelengths (nm)")
                    lines.append(f"    A: {a_wl}")
                    lines.append(f"    Z: {z_wl}")
                    lines.append("")

                    self.mismatch_rows.append({
                        "pair_key": key,
                        "severity": severity,
                        "issue_type": "Wavelengths",
                        "expected_polarity": "",
                        "a_polarity": "",
                        "z_polarity": "",
                        "a_wavelengths_nm": a_wl,
                        "z_wavelengths_nm": z_wl,
                        "a_length": "",
                        "z_length": "",
                        "length_delta": "",
                        "length_threshold": length_threshold,
                    })

                # Length missing
                if length_missing:
                    stats["length_missing"] += 1
                    lines.append("  Length")
                    lines.append(f"    A: {self._fmt(a_len_val)}")
                    lines.append(f"    Z: {self._fmt(z_len_val)}")
                    lines.append("")

                    self.mismatch_rows.append({
                        "pair_key": key,
                        "severity": severity,
                        "issue_type": "LengthMissing",
                        "expected_polarity": "",
                        "a_polarity": "",
                        "z_polarity": "",
                        "a_wavelengths_nm": "",
                        "z_wavelengths_nm": "",
                        "a_length": a_len_val,
                        "z_length": z_len_val,
                        "length_delta": "",
                        "length_threshold": length_threshold,
                    })

                # Length mismatch beyond threshold
                if length_mismatch:
                    stats["length_mismatch"] += 1
                    lines.append("  Length")
                    lines.append(f"    Length mismatch (raw): A={self._fmt(float(a_len_val))}  Z={self._fmt(float(z_len_val))}")
                    lines.append(f"    Delta: {self._fmt(length_delta)}   Threshold: {self._fmt(length_threshold)}")
                    lines.append("")

                    self.mismatch_rows.append({
                        "pair_key": key,
                        "severity": severity,
                        "issue_type": "LengthMismatch",
                        "expected_polarity": "",
                        "a_polarity": "",
                        "z_polarity": "",
                        "a_wavelengths_nm": "",
                        "z_wavelengths_nm": "",
                        "a_length": a_len_val,
                        "z_length": z_len_val,
                        "length_delta": length_delta,
                        "length_threshold": length_threshold,
                    })

                stats["mismatch_blocks"] += 1
                blocks.append((severity, key, lines))

            except Exception as e:
                # Treat compare errors as warnings, exportable
                stats["mismatch_blocks"] += 1
                blocks.append(("WARNING", key, [f"  Compare error: {e}", ""]))
                self.mismatch_rows.append({
                    "pair_key": key,
                    "severity": "WARNING",
                    "issue_type": "CompareError",
                    "expected_polarity": "",
                    "a_polarity": "",
                    "z_polarity": "",
                    "a_wavelengths_nm": "",
                    "z_wavelengths_nm": "",
                    "a_length": "",
                    "z_length": "",
                    "length_delta": "",
                    "length_threshold": length_threshold,
                })

        # Store stats and print summary first
        self.last_compare_stats = stats
        self._log_compare_summary(stats)

        # Then print blocks
        for severity, key, lines in blocks:
            self.log_line(f"{severity}  A/Z Mismatch: {key}")
            for ln in lines:
                self.log_line(ln.rstrip())
            self.log_line()

        # Enable export if anything exists
        if self.mismatch_rows:
            self.export_btn.config(state="normal")

        self.log_line(f"A/Z compare complete. Pairs checked: {stats['pairs_checked']}. Mismatches: {stats['mismatch_blocks']}.")

    # ----------------------------
    # Main convert routine
    # ----------------------------

    def _explain_duplicate_output(self, out_path: Path, src_path: Path) -> str:
        return (
            f"Output file already exists and overwrite is disabled:\n"
            f"  Output: {out_path}\n"
            f"  Input : {src_path}\n\n"
            f"Fix: choose an empty output folder, delete/rename the existing .opm file, or rename the input JSON(s) "
            f"so they produce unique output names."
        )

    def convert(self):
        if not self.input_dir or not self.output_dir:
            messagebox.showerror("Missing folder", "Please select both input and output folders.")
            return

        json_files = list(self.input_dir.glob("*.json"))
        if not json_files:
            messagebox.showwarning("No files", "No JSON files found in input folder.")
            return

        self._persist_paths()
        self._persist_length_threshold()

        # Reset mismatch export data per run
        self.mismatch_rows = []
        self.export_btn.config(state="disabled")

        self.convert_btn.config(state="disabled")
        try:
            total = len(json_files)
            self.progress["value"] = 0
            self.progress["maximum"] = total

            self.log_line(f"Starting conversion of {total} files...\n")
            self._set_status(f"Converting 0 / {total}")

            success = 0
            failed = 0
            produced_opm_paths: list[Path] = []

            for i, src_path in enumerate(json_files, start=1):
                try:
                    src_json = load_json(src_path)
                    opm_json = map_pxm_json_to_opm(src_json)

                    out_path = self.output_dir / (src_path.stem + ".opm")

                    # Option 1: fail on duplicates (no overwrite)
                    if out_path.exists():
                        raise FileExistsError(self._explain_duplicate_output(out_path, src_path))

                    with out_path.open("w", encoding="utf-8") as f:
                        json.dump(opm_json, f, indent=2)

                    produced_opm_paths.append(out_path)
                    self.log_line(f"✔ Converted: {src_path.name}")
                    success += 1
                except Exception as e:
                    self.log_line(f"✖ FAILED: {src_path.name} → {e}")
                    failed += 1

                self.progress["value"] = i
                self._set_status(f"Converting {i} / {total}")

            # Post-pass: compare A/Z pairs and flag mismatches
            self.log_line()
            self._log_az_differences(produced_opm_paths)

            self.log_line("\nDone.")
            self.log_line(f"Success: {success}")
            self.log_line(f"Failed: {failed}")
            self._set_status(f"Done. Success: {success}  Failed: {failed}")

            if failed > 0:
                self.log_error_banner(
                    f"❌❌❌   FAILURES DETECTED   ❌❌❌\n"
                    f"Total failed files: {failed}\n"
                    f"See log above for detailed reasons."
                )

        finally:
            self.convert_btn.config(state="normal")
