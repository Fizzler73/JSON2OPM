import json
import csv
import copy
from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter import scrolledtext

from json2opm.loader import load_json
from json2opm.mapper import map_pxm_json_to_opm


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
        self.geometry("980x680")

        self.settings = _load_settings()
        self.input_dir: Path | None = None
        self.output_dir: Path | None = None

        self.length_delta_var = tk.StringVar()
        self.merge_var = tk.BooleanVar(value=False)

        self.mismatch_rows: list[dict] = []

        self._build_ui()
        self._restore_last_paths()
        self._restore_length_threshold()
        self._restore_merge_toggle()

    # ----------------------------
    # UI
    # ----------------------------

    def _build_ui(self):
        top = tk.Frame(self)
        top.pack(fill="x", padx=10, pady=10)

        tk.Button(top, text="Select Input Folder", command=self.choose_input).grid(row=0, column=0, sticky="w")
        self.input_label = tk.Label(top, text="(none)", anchor="w")
        self.input_label.grid(row=0, column=1, sticky="we", padx=(10, 0))

        tk.Button(top, text="Select Output Folder", command=self.choose_output).grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.output_label = tk.Label(top, text="(none)", anchor="w")
        self.output_label.grid(row=1, column=1, sticky="we", padx=(10, 0), pady=(6, 0))

        thresh_frame = tk.Frame(top)
        thresh_frame.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))
        tk.Label(thresh_frame, text="Length Δ threshold (raw units):").pack(side="left")
        self.length_entry = tk.Entry(thresh_frame, width=10, textvariable=self.length_delta_var)
        self.length_entry.pack(side="left", padx=(8, 0))
        tk.Label(thresh_frame, text="(example: 0.25)").pack(side="left", padx=(8, 0))

        merge_frame = tk.Frame(top)
        merge_frame.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))
        tk.Checkbutton(
            merge_frame,
            text="Merge eligible A/Z pairs into *_MergeMF.opm",
            variable=self.merge_var
        ).pack(side="left")

        top.columnconfigure(1, weight=1)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(0, 6))

        self.status_label = tk.Label(self, text="Ready.", anchor="w")
        self.status_label.pack(fill="x", padx=10)

        btns = tk.Frame(self)
        btns.pack(fill="x", padx=10, pady=(6, 0))

        self.convert_btn = tk.Button(btns, text="Convert JSON → OPM", command=self.convert)
        self.convert_btn.pack(side="left")

        self.export_btn = tk.Button(btns, text="Export Mismatch CSV", command=self.export_mismatches_csv, state="disabled")
        self.export_btn.pack(side="left", padx=(10, 0))

        self.log = scrolledtext.ScrolledText(self, height=26)
        self.log.pack(fill="both", expand=True, padx=10, pady=10)

        # Muted, readable colors (less “laser bright”)
        MUTED_ERR_BG = "#7a1f2b"   # muted brick red
        MUTED_SUM_BG = "#144a78"   # muted slate blue

        self.log.tag_configure("ok", foreground="#1a7f37")  # green

        # Error block (muted red)
        self.log.tag_configure("err", foreground="white", background=MUTED_ERR_BG)

        # Summary block (muted blue)
        self.log.tag_configure("sum", foreground="white", background=MUTED_SUM_BG, font=("Consolas", 11, "bold"))

        # Headers (plain, consistent for all sections)
        self.log.tag_configure("hdr", font=("Consolas", 11, "bold"))

        # Wavelength accent (slightly softer violet)
        self.log.tag_configure("wl", foreground="#c4b5fd", font=("Consolas", 11, "bold"))

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

    def _restore_merge_toggle(self):
        self.merge_var.set(bool(self.settings.get("merge_enabled", False)))

    def _persist_merge_toggle(self):
        self.settings["merge_enabled"] = bool(self.merge_var.get())
        _save_settings(self.settings)

    def _set_status(self, text: str):
        self.status_label.config(text=text)
        self.update_idletasks()

    # ----------------------------
    # Logging helpers
    # ----------------------------

    def _log(self, text: str = "", tag: str | None = None):
        if tag:
            self.log.insert(tk.END, text + "\n", tag)
        else:
            self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.update_idletasks()

    def _log_header_plain(self, title: str):
        # Always plain header: matches Errors and Summary
        self._log(title, "hdr")
        self._log("-" * len(title), "hdr")

    def _log_section_plain(self, title: str):
        self._log("")
        self._log_header_plain(title)

    def _clear_log(self):
        self.log.delete("1.0", tk.END)

    def _log_lambda_line(self, line: str, base_tag: str | None):
        if "λ" not in line:
            self._log(line, base_tag)
            return

        start_index = self.log.index(tk.END)
        self.log.insert(tk.END, line + "\n", base_tag if base_tag else None)

        try:
            pos = line.index("λ")
            base_line, base_col = map(int, start_index.split("."))
            lam_start = f"{base_line}.{base_col + pos}"
            lam_end = f"{base_line}.{base_col + pos + 1}"
            self.log.tag_add("wl", lam_start, lam_end)
        except Exception:
            pass

        self.log.see(tk.END)
        self.update_idletasks()

    # ----------------------------
    # Export mismatches
    # ----------------------------

    def export_mismatches_csv(self):
        if not self.mismatch_rows:
            messagebox.showinfo("No mismatches", "No mismatch rows available to export.")
            return

        out_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="az_mismatches.csv",
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
                    w.writerow({k: row.get(k, "") for k in fields})
            self._log(f"✅ EXPORTED mismatch CSV: {out_path}", "ok")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))

    # ----------------------------
    # A/Z helpers
    # ----------------------------

    def _extract_az_pair_key(self, stem: str):
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

    # ----------------------------
    # Merge helpers
    # ----------------------------

    def _worst_verdict(self, a: str | None, z: str | None) -> str:
        order = {"Fail": 2, "Unknown": 1, "Pass": 0}
        av = order.get(a, 1)
        zv = order.get(z, 1)
        worst = a if av >= zv else z
        if worst not in ("Fail", "Pass"):
            return "Unknown"
        return worst

    def _merge_opm_docs(self, a_doc: dict, z_doc: dict) -> dict:
        merged = copy.deepcopy(a_doc)

        a_od = merged.get("Measurement", {}).get("OpmResultData", {})
        z_od = z_doc.get("Measurement", {}).get("OpmResultData", {})

        a_meas = a_od.get("Measurements", [])
        z_meas = z_od.get("Measurements", [])
        if not isinstance(a_meas, list):
            a_meas = []
        if not isinstance(z_meas, list):
            z_meas = []

        combined: list[dict] = []
        for m in (a_meas + z_meas):
            if not isinstance(m, dict):
                continue
            m2 = copy.deepcopy(m)
            if "ResultState" not in m2:
                m2["ResultState"] = "Active"
            fl = m2.get("FiberLength")
            if isinstance(fl, dict) and "Origin" not in fl:
                fl["Origin"] = "Unknown"
            combined.append(m2)

        a_od["Measurements"] = combined
        if "AutoWavelength" not in a_od:
            a_od["AutoWavelength"] = False

        worst = self._worst_verdict(a_doc.get("GlobalVerdict"), z_doc.get("GlobalVerdict"))
        merged["GlobalVerdict"] = worst
        a_od["Status"] = worst
        return merged

    # ----------------------------
    # Analysis
    # ----------------------------

    def _analyze_az_pairs(self, produced_opm_paths: list[Path]) -> dict:
        pairs: dict[str, dict[str, Path]] = {}
        for p in produced_opm_paths:
            key, side = self._extract_az_pair_key(p.stem)
            if not key or side not in ("A", "Z"):
                continue
            pairs.setdefault(key, {})[side] = p

        length_threshold = self._get_length_threshold()

        stats = {
            "pairs_checked": 0,
            "mismatched_pairs": 0,
            "polarity_mismatches": 0,
            "wavelength_mismatches": 0,
            "length_issue_pairs": 0,
            "eligible_pairs": 0,
        }

        blocks: list[tuple[str, list[str]]] = []
        eligible_pairs: list[tuple[str, Path, Path]] = []
        mismatch_rows: list[dict] = []

        for key, sides in sorted(pairs.items()):
            if "A" not in sides or "Z" not in sides:
                continue

            stats["pairs_checked"] += 1

            try:
                a_doc = load_json(sides["A"])
                z_doc = load_json(sides["Z"])

                expected_pol = self._get_expected_polarity(a_doc) or self._get_expected_polarity(z_doc)
                a_pol = self._get_actual_polarity(a_doc)
                z_pol = self._get_actual_polarity(z_doc)
                polarity_mismatch = (a_pol != z_pol)

                a_wl = self._get_wavelengths_nm(a_doc)
                z_wl = self._get_wavelengths_nm(z_doc)
                wavelength_mismatch = (a_wl != z_wl)

                _, a_len_val = self._get_length_info(a_doc)
                _, z_len_val = self._get_length_info(z_doc)
                a_has_len = isinstance(a_len_val, (int, float))
                z_has_len = isinstance(z_len_val, (int, float))
                length_missing = (not a_has_len) or (not z_has_len)

                length_delta = None
                length_mismatch = False
                if a_has_len and z_has_len:
                    length_delta = abs(float(a_len_val) - float(z_len_val))
                    length_mismatch = (length_delta > length_threshold)

                if not (polarity_mismatch or wavelength_mismatch or length_missing or length_mismatch):
                    eligible_pairs.append((key, sides["A"], sides["Z"]))
                    continue

                stats["mismatched_pairs"] += 1
                lines: list[str] = []

                if polarity_mismatch:
                    stats["polarity_mismatches"] += 1
                    lines += [
                        "  🔀 | Polarity",
                        f"    Expected: {self._fmt(expected_pol)}",
                        f"    A: {self._fmt(a_pol)}",
                        f"    Z: {self._fmt(z_pol)}",
                        "",
                    ]

                if wavelength_mismatch:
                    stats["wavelength_mismatches"] += 1
                    lines += [
                        "  λ | Wavelength",
                        f"    A: {a_wl}",
                        f"    Z: {z_wl}",
                        "",
                    ]

                length_issue_on_pair = False

                if length_missing:
                    length_issue_on_pair = True
                    lines += [
                        "  📏 | Length",
                        f"    A: {self._fmt(a_len_val)}",
                        f"    Z: {self._fmt(z_len_val)}",
                        "",
                    ]

                if length_mismatch:
                    length_issue_on_pair = True
                    lines += [
                        "  📏 | Length",
                        f"    Length mismatch: A={self._fmt(float(a_len_val))}  Z={self._fmt(float(z_len_val))}",
                        f"    Delta: {self._fmt(length_delta)}   Threshold: {self._fmt(length_threshold)}",
                        "",
                    ]

                if length_issue_on_pair:
                    stats["length_issue_pairs"] += 1

                blocks.append((key, lines))

            except Exception as e:
                stats["mismatched_pairs"] += 1
                blocks.append((key, [f"  Compare error: {e}", ""]))

        stats["eligible_pairs"] = len(eligible_pairs)

        return {
            "stats": stats,
            "blocks": blocks,
            "eligible_pairs": eligible_pairs,
            "length_threshold": length_threshold,
        }

    def _merge_eligible_pairs(self, eligible_pairs: list[tuple[str, Path, Path]]) -> dict:
        merged_msgs: list[str] = []
        merge_write_errors: list[str] = []

        merged = 0
        write_errors = 0

        for pair_key, a_path, z_path in eligible_pairs:
            try:
                a_doc = load_json(a_path)
                z_doc = load_json(z_path)
                merged_doc = self._merge_opm_docs(a_doc, z_doc)

                out_path = self.output_dir / f"{a_path.stem}_MergeMF.opm"
                if out_path.exists():
                    raise FileExistsError(f"Output already exists: {out_path.name}")

                with out_path.open("w", encoding="utf-8") as f:
                    json.dump(merged_doc, f, indent=2)

                merged += 1
                merged_msgs.append(f"✅ MERGED     {pair_key}  ->  {out_path.name}")

            except Exception as e:
                write_errors += 1
                merge_write_errors.append(f"❌ MERGE ERR  {pair_key}  ->  {e}")

        return {
            "merged": merged,
            "write_errors": write_errors,
            "merged_msgs": merged_msgs,
            "merge_write_errors": merge_write_errors,
        }

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
        self._persist_merge_toggle()

        self._clear_log()
        self.export_btn.config(state="disabled")

        ok_msgs: list[str] = []
        err_msgs: list[str] = []

        self.convert_btn.config(state="disabled")
        try:
            total = len(json_files)
            self.progress["value"] = 0
            self.progress["maximum"] = total
            self._set_status(f"Converting 0 / {total}")

            produced_opm_paths: list[Path] = []
            success = 0
            failed = 0

            for i, src_path in enumerate(json_files, start=1):
                try:
                    src_json = load_json(src_path)
                    opm_json = map_pxm_json_to_opm(src_json)
                    out_path = self.output_dir / (src_path.stem + ".opm")

                    if out_path.exists():
                        raise FileExistsError(self._explain_duplicate_output(out_path, src_path))

                    with out_path.open("w", encoding="utf-8") as f:
                        json.dump(opm_json, f, indent=2)

                    produced_opm_paths.append(out_path)
                    success += 1
                    ok_msgs.append(f"✅ CONVERTED  {src_path.name}  ->  {out_path.name}")

                except Exception as e:
                    failed += 1
                    err_msgs.append(f"❌ FAILED     {src_path.name}  ->  {e}")

                self.progress["value"] = i
                self._set_status(f"Converting {i} / {total}")

            analysis = self._analyze_az_pairs(produced_opm_paths)
            az_stats = analysis["stats"]
            blocks = analysis["blocks"]
            eligible_pairs = analysis["eligible_pairs"]
            length_threshold = analysis["length_threshold"]

            merge_stats = {"merged": 0, "write_errors": 0}
            if self.merge_var.get() and self.output_dir:
                merge_out = self._merge_eligible_pairs(eligible_pairs)
                merge_stats["merged"] = merge_out["merged"]
                merge_stats["write_errors"] = merge_out["write_errors"]
                ok_msgs.extend(merge_out["merged_msgs"])
                err_msgs.extend(merge_out["merge_write_errors"])

            # Results (plain header, green lines)
            self._log_section_plain("Results")
            for m in ok_msgs:
                self._log(m, "ok")

            # Errors (plain header, red block lines)
            has_any_errors = bool(err_msgs) or bool(blocks) or failed > 0
            if has_any_errors:
                self._log_section_plain("Errors")
                for m in err_msgs:
                    self._log(m, "err")
                for key, lines in blocks:
                    self._log(f"❌ A/Z MISMATCH  {key}", "err")
                    for ln in lines:
                        if "λ" in ln:
                            self._log_lambda_line(ln, "err")
                        else:
                            self._log(ln, "err")
                    self._log("", "err")

            # Summary (plain header, blue block lines)
            self._log_section_plain("Summary")

            mismatched_pairs = az_stats.get("mismatched_pairs", 0)
            pairs_checked = az_stats.get("pairs_checked", 0)
            eligible = az_stats.get("eligible_pairs", 0)
            not_eligible = max(0, pairs_checked - eligible)

            summary_lines = [
                f"JSON Converted: {success}   Failed: {failed}",
                f"A/Z pairs checked: {pairs_checked}",
                f"A/Z mismatched pairs: {mismatched_pairs}",
                "",
                "Issue counts (pairs may include multiple issues):",
                f"  🔀  Polarity mismatches: {az_stats.get('polarity_mismatches', 0)}",
                f"  λ   Wavelength mismatches: {az_stats.get('wavelength_mismatches', 0)}",
                f"  📏  Length missing or mismatched: {az_stats.get('length_issue_pairs', 0)}",
                "",
                "Merge",
                f"  Eligible pairs to merge: {eligible} of {pairs_checked}   (threshold Δ={self._fmt(length_threshold)})",
            ]

            if self.merge_var.get():
                summary_lines += [
                    f"  ✅ Merged: {merge_stats.get('merged', 0)}",
                    f"  ⛔ Not eligible: {not_eligible}",
                    f"  ⚠ Merge write errors: {merge_stats.get('write_errors', 0)}",
                ]
            else:
                summary_lines += [
                    "  Merged: (disabled)",
                    f"  ⛔ Not eligible: {not_eligible}",
                ]

            for ln in summary_lines:
                if "λ" in ln:
                    self._log_lambda_line(ln, "sum")
                else:
                    self._log(ln, "sum")

            self._set_status(f"Done. Success: {success}  Failed: {failed}")

        finally:
            self.convert_btn.config(state="normal")
