import json
import csv
import copy
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter import scrolledtext
from datetime import datetime

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
        self.geometry("1040x760")

        self.settings = _load_settings()
        self.input_dir: Path | None = None
        self.output_dir: Path | None = None

        # Analyze-only folder
        self.opm_results_dir: Path | None = None

        # Options
        self.length_delta_var = tk.StringVar()
        self.merge_var = tk.BooleanVar(value=False)
        self.generate_punch_var = tk.BooleanVar(value=False)

        # Last-run data
        self.last_punch_rows: list[dict] = []
        self.last_punch_path: Path | None = None

        self._build_ui()
        self._restore_last_paths()
        self._restore_length_threshold()
        self._restore_merge_toggle()
        self._restore_punch_toggle()

    def _build_merged_opm_filename(self, a_path: "Path", z_path: "Path") -> str:
        """
        Build merged filename strictly from the normalized A and Z filenames.

        Expected stems:
        A: P1_A03_C01_LC01_NS1_ROW_04_RACK_30_RU_48
        Z: P1_Z03_C01_LC02_DHB_ROW_13_RACK_C_RU_10

        Output (example):
        P1_03_C01_LC01_NS1__LC02_DHB_ROW_04_RACK_30_RU_48_ROW_13_RACK_C_RU_10_MergeMF.opm
        """
        import re

        a_stem = a_path.stem
        z_stem = z_path.stem

        # Pull: P#, side, enclosure(2 digits), connector(2 digits), rest
        # Example: P1_A03_C01_<rest...>
        rx = r"^(P\d+)_([AZ])(\d{2})_C(\d{2})_(.+)$"

        ma = re.match(rx, a_stem)
        mz = re.match(rx, z_stem)

        if not ma or not mz:
            # fallback if unexpected
            return f"{a_stem}__{z_stem}_MergeMF.opm"

        p_a, side_a, enc_a, c_a, rest_a = ma.groups()
        p_z, side_z, enc_z, c_z, rest_z = mz.groups()

        # Pair key is project + enclosure + connector (but without side letter)
        # P1_03_C01
        pair_key = f"{p_a}_{enc_a}_C{c_a}"

        # Build merged name: <pair_key>_<Arest>__<Zrest>_MergeMF.opm
        merged_stem = f"{pair_key}_{rest_a}__{rest_z}_MergeMF"

        # Safe tokenization: keep underscores/hyphens, remove illegal filename chars
        def _safe(s: str) -> str:
            bad = '<>:"/\\|?*'
            for ch in bad:
                s = s.replace(ch, "_")
            return s.strip().strip(".")

        return _safe(merged_stem) + ".opm"

    # ----------------------------
    # UI
    # ----------------------------

    def _build_ui(self):
        top = tk.Frame(self)
        top.pack(fill="x", padx=10, pady=10)

        # Input/Output (Convert mode)
        tk.Button(top, text="Select Input Folder (JSON)", command=self.choose_input).grid(
            row=0, column=0, sticky="w"
        )
        self.input_label = tk.Label(top, text="(none)", anchor="w")
        self.input_label.grid(row=0, column=1, sticky="we", padx=(10, 0))

        tk.Button(top, text="Select Output Folder", command=self.choose_output).grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )
        self.output_label = tk.Label(top, text="(none)", anchor="w")
        self.output_label.grid(row=1, column=1, sticky="we", padx=(10, 0), pady=(6, 0))

        # Analyze-only (OPM folder)
        tk.Button(top, text="Select OPM Results Folder", command=self.choose_opm_results).grid(
            row=2, column=0, sticky="w", pady=(6, 0)
        )
        self.opm_label = tk.Label(top, text="(none)", anchor="w")
        self.opm_label.grid(row=2, column=1, sticky="we", padx=(10, 0), pady=(6, 0))

        # Length threshold
        thresh_frame = tk.Frame(top)
        thresh_frame.grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))
        tk.Label(thresh_frame, text="Length Δ threshold (raw units):").pack(side="left")
        self.length_entry = tk.Entry(thresh_frame, width=10, textvariable=self.length_delta_var)
        self.length_entry.pack(side="left", padx=(8, 0))
        tk.Label(thresh_frame, text="(example: 0.25)").pack(side="left", padx=(8, 0))

        # Merge toggle
        merge_frame = tk.Frame(top)
        merge_frame.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))
        tk.Checkbutton(
            merge_frame,
            text="Merge eligible A/Z pairs into *_MergeMF.opm",
            variable=self.merge_var
        ).pack(side="left")

        # Punch list toggle
        punch_frame = tk.Frame(top)
        punch_frame.grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))
        tk.Checkbutton(
            punch_frame,
            text="Export Punch List CSV (issues only)",
            variable=self.generate_punch_var
        ).pack(side="left")

        top.columnconfigure(1, weight=1)

        # Progress + status
        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(0, 6))

        self.status_label = tk.Label(self, text="Ready.", anchor="w")
        self.status_label.pack(fill="x", padx=10)

        # Buttons
        btns = tk.Frame(self)
        btns.pack(fill="x", padx=10, pady=(6, 0))

        self.convert_btn = tk.Button(btns, text="Convert JSON → OPM", command=self.convert)
        self.convert_btn.pack(side="left")

        self.analyze_btn = tk.Button(btns, text="Analyze OPM Folder", command=self.analyze_opm_folder)
        self.analyze_btn.pack(side="left", padx=(10, 0))

        self.export_punch_btn = tk.Button(
            btns, text="Export Last Punch List CSV", command=self.export_last_punch_csv, state="disabled"
        )
        self.export_punch_btn.pack(side="left", padx=(10, 0))

        # Log
        self.log = scrolledtext.ScrolledText(self, height=30)
        self.log.pack(fill="both", expand=True, padx=10, pady=10)

        self._make_text_readonly_but_copyable(self.log)

        # Muted, readable colors (less “laser bright”)
        MUTED_ERR_BG = "#7a1f2b"   # muted brick red
        MUTED_SUM_BG = "#144a78"   # muted slate blue

        self.log.tag_configure("ok", foreground="#1a7f37")  # green
        self.log.tag_configure("err", foreground="white", background=MUTED_ERR_BG)
        self.log.tag_configure("sum", foreground="white", background=MUTED_SUM_BG, font=("Consolas", 11, "bold"))
        self.log.tag_configure("hdr", font=("Consolas", 11, "bold"))
        self.log.tag_configure("wl", foreground="#c4b5fd", font=("Consolas", 11, "bold"))  # violet-ish lambda

    # ----------------------------
    # Folder pickers / settings
    # ----------------------------

    def choose_input(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_dir = Path(folder)
            self.input_label.config(text=str(self.input_dir))
            self._persist_paths()

    def choose_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir = Path(folder)
            self.output_label.config(text=str(self.output_dir))
            self._persist_paths()

    def choose_opm_results(self):
        folder = filedialog.askdirectory()
        if folder:
            self.opm_results_dir = Path(folder)
            self.opm_label.config(text=str(self.opm_results_dir))
            self.settings["last_opm_results_dir"] = str(self.opm_results_dir)
            _save_settings(self.settings)

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

        last_opm = self.settings.get("last_opm_results_dir")
        if last_opm and Path(last_opm).exists():
            self.opm_results_dir = Path(last_opm)
            self.opm_label.config(text=str(self.opm_results_dir))

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

    def _restore_punch_toggle(self):
        self.generate_punch_var.set(bool(self.settings.get("generate_punch_csv", False)))

    def _persist_punch_toggle(self):
        self.settings["generate_punch_csv"] = bool(self.generate_punch_var.get())
        _save_settings(self.settings)

    # ----------------------------
    # Logging helpers
    # ----------------------------

    def _set_status(self, text: str):
        self.status_label.config(text=text)
        self.update_idletasks()

    def _clear_log(self):
        self.log.delete("1.0", tk.END)

    def _log(self, text: str = "", tag: str | None = None):
        if tag:
            self.log.insert(tk.END, text + "\n", tag)
        else:
            self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.update_idletasks()

    def _log_header_plain(self, title: str):
        self._log(title, "hdr")
        self._log("-" * len(title), "hdr")

    def _log_section_plain(self, title: str):
        self._log("")
        self._log_header_plain(title)

    def _log_lambda_line(self, line: str, base_tag: str | None):
        # apply special styling to λ character inside a line
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
    # Core: Convert
    # ----------------------------

    def _explain_duplicate_output(self, out_path: Path, src_path: Path) -> str:
        return (
            f"Output file already exists and overwrite is disabled:\n"
            f"  Output: {out_path}\n"
            f"  Input : {src_path}\n\n"
            f"Fix: choose an empty output folder, delete/rename the existing .opm file, or rename the input JSON(s) "
            f"so they produce unique output names."
        )

    def _idents_to_map(self, src_json: dict) -> dict[str, str]:
        """
        Exchange JSON has identifiers as a list of {Name, Value} under brief.Identifiers.
        Return a simple dict mapping Name -> Value (string).
        """
        out: dict[str, str] = {}
        brief = src_json.get("brief")
        if not isinstance(brief, dict):
            return out

        ids = brief.get("Identifiers")
        if not isinstance(ids, list):
            return out

        for item in ids:
            if not isinstance(item, dict):
                continue
            name = item.get("Name")
            val = item.get("Value")
            if isinstance(name, str) and val is not None:
                out[name] = str(val)

        return out


    def _safe_token(self, s: str) -> str:
        """
        Windows-safe token for filenames. No ?, :, *, etc. Also strips spaces.
        """
        if s is None:
            return "NA"
        s = str(s).strip()
        if not s:
            return "NA"

        # Replace invalid filename chars
        for ch in ['<', '>', ':', '"', '/', '\\', '|', '?', '*']:
            s = s.replace(ch, "_")

        # Collapse whitespace
        s = s.replace(" ", "_")
        while "__" in s:
            s = s.replace("__", "_")

        return s.strip("._") or "NA"


    def _split_side_tokens(self, value: str, side: str, expect_pairs: int) -> str:
        """
        Split a concatenated identifier like:
        BLDG_RM: LC01_NS1_LC02_DHB  (expect_pairs=2) -> A gets first 2 tokens, Z gets last 2 tokens
        ROW_RACK_RU: 04_30_48_13_C_10 (expect_pairs=3) -> A gets first 3 tokens, Z gets last 3 tokens

        If token count isn't 2*expect_pairs, we just return the whole thing.
        """
        if not value:
            return "NA"
        toks = [t for t in str(value).split("_") if t]

        if len(toks) == 2 * expect_pairs:
            if side.upper() == "A":
                use = toks[:expect_pairs]
            else:
                use = toks[expect_pairs:]
            return "_".join(use)

        # unexpected format; return original (still safe-tokenized later)
        return "_".join(toks) if toks else "NA"


    def _build_normalized_output_name(self, src_json: dict, src_path: Path) -> str:
        """
        Build output OPM filename WITHOUT changing JSON content.

        Rules:
        - PATH/TRUNK + A/Z + Cxx come from the source JSON filename (src_path.stem)
        Example: P1_Z03_C01_...
        - Location + ROW/RACK/RU come from identifiers but split by side:
        BLDG_RM: A=first half, Z=second half
        ROW_RACK_RU: A=first half, Z=second half
        """
        stem = src_path.stem

        # Determine side from the source filename
        # e.g. "P1_A03_C01_..." or "P1_Z03_C01_..."
        m = re.match(r"^(P\d+_[AZ]\d+_C\d{2})_", stem)
        prefix = m.group(1) if m else stem.split("_")[0]  # fallback

        side = "A"
        m2 = re.match(r"^P\d+_([AZ])\d+_C\d{2}", prefix)
        if m2:
            side = m2.group(1).upper()

        ids = self._idents_to_map(src_json)
        bldg_rm_full = ids.get("BLDG_RM", "")
        rrr_full = ids.get("ROW_RACK_RU", "")

        bldg_rm_side = self._split_side_tokens(bldg_rm_full, side=side, expect_pairs=2)
        rrr_side = self._split_side_tokens(rrr_full, side=side, expect_pairs=3)

        # rrr_side is like "04_30_48" or "13_C_10"
        rrr_toks = rrr_side.split("_") if rrr_side else []
        row = rrr_toks[0] if len(rrr_toks) > 0 else "NA"
        rack = rrr_toks[1] if len(rrr_toks) > 1 else "NA"
        ru = rrr_toks[2] if len(rrr_toks) > 2 else "NA"

        filename = (
            f"{prefix}_"
            f"{bldg_rm_side}_"
            f"ROW_{row}_RACK_{rack}_RU_{ru}"
            f".opm"
        )

        return self._safe_token(filename)

    def convert(self):
        if not self.input_dir or not self.output_dir:
            messagebox.showerror("Missing folder", "Please select both input (JSON) and output folders.")
            return

        json_files = list(self.input_dir.glob("*.json"))
        if not json_files:
            messagebox.showwarning("No files", "No JSON files found in input folder.")
            return

        self._persist_paths()
        self._persist_length_threshold()
        self._persist_merge_toggle()
        self._persist_punch_toggle()

        self._clear_log()
        self.export_punch_btn.config(state="disabled")
        self.last_punch_rows = []
        self.last_punch_path = None

        ok_msgs: list[str] = []
        err_msgs: list[str] = []

        self.convert_btn.config(state="disabled")
        try:
            total = len(json_files)
            self.progress["value"] = 0
            self.progress["maximum"] = total
            self._set_status(f"Converting 0 / {total}")

            produced_opm_paths: list[Path] = []
            json_ok = 0
            json_fail = 0

            for i, src_path in enumerate(json_files, start=1):
                try:
                    src_json = load_json(src_path)
                    opm_json = map_pxm_json_to_opm(src_json)

                    out_name = self._build_normalized_output_name(src_json, src_path)
                    out_path = self.output_dir / out_name

                    if out_path.exists():
                        raise FileExistsError(self._explain_duplicate_output(out_path, src_path))

                    with out_path.open("w", encoding="utf-8") as f:
                        json.dump(opm_json, f, indent=2)

                    produced_opm_paths.append(out_path)
                    json_ok += 1
                    ok_msgs.append(f"✅ CONVERTED  {src_path.name}  ->  {out_path.name}")

                except Exception as e:
                    json_fail += 1
                    err_msgs.append(f"❌ FAILED     {src_path.name}  ->  {e}")

                self.progress["value"] = i
                self._set_status(f"Converting {i} / {total}")

            # Analyze A/Z from produced outputs
            analysis = self._analyze_pairs_from_opm_paths(produced_opm_paths)
            stats = analysis["stats"]
            error_blocks = analysis["error_blocks"]
            eligible_pairs = analysis["eligible_pairs"]
            punch_rows = analysis["punch_rows"]
            length_threshold = analysis["length_threshold"]

            # Merge (optional)
            merge_stats = {"merged": 0, "write_errors": 0, "not_eligible": 0}
            if self.merge_var.get():
                merge_out = self._merge_eligible_pairs(eligible_pairs, self.output_dir)
                merge_stats["merged"] = merge_out["merged"]
                merge_stats["write_errors"] = merge_out["write_errors"]
                ok_msgs.extend(merge_out["merged_msgs"])
                err_msgs.extend(merge_out["merge_write_errors"])

            pairs_checked = stats.get("pairs_checked", 0)
            eligible = stats.get("eligible_pairs", 0)
            merge_stats["not_eligible"] = max(0, pairs_checked - eligible)

            # Punch list (optional) - CSV only
            punch_out_path = None
            if self.generate_punch_var.get():
                if punch_rows:
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    ose = self._guess_ose_from_any(punch_rows) or "OSE"
                    punch_out_path = self.output_dir / f"Punch List - {ose} - {ts}.csv"
                    try:
                        self._write_punch_list_csv(punch_out_path, punch_rows)
                        ok_msgs.append(f"✅ PUNCH LIST  Created  ->  {punch_out_path.name}")
                        self.last_punch_rows = punch_rows
                        self.last_punch_path = punch_out_path
                        self.export_punch_btn.config(state="normal")
                    except Exception as e:
                        err_msgs.append(f"❌ PUNCH LIST  Failed to write CSV: {e}")
                else:
                    ok_msgs.append("✅ PUNCH LIST  No issues found; no Punch List created.")

            # RESULTS (top)
            self._log_section_plain("Results")
            for m in ok_msgs:
                self._log(m, "ok")

            # ERRORS
            has_any_errors = bool(err_msgs) or bool(error_blocks) or json_fail > 0
            if has_any_errors:
                self._log_section_plain("Errors")
                for m in err_msgs:
                    self._log(m, "err")

                for block in error_blocks:
                    for ln in block:
                        if "λ" in ln:
                            self._log_lambda_line(ln, "err")
                        else:
                            self._log(ln, "err")
                    self._log("", "err")

            # SUMMARY
            self._log_section_plain("Summary")

            summary_lines = [
                f"JSON Converted: {json_ok}   Failed: {json_fail}",
                f"A/Z pairs checked: {pairs_checked}",
                "",
                "Issue counts (pairs may include multiple issues):",
                f"  🔥  High Loss failures: {stats.get('high_loss_pairs', 0)}",
                f"  🔀  Polarity issues (missing/mismatch/unknown): {stats.get('polarity_issue_pairs', 0)}",
                f"  λ  Wavelength mismatches: {stats.get('wavelength_mismatches', 0)}",
                f"  📏  Length missing/mismatched: {stats.get('length_issue_pairs', 0)}",
                "",
                "Merge",
                f"  Eligible pairs to merge: {eligible} of {pairs_checked}   (threshold Δ={self._fmt(length_threshold)})",
                f"  ✅  Merged: {merge_stats['merged']}" if self.merge_var.get() else "  (merge disabled)",
            ]
            if self.merge_var.get():
                summary_lines += [
                    f"  ⛔  Not eligible: {merge_stats['not_eligible']}",
                    f"  ⚠  Merge write errors: {merge_stats['write_errors']}",
                ]

            if self.generate_punch_var.get():
                if punch_out_path:
                    summary_lines += ["", f"Punch List: {punch_out_path.name}"]
                else:
                    summary_lines += ["", "Punch List: (not created)"]

            for ln in summary_lines:
                if "λ" in ln:
                    self._log_lambda_line(ln, "sum")
                else:
                    self._log(ln, "sum")

            self._set_status(f"Done. Converted: {json_ok}  Failed: {json_fail}")

        finally:
            self.convert_btn.config(state="normal")

    def _make_text_readonly_but_copyable(self, text):
        """
        Make a tk.Text / ScrolledText widget behave like a read-only log:
        - user can select and copy
        - user cannot type/edit
        - supports Ctrl+C, Ctrl+A, and right-click Copy
        """
        # Block typing/editing (but allow selection)
        text.bind("<Key>", lambda e: "break")

        # Ctrl+C (copy selection)
        text.bind("<Control-c>", lambda e: self._copy_selection(text))
        text.bind("<Control-C>", lambda e: self._copy_selection(text))

        # Ctrl+A (select all)
        text.bind("<Control-a>", lambda e: self._select_all(text))
        text.bind("<Control-A>", lambda e: self._select_all(text))

        # Right-click menu (Windows)
        text.bind("<Button-3>", lambda e: self._show_copy_menu(e, text))

    def _copy_selection(self, text):
        try:
            selection = text.get("sel.first", "sel.last")
        except Exception:
            return "break"
        text.clipboard_clear()
        text.clipboard_append(selection)
        return "break"


    def _select_all(self, text):
        text.tag_add("sel", "1.0", "end-1c")
        return "break"


    def _show_copy_menu(self, event, text):
        menu = tk.Menu(text, tearoff=0)
        menu.add_command(label="Copy", command=lambda: self._copy_selection(text))
        menu.add_command(label="Select All", command=lambda: self._select_all(text))
        menu.tk_popup(event.x_root, event.y_root)

    # ----------------------------
    # Analyze-only mode
    # ----------------------------

    def analyze_opm_folder(self):
        if not self.opm_results_dir or not self.opm_results_dir.exists():
            messagebox.showerror("Missing folder", "Please select an OPM results folder first.")
            return

        out_dir = self.output_dir or self.opm_results_dir

        self._persist_length_threshold()
        self._persist_merge_toggle()
        self._persist_punch_toggle()

        opm_files = list(self.opm_results_dir.glob("*.opm"))
        if not opm_files:
            messagebox.showwarning("No files", "No .opm files found in the selected results folder.")
            return

        self._clear_log()
        self.export_punch_btn.config(state="disabled")
        self.last_punch_rows = []
        self.last_punch_path = None

        ok_msgs: list[str] = []
        err_msgs: list[str] = []

        self.analyze_btn.config(state="disabled")
        try:
            self.progress["value"] = 0
            self.progress["maximum"] = len(opm_files)
            self._set_status("Scanning OPM files...")

            for i, _ in enumerate(opm_files, start=1):
                self.progress["value"] = i
                if i % 25 == 0 or i == len(opm_files):
                    self._set_status(f"Found {i} / {len(opm_files)} OPM files")

            analysis = self._analyze_pairs_from_opm_paths(opm_files)
            stats = analysis["stats"]
            error_blocks = analysis["error_blocks"]
            eligible_pairs = analysis["eligible_pairs"]
            punch_rows = analysis["punch_rows"]
            length_threshold = analysis["length_threshold"]

            # Merge (optional)
            merge_stats = {"merged": 0, "write_errors": 0, "not_eligible": 0}
            if self.merge_var.get():
                merge_out = self._merge_eligible_pairs(eligible_pairs, out_dir)
                merge_stats["merged"] = merge_out["merged"]
                merge_stats["write_errors"] = merge_out["write_errors"]
                ok_msgs.extend(merge_out["merged_msgs"])
                err_msgs.extend(merge_out["merge_write_errors"])

            pairs_checked = stats.get("pairs_checked", 0)
            eligible = stats.get("eligible_pairs", 0)
            merge_stats["not_eligible"] = max(0, pairs_checked - eligible)

            # Punch list (optional) - CSV
            punch_out_path = None
            if self.generate_punch_var.get():
                if punch_rows:
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    ose = self._guess_ose_from_any(punch_rows) or "OSE"
                    punch_out_path = out_dir / f"Punch List - {ose} - {ts}.csv"
                    try:
                        self._write_punch_list_csv(punch_out_path, punch_rows)
                        ok_msgs.append(f"✅ PUNCH LIST  Created  ->  {punch_out_path.name}")
                        self.last_punch_rows = punch_rows
                        self.last_punch_path = punch_out_path
                        self.export_punch_btn.config(state="normal")
                    except Exception as e:
                        err_msgs.append(f"❌ PUNCH LIST  Failed to write CSV: {e}")
                else:
                    ok_msgs.append("✅ PUNCH LIST  No issues found; no Punch List created.")

            # RESULTS
            self._log_section_plain("Results")
            for m in ok_msgs:
                self._log(m, "ok")

            # ERRORS
            has_any_errors = bool(err_msgs) or bool(error_blocks)
            if has_any_errors:
                self._log_section_plain("Errors")
                for m in err_msgs:
                    self._log(m, "err")

                for block in error_blocks:
                    for ln in block:
                        if "λ" in ln:
                            self._log_lambda_line(ln, "err")
                        else:
                            self._log(ln, "err")
                    self._log("", "err")

            # SUMMARY
            self._log_section_plain("Summary")

            summary_lines = [
                f"JSON Converted: 0   Failed: 0",
                f"A/Z pairs checked: {pairs_checked}",
                "",
                "Issue counts (pairs may include multiple issues):",
                f"  🔥  High Loss failures: {stats.get('high_loss_pairs', 0)}",
                f"  🔀  Polarity mismatches/unknown: {stats.get('polarity_issue_pairs', 0)}",
                f"  λ  Wavelength mismatches: {stats.get('wavelength_mismatches', 0)}",
                f"  📏  Length missing/mismatched: {stats.get('length_issue_pairs', 0)}",
                "",
                "Merge",
                f"  Eligible pairs to merge: {eligible} of {pairs_checked}   (threshold Δ={self._fmt(length_threshold)})",
                f"  ✅  Merged: {merge_stats['merged']}" if self.merge_var.get() else "  (merge disabled)",
            ]
            if self.merge_var.get():
                summary_lines += [
                    f"  ⛔  Not eligible: {merge_stats['not_eligible']}",
                    f"  ⚠  Merge write errors: {merge_stats['write_errors']}",
                ]

            if self.generate_punch_var.get():
                if punch_out_path:
                    summary_lines += ["", f"Punch List: {punch_out_path.name}"]
                else:
                    summary_lines += ["", "Punch List: (not created)"]

            for ln in summary_lines:
                if "λ" in ln:
                    self._log_lambda_line(ln, "sum")
                else:
                    self._log(ln, "sum")

            self._set_status("Done.")

        finally:
            self.analyze_btn.config(state="normal")

    # ----------------------------
    # Export last punch list again
    # ----------------------------

    def export_last_punch_csv(self):
        if not self.last_punch_rows:
            messagebox.showinfo("Nothing to export", "No punch list rows are available from the last run.")
            return

        fp = filedialog.asksaveasfilename(
            title="Save Punch List CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not fp:
            return

        try:
            self._write_punch_list_csv(Path(fp), self.last_punch_rows)
            messagebox.showinfo("Saved", f"Punch List saved:\n{fp}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))

    # ----------------------------
    # A/Z Pairing + analysis logic
    # ----------------------------

    def _extract_az_pair_key(self, stem: str) -> tuple[str | None, str | None]:
        """
        Extract a stable A/Z pairing key from an OPM filename stem.

        Works with normalized filenames that start with:
        P1_A03_C01_...
        P1_Z03_C01_...

        Returns:
        (pair_key, side) where side is "A" or "Z"
        pair_key example: "P1_03_C01"
        """
        import re

        m = re.match(r"^(P\d+)_([AZ])(\d{2})_C(\d{2})(?:_|$)", stem)
        if not m:
            return None, None

        path = m.group(1)          # P1
        side = m.group(2)          # A or Z
        trunk = m.group(3)         # 03
        conn = m.group(4)          # 01

        key = f"{path}_{trunk}_C{conn}"
        return key, side

    def _analyze_pairs_from_opm_paths(self, opm_paths: list[Path]) -> dict:
        pairs: dict[str, dict[str, Path]] = {}
        for p in opm_paths:
            key, side = self._extract_az_pair_key(p.stem)
            if not key or side not in ("A", "Z"):
                continue
            pairs.setdefault(key, {})[side] = p

        length_threshold = self._get_length_threshold()

        stats = {
            "pairs_checked": 0,

            # merge-blocking mismatches ONLY
            "mismatched_pairs": 0,
            "polarity_issue_pairs": 0,     # missing OR mismatch OR PolarityStatus unknown
            "wavelength_mismatches": 0,
            "length_issue_pairs": 0,

            # failure-only (not merge-blocking)
            "high_loss_pairs": 0,

            "eligible_pairs": 0,
        }

        error_blocks: list[list[str]] = []
        eligible_pairs: list[tuple[str, Path, Path]] = []
        punch_pair_docs: list[dict] = []
        punch_rows: list[dict] = []

        for key, sides in sorted(pairs.items()):
            if "A" not in sides or "Z" not in sides:
                continue

            stats["pairs_checked"] += 1

            try:
                a_doc = load_json(sides["A"])
                z_doc = load_json(sides["Z"])
                punch_pair_docs.append({"pair_key": key, "A_doc": a_doc, "Z_doc": z_doc})

                # -------- High loss (failure-only) --------
                a_high_loss = self._has_high_loss(a_doc)
                z_high_loss = self._has_high_loss(z_doc)
                high_loss = a_high_loss or z_high_loss
                if high_loss:
                    stats["high_loss_pairs"] += 1

                # -------- Polarity (merge blocker) --------
                expected_pol = self._get_expected_polarity(a_doc) or self._get_expected_polarity(z_doc)
                a_pol = self._get_actual_polarity(a_doc)
                z_pol = self._get_actual_polarity(z_doc)
                pol_status_a = self._get_polarity_status(a_doc)
                pol_status_z = self._get_polarity_status(z_doc)

                polarity_missing = (a_pol is None) or (z_pol is None)
                polarity_mismatch = (a_pol is not None and z_pol is not None and a_pol != z_pol)
                polarity_unknown = (pol_status_a == "Unknown") or (pol_status_z == "Unknown")

                polarity_issue = polarity_missing or polarity_mismatch or polarity_unknown
                if polarity_issue:
                    stats["polarity_issue_pairs"] += 1

                # -------- Wavelength (merge blocker) --------
                a_wl = self._get_wavelengths_nm(a_doc)
                z_wl = self._get_wavelengths_nm(z_doc)
                wavelength_mismatch = (a_wl != z_wl)
                if wavelength_mismatch:
                    stats["wavelength_mismatches"] += 1

                # -------- Length (merge blocker) --------
                a_len, a_missing = self._get_length_numeric_or_missing(a_doc)
                z_len, z_missing = self._get_length_numeric_or_missing(z_doc)
                length_missing = a_missing or z_missing

                length_delta = None
                length_mismatch = False
                if a_len is not None and z_len is not None:
                    length_delta = abs(float(a_len) - float(z_len))
                    length_mismatch = (length_delta > length_threshold)

                length_issue = length_missing or length_mismatch
                if length_issue:
                    stats["length_issue_pairs"] += 1

                # -------- Merge blockers (your definition) --------
                # Merge allowed ONLY when there are NO polarity/length/wavelength mismatches.
                has_merge_blocker = polarity_issue or length_issue or wavelength_mismatch

                # Punch row if any issue (failure or mismatch)
                if has_merge_blocker or high_loss:
                    punch_rows.append(self._build_punch_row(
                        pair_key=key,
                        a_doc=a_doc,
                        z_doc=z_doc,
                        mismatch=has_merge_blocker,
                        failure_high_loss=high_loss,
                        a_fail=a_high_loss,
                        z_fail=z_high_loss,
                        polarity_issue=polarity_issue,
                        wavelength_issue=wavelength_mismatch,
                        length_issue=length_issue,
                    ))

                # -------- Render errors --------
                # FAILURE-only (no merge blocker)
                if high_loss and not has_merge_blocker:
                    side_txt = "A" if a_high_loss and not z_high_loss else "Z" if z_high_loss and not a_high_loss else "A+Z"
                    error_blocks.append([
                        f"❌ A/Z FAILURE   {key}",
                        "  🔥 | High Loss",
                        f"    Side: {side_txt}",
                        "    One or more readings have Status=Fail",
                        "",
                    ])
                    # still eligible to merge (your rule)
                    eligible_pairs.append((key, sides["A"], sides["Z"]))
                    continue

                # No issues at all → eligible
                if not has_merge_blocker and not high_loss:
                    eligible_pairs.append((key, sides["A"], sides["Z"]))
                    continue

                # MISMATCH (merge blocker)
                if has_merge_blocker:
                    stats["mismatched_pairs"] += 1
                    lines: list[str] = [f"❌ A/Z MISMATCH  {key}"]

                    if high_loss:
                        side_txt = "A" if a_high_loss and not z_high_loss else "Z" if z_high_loss and not a_high_loss else "A+Z"
                        lines += [
                            "  🔥 | High Loss",
                            f"    Side: {side_txt}",
                            "    One or more readings have Status=Fail",
                            "",
                        ]

                    if polarity_issue:
                        lines += [
                            "  🔀 | Polarity",
                            f"    Expected: {self._fmt(expected_pol)}",
                            f"    A: {self._fmt(a_pol)}",
                            f"    Z: {self._fmt(z_pol)}",
                            f"    Status: A={self._fmt(pol_status_a)}  Z={self._fmt(pol_status_z)}",
                            "",
                        ]

                    if wavelength_mismatch:
                        lines += [
                            "  λ | Wavelength",
                            f"    A: {a_wl}",
                            f"    Z: {z_wl}",
                            "",
                        ]

                    if length_missing:
                        lines += [
                            "  📏 | Length",
                            f"    A: {self._fmt(a_len)}",
                            f"    Z: {self._fmt(z_len)}",
                            "",
                        ]

                    if length_mismatch:
                        lines += [
                            "  📏 | Length",
                            f"    Length mismatch: A={self._fmt(float(a_len))}  Z={self._fmt(float(z_len))}",
                            f"    Delta: {self._fmt(length_delta)}   Threshold: {self._fmt(length_threshold)}",
                            "",
                        ]

                    error_blocks.append(lines)
                    continue

            except Exception as e:
                stats["mismatched_pairs"] += 1
                error_blocks.append([f"❌ A/Z MISMATCH  {key}", f"  Compare error: {e}", ""])
                punch_rows.append(self._build_punch_row(
                    pair_key=key,
                    a_doc=None,
                    z_doc=None,
                    mismatch=True,
                    failure_high_loss=False,
                    a_fail=False,
                    z_fail=False,
                    polarity_issue=False,
                    wavelength_issue=False,
                    length_issue=False,
                    note=f"Compare error: {e}",
                ))

        stats["eligible_pairs"] = len(eligible_pairs)

        return {
            "stats": stats,
            "error_blocks": error_blocks,
            "eligible_pairs": eligible_pairs,
            "punch_pair_docs": punch_pair_docs,
            "punch_rows": punch_rows,
            "length_threshold": length_threshold,
        }
        
    def _build_normalized_opm_filename_from_src(self, src_doc: dict, src_stem: str) -> str:
        """
        Build normalized OPM filename using the INPUT JSON (src_doc) identifiers.
        Does NOT modify any JSON data. Filename only.

        Target format:
        {ENCLOSURE}_{PATH_TRUNK}_{CONNECTOR}_{BLDG_RM_SIDE}_ROW_{row}_RACK_{rack}_RU_{ru}.opm

        Where:
        - ENCLOSURE / PATH_TRUNK / CONNECTOR come from src_stem (reliable in your dataset)
        - BLDG_RM and ROW/RACK/RU come from concatenated identifiers when available
            (BLDG_RM: LC01_NS1_LC02_DHB ; ROW_RACK_RU: 04_30_48_13_C_10)
        - side is inferred from PATH_TRUNK in src_stem (contains _Axx_ or _Zxx_)
        """
        import re

        def safe(s: str) -> str:
            s = str(s) if s is not None else ""
            for ch in '<>:"/\\|?*':
                s = s.replace(ch, "-")
            return s.strip() or "NA"

        # ---- Parse src_stem for enclosure / side / trunk / connector ----
        # Example: P1_A03_C01_LC01_NS1_ROW_04_RACK_30_RU_48
        m = re.match(r"^([A-Za-z0-9]+)_([AZ])(\d+)_C(\d+)", src_stem, re.IGNORECASE)
        if not m:
            # If somehow stem doesn't match, fall back to the original stem + .opm safely.
            return safe(src_stem) + ".opm"

        enclosure = m.group(1)
        side = m.group(2).upper()
        trunk_num = m.group(3)
        connector = f"C{m.group(4)}"

        path_trunk = f"{enclosure}_{side}{trunk_num}"

        # ---- Pull concatenated identifiers from INPUT JSON ----
        # We will NOT modify them; just read.
        identifiers = None
        if isinstance(src_doc, dict):
            b = src_doc.get("brief")
            if isinstance(b, dict):
                identifiers = b.get("Identifiers")
        if not isinstance(identifiers, dict):
            identifiers = {}

        bldg_rm_concat = identifiers.get("BLDG_RM")
        row_rack_ru_concat = identifiers.get("ROW_RACK_RU")

        # ---- Derive BLDG_RM for the correct side ----
        # LC01_NS1_LC02_DHB -> A=LC01_NS1 ; Z=LC02_DHB
        bldg_rm_side = None
        if isinstance(bldg_rm_concat, str) and bldg_rm_concat.strip():
            parts = [p for p in bldg_rm_concat.split("_") if p]
            if len(parts) >= 4:
                a_bldg = "_".join(parts[0:2])
                z_bldg = "_join".join([])  # placeholder to avoid accidental NameError

                z_bldg = "_".join(parts[2:4])
                bldg_rm_side = a_bldg if side == "A" else z_bldg

        # Fallback: parse BLDG_RM from src_stem (A-side naming was correct there)
        if not bldg_rm_side:
            m_loc = re.search(r"_C\d+_([A-Za-z0-9]+_[A-Za-z0-9]+)_ROW_", src_stem)
            if m_loc:
                bldg_rm_side = m_loc.group(1)
            else:
                bldg_rm_side = "BLDG_RM"

        # ---- Derive ROW/RACK/RU for the correct side ----
        # 04_30_48_13_C_10 -> A=(04,30,48) ; Z=(13,C,10)
        row = rack = ru = None
        if isinstance(row_rack_ru_concat, str) and row_rack_ru_concat.strip():
            parts = [p for p in row_rack_ru_concat.split("_") if p]
            if len(parts) >= 6:
                if side == "A":
                    row, rack, ru = parts[0], parts[1], parts[2]
                else:
                    row, rack, ru = parts[3], parts[4], parts[5]

        # Fallback: parse ROW/RACK/RU from src_stem
        if row is None or rack is None or ru is None:
            m_rrr = re.search(r"_ROW_([^_]+)_RACK_([^_]+)_RU_([^_]+)$", src_stem)
            if m_rrr:
                row, rack, ru = m_rrr.group(1), m_rrr.group(2), m_rrr.group(3)
            else:
                row, rack, ru = "NA", "NA", "NA"

        # Assemble exactly in your preferred style
        filename = f"{enclosure}_{side}{trunk_num}_{connector}_{bldg_rm_side}_ROW_{row}_RACK_{rack}_RU_{ru}.opm"
        return safe(filename).replace(" ", "_")

    # ----------------------------
    # Merge
    # ----------------------------

    def _worst_verdict(self, a: str | None, z: str | None) -> str:
        # conservative: Fail beats Pass; Unknown beats Pass
        order = {"Fail": 3, "Unknown": 2, "Pass": 1, None: 0}
        return a if order.get(a, 0) >= order.get(z, 0) else z

    def _merge_opm_docs(self, a_doc: dict, z_doc: dict) -> dict:
        """
        Merge A + Z OPM docs into a single "MergeMF" output.

        IMPORTANT:
        - We do NOT change original JSON data in the individual A/Z files.
        - This merged output combines the measurement lists:
            A stays as fibers 1..N
            Z is shifted to fibers (N+1)..(2N)
        - The "fiber number" in this schema is stored in each measurement object's "Name"
        (string digits like "1", "2", ...).
        """
        import copy

        merged = copy.deepcopy(a_doc)

        # Locate measurements list in the OPM structure we generate
        def _get_measurements(doc: dict) -> list:
            try:
                meas = doc["Measurement"]["OpmResultData"]["Measurements"]
                return meas if isinstance(meas, list) else []
            except Exception:
                return []

        a_meas = _get_measurements(a_doc)
        z_meas = _get_measurements(z_doc)

        # If either side is missing measurements, just return A copy (safe fallback)
        if not a_meas or not z_meas:
            return merged

        # Find the max fiber index on A by reading numeric "Name" fields
        a_max = 0
        for m in a_meas:
            name = m.get("Name")
            if isinstance(name, str) and name.isdigit():
                a_max = max(a_max, int(name))

        # If we couldn't determine A max, assume 12 (common MPO-12), but keep it safe.
        if a_max <= 0:
            a_max = 12

        # Shift Z measurement fiber numbers by +a_max
        z_shifted = []
        for m in z_meas:
            mm = copy.deepcopy(m)
            name = mm.get("Name")
            if isinstance(name, str) and name.isdigit():
                mm["Name"] = str(int(name) + a_max)
            z_shifted.append(mm)

        # Write merged measurement list
        try:
            merged["Measurement"]["OpmResultData"]["Measurements"] = list(a_meas) + z_shifted
        except Exception:
            # If structure isn't what we expect, just return A copy
            return merged

        # Best-effort: set verdict to worst-case of A/Z (if present)
        def _verdict(doc: dict) -> str | None:
            try:
                v = doc["Measurement"]["OpmResultData"].get("GlobalVerdict")
                return str(v) if v is not None else None
            except Exception:
                return None

        a_v = _verdict(a_doc)
        z_v = _verdict(z_doc)
        # If either is Fail, set Fail; else keep A's
        if (a_v == "Fail") or (z_v == "Fail"):
            try:
                merged["Measurement"]["OpmResultData"]["GlobalVerdict"] = "Fail"
            except Exception:
                pass

        return merged

    def _merge_eligible_pairs(self, eligible_pairs: list[tuple[str, "Path", "Path"]], output_dir: "Path") -> dict:
        """
        Merge eligible A/Z pairs to MergeMF files.

        eligible_pairs items are: (pair_key, a_path, z_path)
        NOTE: pair_key is used only for logging; filename is built from the paths.
        """
        merged_msgs: list[str] = []
        merge_write_errors: list[str] = []
        merged = 0
        write_errors = 0

        for pair_key, a_path, z_path in eligible_pairs:
            try:
                a_doc = load_json(a_path)
                z_doc = load_json(z_path)

                merged_doc = self._merge_opm_docs(a_doc, z_doc)

                out_name = self._build_merged_opm_filename(a_path, z_path)
                out_path = output_dir / out_name

                if out_path.exists():
                    raise FileExistsError(f"Output already exists: {out_path.name}")

                with out_path.open("w", encoding="utf-8") as f:
                    import json
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
    # Punch List (CSV) helpers
    # ----------------------------

    def _guess_ose_from_any(self, rows: list[dict]) -> str | None:
        for r in rows:
            v = r.get("OSE")
            if v:
                return str(v)
        return None

    def _write_punch_list_csv(self, out_path: Path, rows: list[dict]) -> None:
        headers = [
            "OSE",
            "Cable ID",
            "Location A",
            "Location B",
            "Test Date",
            "Tester",
            "Pair Key",
            "Issue Type",
            "Details",
        ]
        with out_path.open("w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=headers)
            w.writeheader()
            for r in rows:
                w.writerow({h: r.get(h, "") for h in headers})

    def _build_punch_row(
        self,
        pair_key: str,
        a_doc: dict | None,
        z_doc: dict | None,
        mismatch: bool,
        failure_high_loss: bool,
        a_fail: bool,
        z_fail: bool,
        polarity_issue: bool,
        wavelength_issue: bool,
        length_issue: bool,
        note: str | None = None,
    ) -> dict:
        # Use A doc as primary for metadata, fallback to Z
        doc = a_doc or z_doc or {}
        cable_id = self._get_job_cable_id(doc) or ""
        loc_a, loc_b = self._split_locations_from_cable_id(cable_id)

        test_date = self._get_test_datetime(doc) or ""
        tester = self._get_tester_string(doc) or ""

        # OSE = “option A for Short label” in your terms; we’ll derive from JobId if possible
        ose = self._get_ose_from_job_id(doc) or ""

        issue_parts = []
        if mismatch:
            issue_parts.append("MISMATCH")
        if failure_high_loss:
            issue_parts.append("FAILURE")
        issue_type = "+".join(issue_parts) if issue_parts else "ISSUE"

        details = []
        if note:
            details.append(note)
        if failure_high_loss:
            sides = []
            if a_fail:
                sides.append("A")
            if z_fail:
                sides.append("Z")
            details.append(f"High Loss (Status=Fail) side={('/'.join(sides) if sides else '?')}")
        if polarity_issue:
            details.append("Polarity issue")
        if wavelength_issue:
            details.append("Wavelength mismatch")
        if length_issue:
            details.append("Length missing/mismatch")

        return {
            "OSE": ose,
            "Cable ID": cable_id,
            "Location A": loc_a,
            "Location B": loc_b,
            "Test Date": test_date,
            "Tester": tester,
            "Pair Key": pair_key,
            "Issue Type": issue_type,
            "Details": "; ".join(details),
        }

    # ----------------------------
    # Data extraction helpers
    # ----------------------------

    def _fmt(self, v) -> str:
        if v is None:
            return "(missing)"
        if isinstance(v, float):
            # keep it readable without tons of noise
            return f"{v:.3f}".rstrip("0").rstrip(".")
        return str(v)

    def _get_job_cable_id(self, doc: dict) -> str | None:
        # Your note: Cable ID is populated from JobId
        job = doc.get("JobId")
        if not job:
            return None
        # Example: "PATH_1_LCO1-NS3-LCO2-DHB-00001.A03"
        # We want the middle “cable id chunk”
        m = re.search(r"PATH_\d+_(.+)", str(job))
        if not m:
            return str(job)
        return m.group(1)

    def _split_locations_from_cable_id(self, cable_id: str) -> tuple[str, str]:
        # Example: LCO1-NS3-LCO2-DHB-00001.A03
        # You said: LCO1-NS3 is Location A and LCO2-DHB is Location B.
        # So we take first two hyphen chunks for A and next two for B.
        if not cable_id:
            return "", ""
        base = cable_id.split(".")[0]
        parts = base.split("-")
        if len(parts) >= 4:
            loc_a = "-".join(parts[0:2])
            loc_b = "-".join(parts[2:4])
            return loc_a, loc_b
        return "", ""

    def _get_ose_from_job_id(self, doc: dict) -> str | None:
        job = doc.get("JobId")
        if not job:
            return None
        # Try to find something like "...A03" (path / segment)
        m = re.search(r"\.(A\d+|Z\d+)$", str(job))
        if m:
            return m.group(1)
        return None

    def _get_test_datetime(self, doc: dict) -> str | None:
        # We’ll try a few common spots
        for k in ("DateTime", "TestDateTime", "Timestamp", "CreatedAt", "TestDate"):
            v = doc.get(k)
            if v:
                return str(v)
        # Sometimes under OpticalData
        od = doc.get("OpticalData")
        if isinstance(od, dict):
            for k in ("DateTime", "TestDateTime", "Timestamp", "TestDate"):
                v = od.get(k)
                if v:
                    return str(v)
        return None

    def _get_tester_string(self, doc: dict) -> str | None:
        # Try a bunch of likely keys without being brittle
        candidates = []
        for k in ("TestSet", "TestSetName", "TestSetModel", "Instrument", "InstrumentName", "Tester", "Operator"):
            v = doc.get(k)
            if v:
                candidates.append(str(v))
        od = doc.get("OpticalData")
        if isinstance(od, dict):
            for k in ("TestSet", "Instrument", "InstrumentName", "Tester", "Operator"):
                v = od.get(k)
                if v:
                    candidates.append(str(v))
        if candidates:
            return " | ".join(dict.fromkeys(candidates))  # de-dupe while keeping order
        return None
    # ---- Polarity ----

    def _get_opm_root(self, doc: dict) -> dict:
        """
        Return the dict where OPM result fields usually live.

        Handles both schemas:
          - full OPM: doc["Measurement"]["OpmResultData"] (common for native test-set exports)
          - already-rooted: doc contains "Connectors"/"Measurements" directly (some generated outputs)
        """
        if not isinstance(doc, dict):
            return {}
        m = doc.get("Measurement")
        if isinstance(m, dict):
            od = m.get("OpmResultData")
            if isinstance(od, dict):
                return od
        return doc

    def _get_expected_polarity(self, doc: dict) -> str | None:
        root = self._get_opm_root(doc)
        con = root.get("Connectors")
        if not isinstance(con, dict):
            return None
        exp = con.get("ExpectedConnectors")
        if not isinstance(exp, dict):
            return None
        pol = exp.get("PolarityType")
        return self._normalize_polarity(pol) if pol else None

    def _get_actual_polarity(self, doc: dict) -> str | None:
        root = self._get_opm_root(doc)
        con = root.get("Connectors")
        if not isinstance(con, dict):
            return None
        act = con.get("ActualConnectors")
        if not isinstance(act, dict):
            return None
        pol = act.get("PolarityType")
        return self._normalize_polarity(pol) if pol else None

    def _get_polarity_status(self, doc: dict) -> str | None:
        """
        Return polarity status, if present.

        Primary:
          Measurement.OpmResultData.Connectors.PolarityStatus

        Fallback:
          Connectors.PolarityStatus
        """
        if not isinstance(doc, dict):
            return None

        root = self._get_opm_root(doc)
        con = root.get("Connectors")
        if isinstance(con, dict):
            ps = con.get("PolarityStatus")
            if ps is not None:
                return str(ps)

        return None

    def _get_wavelengths_nm(self, doc: dict) -> list[int]:
        """Return sorted unique list of wavelengths (nm) seen in Measurements."""
        root = self._get_opm_root(doc)

        out: set[int] = set()
        meas = root.get("Measurements")
        if not isinstance(meas, list):
            # fallback: some variants store measurements under OpticalData
            od = doc.get("OpticalData")
            if isinstance(od, dict):
                meas = od.get("Measurements")
        if not isinstance(meas, list):
            return []

        for m in meas:
            if not isinstance(m, dict):
                continue
            w = m.get("Wavelength")
            if isinstance(w, (int, float)):
                out.add(int(w))
        return sorted(out)

    def _normalize_polarity(self, pol) -> str:
        s = str(pol).strip()
        # unify separators: "MPO B" -> "MPO_B"
        s = s.replace(" ", "_")
        # collapse double underscores just in case
        while "__" in s:
            s = s.replace("__", "_")
        return s
    # ---- Length ----

    def _get_length_numeric_or_missing(self, doc: dict) -> tuple[float | None, bool]:
        """
        ONLY care about FiberLength.LengthInfo.Length numeric vs null.
        Ignore FiberLength.Status entirely.

        Primary OPM path:
        Measurement.OpmResultData.Measurements[i].FiberLength.LengthInfo.Length

        Fallbacks:
        Measurement.OpmResultData.FiberLength.LengthInfo.Length
        OpticalData.Measurements[...] (older/alternate)
        doc.FiberLength.LengthInfo.Length
        """
        if not isinstance(doc, dict):
            return None, True

        # -----------------------------
        # Primary: Measurement -> OpmResultData -> Measurements[]
        # -----------------------------
        m = doc.get("Measurement")
        if isinstance(m, dict):
            od = m.get("OpmResultData")
            if isinstance(od, dict):
                meas = od.get("Measurements")
                if isinstance(meas, list):
                    for meas_row in meas:
                        if not isinstance(meas_row, dict):
                            continue
                        fl = meas_row.get("FiberLength")
                        if not isinstance(fl, dict):
                            continue
                        li = fl.get("LengthInfo")
                        if isinstance(li, dict) and ("Length" in li):
                            val = li.get("Length")
                            if val is None:
                                return None, True
                            try:
                                return float(val), False
                            except Exception:
                                return None, True

                # Some variants store FiberLength at OpmResultData level
                fl = od.get("FiberLength")
                if isinstance(fl, dict):
                    li = fl.get("LengthInfo")
                    if isinstance(li, dict) and ("Length" in li):
                        val = li.get("Length")
                        if val is None:
                            return None, True
                        try:
                            return float(val), False
                        except Exception:
                            return None, True

        # -----------------------------
        # Fallback: OpticalData (older/alternate schema)
        # -----------------------------
        od2 = doc.get("OpticalData")
        if isinstance(od2, dict):
            meas = od2.get("Measurements")
            if isinstance(meas, list):
                for meas_row in meas:
                    if not isinstance(meas_row, dict):
                        continue
                    fl = meas_row.get("FiberLength")
                    if not isinstance(fl, dict):
                        continue
                    li = fl.get("LengthInfo")
                    if isinstance(li, dict) and ("Length" in li):
                        val = li.get("Length")
                        if val is None:
                            return None, True
                        try:
                            return float(val), False
                        except Exception:
                            return None, True

        # -----------------------------
        # Last fallback: doc-level FiberLength
        # -----------------------------
        fl = doc.get("FiberLength")
        if isinstance(fl, dict):
            li = fl.get("LengthInfo")
            if isinstance(li, dict) and ("Length" in li):
                val = li.get("Length")
                if val is None:
                    return None, True
                try:
                    return float(val), False
                except Exception:
                    return None, True

        return None, True

    # ---- High loss ----

    def _has_high_loss(self, doc: dict) -> bool:
        """
        True if the result contains any FAIL indication.
        This is NOT used as a merge blocker (per your rules) — it's for reporting only.
        """
        try:
            gv = doc.get("GlobalVerdict")
            if gv == "Fail":
                return True

            od = (doc.get("Measurement") or {}).get("OpmResultData") or {}
            if od.get("Status") == "Fail":
                return True

            measurements = od.get("Measurements", [])
            if not isinstance(measurements, list):
                return False

            for m in measurements:
                if not isinstance(m, dict):
                    continue

                # Sometimes individual measurement may have a Status/Verdict
                if m.get("Status") == "Fail" or m.get("Verdict") == "Fail":
                    return True

                readings = m.get("Readings", [])
                if not isinstance(readings, list):
                    continue

                for r in readings:
                    if isinstance(r, dict) and r.get("Status") == "Fail":
                        return True

            return False
        except Exception:
            # If something is malformed, don't crash analysis;
            # just assume not high loss here and let other logic catch issues.
            return False

if __name__ == "__main__":
    app = JSON2OPMApp()
    app.mainloop()
