# JSON2OPM

JSON2OPM is a desktop utility that converts EXFO Exchange JSON test results into OPM (.opm) files, analyzes A/Z fiber pairs for consistency, and optionally merges clean A/Z pairs into a single Multi-Fiber OPM result.

This tool is designed for fiber test engineers working with high-density MPO/MTP links who need deterministic, auditable results.

---

## What it does

- Converts Exchange-exported JSON files into OPM format
- Automatically pairs A-side and Z-side results
- Detects A/Z mismatches in:
  - 🔀 Polarity
  - λ Wavelength
  - 📏 Fiber length (missing or delta over threshold)
- Produces a clear, color-coded log:
  - ✅ Converted / Merged results (green)
  - ❌ A/Z mismatches (red)
  - 📊 One concise summary block (blue)
- Optionally merges **eligible** A/Z pairs into `*_MergeMF.opm`
- Exports mismatch details to CSV for reporting or escalation

---

## What “eligible to merge” means

An A/Z pair is eligible for merge **only if**:
- No polarity mismatch
- No wavelength mismatch
- No missing length
- Length delta is within the configured threshold

If any of the above fail, the pair is **not merged** and is reported as an error.

---

## How to run

From the project root:

```powershell
python gui.py
