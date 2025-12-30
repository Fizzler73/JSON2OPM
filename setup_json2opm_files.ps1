# ============================================================
# JSON2OPM project bootstrap (folders + python files)
# ============================================================

$Base = "E:\OneDrive\Projects\PythonProjects\JSON2OPM"

# -------------------------
# Folders
# -------------------------
$Folders = @(
    "$Base\input_json",
    "$Base\output_opm",
    "$Base\reference",
    "$Base\json2opm"
)

# -------------------------
# Files with exact locations
# -------------------------
$Files = @(
    "$Base\main.py",
    "$Base\gui.py",

    "$Base\json2opm\__init__.py",
    "$Base\json2opm\loader.py",
    "$Base\json2opm\mapper.py",
    "$Base\json2opm\writer.py",
    "$Base\json2opm\diff.py",
    "$Base\json2opm\app_ui.py",
    "$Base\json2opm\dnd.py"
)

Write-Host "`n--- Creating folders ---"
foreach ($folder in $Folders) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder | Out-Null
        Write-Host "Created folder: $folder"
    } else {
        Write-Host "Exists: $folder"
    }
}

Write-Host "`n--- Creating files (no overwrite) ---"
foreach ($file in $Files) {
    if (-not (Test-Path $file)) {
        New-Item -ItemType File -Path $file | Out-Null
        Write-Host "Created file: $file"
    } else {
        Write-Host "Exists: $file"
    }
}

Write-Host "`nJSON2OPM structure verified."
Write-Host "You can now safely run:  python gui.py"
