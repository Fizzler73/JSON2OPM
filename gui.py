import sys
from pathlib import Path

def _ensure_runtime_path():
    # When running as an EXE (PyInstaller), sys._MEIPASS points to the temp extraction dir.
    # When running as a .py script, we want the project root (folder containing gui.py).
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    else:
        base = Path(__file__).resolve().parent

    if str(base) not in sys.path:
        sys.path.insert(0, str(base))

_ensure_runtime_path()

from json2opm.app_ui import JSON2OPMApp


def main():
    app = JSON2OPMApp()
    app.mainloop()


if __name__ == "__main__":
    main()
