#!/usr/bin/env python3
"""
import_vba.py — re-import text modules from src/ back into an .xlsm.

Round-trip counterpart to export_vba.py. Because oletools cannot WRITE VBA,
re-import is done through Excel COM automation, so this script runs on
Windows with Excel installed (your primary dev environment).

It removes existing non-document modules of the same name, then imports the
files from src/. Document modules (ThisWorkbook, Sheet*) cannot be replaced by
import — their code is pasted in instead.

Prerequisites (one-time, in Excel):
    File > Options > Trust Center > Trust Center Settings
      > Macro Settings > "Trust access to the VBA project object model" = ON

Requires: pywin32  (pip install pywin32)

Usage:
    python tools/import_vba.py path/to/PCD_For_TA.xlsm [--src src]
"""
import argparse
import sys
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:
    sys.exit("pywin32 not installed (Windows + Excel required). Run: pip install pywin32")

DOC_MODULE = 100  # vbext_ct_Document


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsm")
    ap.add_argument("--src", default="src")
    args = ap.parse_args()

    xlsm = Path(args.xlsm).resolve()
    src = Path(args.src)
    if not xlsm.exists():
        sys.exit(f"file not found: {xlsm}")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(str(xlsm))
    try:
        proj = wb.VBProject
        existing = {c.Name: c for c in proj.VBComponents}

        for f in sorted(src.glob("*")):
            if f.suffix.lower() not in {".bas", ".cls", ".frm"}:
                continue
            name = f.stem
            comp = existing.get(name)
            if comp is not None and comp.Type == DOC_MODULE:
                # document module: overwrite code lines, don't remove component
                mod = comp.CodeModule
                if mod.CountOfLines:
                    mod.DeleteLines(1, mod.CountOfLines)
                mod.AddFromString(f.read_text(encoding="utf-8"))
                print(f"  code-set  {name}")
            else:
                if comp is not None:
                    proj.VBComponents.Remove(comp)
                proj.VBComponents.Import(str(f.resolve()))
                print(f"  imported  {name}")

        wb.Save()
        print(f"\nsaved {xlsm}")
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
