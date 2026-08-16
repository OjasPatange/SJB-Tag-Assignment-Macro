#!/usr/bin/env python3
"""
export_vba.py — extract VBA modules from an .xlsm into src/ as text.

This is the source-of-truth export: the .xlsm is a binary build artifact,
the exported modules under src/ are what git actually diffs.

Usage:
    python tools/export_vba.py path/to/PCD_For_TA.xlsm [--out src]

Extensions follow VBA convention so re-import is unambiguous:
    .bas  standard module
    .cls  class module / document module (ThisWorkbook, Sheet*)
    .frm  userform code (form binary itself is not round-tripped here)

Requires: oletools  (pip install oletools)
"""
import argparse
import sys
from pathlib import Path

try:
    from oletools.olevba import VBA_Parser
except ImportError:
    sys.exit("oletools not installed. Run: pip install oletools")


def ext_for(code: str, name: str) -> str:
    head = code.lstrip().upper()
    if head.startswith("VERSION") and "BEGIN" in head[:400] and "MULTIUSE" in head[:400]:
        return ".cls"
    if name in {"THISWORKBOOK"} or name.upper().startswith("SHEET"):
        return ".cls"
    return ".bas"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsm", help="path to the .xlsm workbook")
    ap.add_argument("--out", default="src", help="output dir (default: src)")
    args = ap.parse_args()

    xlsm = Path(args.xlsm)
    if not xlsm.exists():
        sys.exit(f"file not found: {xlsm}")

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    parser = VBA_Parser(str(xlsm))
    if not parser.detect_vba_macros():
        sys.exit("no VBA macros found in workbook")

    count = 0
    for _, _, vba_filename, vba_code in parser.extract_macros():
        name = Path(vba_filename).stem
        if not name:
            continue
        target = out / f"{name}{ext_for(vba_code, name)}"
        target.write_text(vba_code, encoding="utf-8", newline="\n")
        print(f"  exported  {target}")
        count += 1

    parser.close()
    print(f"\n{count} module(s) written to {out}/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
