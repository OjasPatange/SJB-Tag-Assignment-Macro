# SJB Tag Assignment Macro

VBA tooling for an industrial instrumentation project: assigns hardwired
instrument tags to cards and channels within SJBs (Smart Junction Boxes),
producing assignment reports for DCS, SIS, GDS, and PDS systems.

## Source of truth
The VBA modules under `src/` are the tracked source. The `.xlsm` workbook is a
binary **build artifact** (git-ignored) — you export from it and import back
into it, but git only diffs the text under `src/`.

```
sjb-tag-assignment/
├── src/            exported VBA modules (.bas / .cls) — tracked
├── tools/          export/import round-trip scripts
├── tests/          Python replica harness + fixtures
├── docs/           error log, changelog, corrective-action decisions
├── build/          .xlsm output (git-ignored)
├── requirements.txt
└── README.md
```

## Round-trip

Export VBA out of the workbook (any OS, uses `oletools`):
```bash
python tools/export_vba.py path/to/PCD_For_TA.xlsm --out src
```

Import edited modules back in (Windows + Excel, uses `pywin32`):
```bash
python tools/import_vba.py path/to/PCD_For_TA.xlsm --src src
```
> Import requires Excel's *"Trust access to the VBA project object model"*
> setting to be enabled.

## Workflow
1. Export current VBA → `src/`, commit as the baseline for a version.
2. Validate logic with a Python replica in `tests/` against the fixture.
3. Resolve the pending decisions in `docs/corrective_actions.md`.
4. Implement the change in `src/`, re-import, tag the version.

## Status
- **v14** — submitted, nine errors found (`docs/v14_error_log.md`).
- **v15** — pending five corrective-action decisions (`docs/corrective_actions.md`).

## Setup
```bash
python -m venv .venv && source .venv/bin/activate   # or .venv\Scripts\activate on Windows
pip install -r requirements.txt
```

## Files
- `PCD_For_TA.xlsm` — primary workbook (place in repo root or `build/`; not tracked).
- `tests/fixtures/Sample_data.xlsx` — test dataset.
