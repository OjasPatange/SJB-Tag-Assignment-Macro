# Tests — Python replica harness

The validation method for this project is a **Python replica of the VBA logic**,
run against the sample dataset, used to error-check before any change lands in
VBA.

## Layout
- `fixtures/Sample_data.xlsx` — test dataset (3,772 tag rows, 70 SJBs).
- `replica_v14.py` — replica of the v14 assignment logic (paste in / port here).
- `replica_v15.py` — replica of proposed v15 logic once decisions are approved.
- `check_errors.py` — asserts E1–E9 conditions against replica output.

## Workflow
1. Port the current VBA logic into a `replica_vNN.py`.
2. Run it against the fixture.
3. Diff the placement output and assert the E1–E9 conditions are resolved.
4. Only then implement the change in VBA under `src/`.
