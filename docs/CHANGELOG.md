# Changelog

## v15 — in progress
Corrective actions for E1–E9. Five decisions posed with recommended defaults;
awaiting approval before implementation. See `corrective_actions.md`.

## v14
Independently developed and submitted for error checking. Nine errors found
(E1–E9) — see `v14_error_log.md`. Code committed as
`src/SJB_Tag_Assignment.bas`; a worked input/output example is in `examples/`.

Only behavioural change over v13: **system is derived from the SJB name**
(`SystemFromSJB`) instead of Column I — `PDSSB`→PDS, `SSB`→SIS, `DSB`→DCS,
else DCS. The Column-I segregation rule and `SegregationReport` are removed.

## v11
`IsTrainToken()` introduced: any `_<Letter≠G><digits>` suffix qualifies as a
train/family token; full uppercased token kept as card-lock identity.
Unplaced tags reduced 103 → 86.

## v10
Core defect: `ParseGrouping` detected trains only via the `_T` substring and
discarded the family prefix, so families like `VRC_C1` / `VRC_C2` were treated
as independent tags.
