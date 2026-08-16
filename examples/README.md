# Worked example

`SJB_Tag_Assignment_Generator.xlsx` is a full input-and-output snapshot showing
what the v14 macro (`src/SJB_Tag_Assignment.bas`) produces. It's kept here as a
reference for how the reports are generated — the code is the source of truth,
this workbook shows the result.

## Sheets

**Inputs**
- `Tag Details` — 3,772 hardwired instrument tags (the source data the macro
  reads; the macro treats this sheet as read-only).
- `SJBConfig` — 70 SJBs with IO capacity, used IOs, an auto-size suggestion,
  and the derived system. Built by `BuildSJBConfig`.

**Generated reports** (produced by `AssignTags`)
- `AssignedReport` — per-SJB summary: capacity, cards, used IOs, spare count,
  % spare, status, plus a TOTAL row.
- `DCS Tag Assignment Report` — every DCS tag with its assigned Card → Channel,
  SPARE channels included, sorted by SJB → Card → Channel.
- `SIS Tag Assignment Report` — same, for SIS.
- `PDS Tag Assignment Report` — same, for PDS.
- `UnplacedReport` — tags that couldn't be placed, with the reason.

> Note: there is **no GDS report sheet**. In v14 the system is derived from the
> SJB name and no GDS name-token was supplied, so GDS is not produced — this
> matches error **E7** in `docs/v14_error_log.md`.

## How it was produced
1. Open the macro-enabled workbook with `src/SJB_Tag_Assignment.bas` imported.
2. Run `BuildSJBConfig` to (re)build the `SJBConfig` sheet.
3. Run `AssignTags` to generate the report sheets above.

## Note on data
This workbook contains real project tag data. Keep the repository **private**.
