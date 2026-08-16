# v14 Error Log

Nine errors identified in v14 via a Python replica of the assignment logic
run against the sample dataset (3,772 tag rows, 70 SJBs). Each is a candidate
fix for v15.

| ID | Summary | Root cause |
|----|---------|-----------|
| **E1** | Non-train groups split across cards | `GB\|<index>` key discards the loop base, so a single group fragments. |
| **E2** | Loop-to-channel mapping broken in multi-index loops | Fill-order assignment in pooled cards breaks the loop→channel correspondence. |
| **E3** | Tags from unrelated plant areas share cards | Plant-wide index pooling ignores area boundaries. |
| **E4** | Train groups severely fragmented in 16-ch SIS cards | Per-index `GT\|train\|idx` keys leave tags unplaced despite free channels. |
| **E5 / E6** | SIS box overflow | IS/NIS exec conflicts within shared group indices. |
| **E7** | GDS unreachable; unknown SJB tokens fall through | GDS is dead code; unrecognized name tokens silently route to the DCS hardware model. |
| **E8** | Column I read but unused | Removes the only cross-check between source data and derived system. |
| **E9** | SPARE channels share boxes with unplaced tags | No linking marker connecting a SPARE to the tag that should occupy it. |

## Cross-cutting cause
Introducing pure-index grouping keys (v13/v14) discards contextual
information — loop base, plant area — which cascades into fragmentation (E1,
E4) and mixing (E2, E3) errors.

## Validation method
Python replicas of the VBA logic, run against `tests/fixtures/Sample_data.xlsx`,
are the standard error-checking method before any change is implemented in VBA.
