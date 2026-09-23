# Checks

Read with [shared guidance](shared-guidance.md). Checks provides the future home for model checks and the shared status message. **The base template contains no specific checks**, including calendar checks, balance-sheet checks or thresholds.

Checks uses [Monthly](monthly.md)'s dates, first eleven flags (rows 10:20), widths, hidden columns and freeze panes. Its flag group is 8:20; do not create Monthly's quarter-boundary flags in rows 21:22, where the CHECKS section follows. Tab and section colour: #F1A983. The inherited calendar is template infrastructure, not a configured check. Existing Checks sheets are not automatically expanded when Monthly gains new flags.

## Blank-template layout

| Cell/range | Content |
|---|---|
| A22 | CHECKS section band |
| Rows 23:26 | Blank space for future check configuration |
| C27 | “No checks configured” |
| Rows 28:34 | Blank; no check labels, source links or result formulas |
| C35 | “Failed check types” |
| E35 | Blank while no checks are configured |
| C37 | “Check messages”, bold |
| C38 | Editable text: “Checks not configured” |
| C39 | Blank, available for later messages |
| C40 | `=C38` |

Group rows 23:40. C38 uses the shared manual-input style. Create workbook-scoped `model.check.message` referring to `=Checks!$C$40`. Every model sheet, including Checks, displays it in C2. The source message must never depend on those header cells; that would create a circular reference.

Do not insert links to absent reporting sheets, a default financial tolerance, or formulas treating empty result cells as successful checks. No prebuilt check inventory or fixed number of future check rows is required; the blank area above preserves familiar status anchors.

## When checks are later requested

Add only the requested checks and their required inputs. Reuse available source results and bounded ranges. Record each check's applicability and pass/fail rule. Missing, invalid or unevaluated required results are unavailable, not passed; don't convert formula errors to success.

Expand the result/summary block and update the map, grouping and message name as needed. E35 then counts **check types with at least one failure**, not failed monthly cells. Repeating a quarterly result across monthly columns must not inflate a reported period count.

The message must distinguish no configured checks, failures, unavailable results and all configured checks passed. Report failures and unavailable checks together when both exist. Use wording limited to configured checks; do not promise that the entire model is error-free.

These are extension rules, not instructions to populate or test a check library in the blank template.
