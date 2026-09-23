# Quarterly

Read with [shared guidance](shared-guidance.md). Quarterly is the template for quarterly reporting. Tab/section colour: #92D050; it does not require an FS sheet to obtain its colour. Freeze Z6.

Use `ROUNDUP(N/3,0)` columns from Z, linked to [Controls](controls.md). For 48 months, the timeline is Z:AO, spacer AP and hidden columns AQ:XFD. Quarter boundaries follow the financial year; a final partial quarter is identified by Controls.

## Header and flags

Apply the shared header/status convention. Fill these anchors right through the last quarterly column:

| Row | X label | Z anchor |
|---|---|---|
| 1 | Date | `=Controls!Z28` |
| 2 | Date | `=Controls!Z29` |
| 3 | Counter | `=Controls!Z30` |
| 4 | Label | `=Controls!Z31` |

Row 5 has no timeline contents. A7 is MODEL FLAGS; C9 is “Quarterly flags”. C10 is “Quarter-end monthly position”, X10 is #, and Z10 is `=Controls!Z59`, filled across. The value is a one-based position in the monthly range, not an Excel column number.

Group rows 8:10. For new templates, row 11 is blank and new reporting sections may start at row 12. Do not add empty flag placeholders. Existing templates may reserve rows 11:13; preserve those anchors during unrelated edits.

Copy Quarterly for quarterly reports. When reports are later added, specify aggregation per row: sum flows over the represented months, take period-end closing balances, and calculate ratios from the appropriate components. The blank template contains no financial outputs or aggregation formulas.
