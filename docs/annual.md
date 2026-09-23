# Annual

Read with [shared guidance](shared-guidance.md). Annual is the template for annual reporting. Use [Quarterly](quarterly.md)'s formatting, including green tab/section colour #92D050, widths, header/status styling, grouping and freeze panes at Z6.

Use `ROUNDUP(N/12,0)` columns from Z, linked to [Controls](controls.md). For a 60-month workbook starting January 2025, the five years 2025–2029 occupy Z:AD, followed by spacer AE and hidden AF:XFD. The 48-month default uses Z:AC, spacer AD and hidden AE:XFD. Leave columns beyond the annual timeline empty; keep the spacer unfilled and width 1.00.

## Header and flags

Apply the shared header/status convention. Fill these anchors right through the last annual column:

| Row | X label | Z anchor |
|---|---|---|
| 1 | Date | `=Controls!Z34` |
| 2 | Date | `=Controls!Z35` |
| 3 | #months | `=Controls!Z36` |
| 4 | Label | `=Controls!Z37` |
| 5 | Counter | `=Controls!Z38` |

Rows 1:2 are dates; rows 3 and 5 use the standard integer format. Row 4 identifies actual, forecast or mixed periods using Controls' configured labels. With actuals through August 2026, that year's forecast-month count is 4 and its period type is A8+F4. Any final partial year is identified by Controls.

A7 is MODEL FLAGS; C9 is “Annual flags”. C10 is “Year-end monthly position”, X10 is #, and Z10 is `=Controls!Z65`, filled across. Values are one-based positions in the monthly range, not Excel column numbers. For a 60-month horizon they are 12, 24, 36, 48 and 60. Use the standard integer format.

Group rows 8:10. Row 11 is blank; new reporting sections may start at row 12. Copy Annual for new annual reports. The base template contains no financial outputs or aggregation formulas; when reporting rows are added, sum flows over represented months, take year-end closing balances, and calculate ratios from the appropriate components.
