# Controls

Read with [shared guidance](shared-guidance.md). Controls owns the calendar, model flags and shared lists. New-template tab/section colour: #FFD966; freeze Z6. Monthly columns begin at Z and end after N months; the next column is a spacer. Header formatting/status follows shared guidance.

## Configuration and layout

Use these anchors for new templates. Labels are in C, values in D, notes in E and units in X. Existing workbooks retain their mapped anchors.

| Cell | Control | Rule |
|---|---|---|
| D9 | Timeline start | First day of the financial year |
| D10 | Last actual date | Month-end between D9−1 and the final model date; D9−1 means no actual months |
| D11 | Forecast start | `=D10+1` |
| D12 | Horizon N | Whole months, at least 12; default 48 |
| D13 | Final model date | `=EOMONTH(D9,D12-1)` |
| D14 | Financial year-end month | Integer 1–12; consistent with D9 |
| D15 | Last actual month position H | Exact count of monthly end dates on/before D10; 0 through N |
| D16:D17 | Period labels | Distinct labels; defaults A and F |

D9, D10, D12, D14, D16 and D17 are inputs; the others are formulas. Validate configuration at build/edit time. E12 warns that a horizon change requires resizing the model: changing D12 alone does not create or remove columns.

| Block | Section/title | Content rows | Detail group |
|---|---|---|---|
| TIME | A7 | Controls 9:17; calendars below | 8:39 |
| Monthly calendar | C19 | 20:25 | Within TIME |
| Quarterly calendar | C27 | 28:31 | Within TIME |
| Annual calendar | C33 | 34:38 | Within TIME |
| MODEL FLAGS | A41 | Flags below | 42:69 |
| Monthly flags | C43 | 44:56 | Within MODEL FLAGS |
| Quarterly flags | C58 | 59 | Within MODEL FLAGS |
| Annual flags | C64 | 65 | Within MODEL FLAGS |
| LISTS | A71 | Currency and calendar months | 72:87 initially |

Calendar/flag labels go in C, types in X, formulas from Z across the relevant period count. Section labels are not repeated on every row. Omit unused placeholder labels and formulas.

## Calendar semantics

For month position k = 1…N:

| Row | Meaning |
|---|---|
| 20 | Start: D9 for the first month, then previous month-end + 1 |
| 21 | End: EOMONTH of that month's start |
| 22 | Type: D16 when k ≤ H; otherwise D17 |
| 23 | Monthly position k |
| 24 | Financial year, labelled by its ending calendar year |
| 25 | Fiscal quarter Q1…Q4 followed by the financial-year label |

For clarity, the first financial-year label is `YEAR(EOMONTH(D9,11))`; increment it after each 12 months. Fiscal quarter within a year is `1+INT(MOD(k-1,12)/3)`. Use exact date/month logic; never estimate month positions using days divided by 30.

Quarterly rows 28:31 hold start, end, sequential quarter position, and fiscal-quarter label. There are `ROUNDUP(N/3,0)` periods. Quarter q covers monthly positions `3q−2` through `MIN(3q,N)`; take dates from those monthly columns. Add “(partial)” to the label if fewer than three months are represented.

Annual rows 34:38 hold start, end, forecast-month count, period type and sequential year position. There are `ROUNDUP(N/12,0)` periods. Year y covers positions `12y−11` through `MIN(12y,N)`. Count actual and forecast months within that span. The type is A or F when uniform, otherwise e.g. A5+F7; use the configured labels. Append “(partial)” for fewer than 12 months. No annual dates extend beyond D13.

All monthly dates must be contiguous. Quarter/year flags below identify the final represented month, including a partial final period. Quarter/year end **positions** are one-based offsets into the monthly range, not worksheet column numbers.

## Monthly flags

H is the last actual position; k is the current month position. Flags are numeric.

| Row | Flag | Definition |
|---|---|---|
| 44 | Month | Calendar month number |
| 45 | Days | Days in that month |
| 46 | Actuals | 1 when k ≤ H, else 0 |
| 47 | Forecast | 1 when k > H, else 0 |
| 48 | First forecast | 1 only when k = H+1 and H < N |
| 49 | Forecast counter | 0 for actuals; otherwise k−H |
| 50 | Last actual | 1 only when k = H and H > 0 |
| 51 | Year counter | `1+INT((k-1)/12)` |
| 52 | Financial year end | 1 when the calendar month equals D14, else 0 |
| 53 | Quarter counter | `1+INT((k-1)/3)` |
| 54 | Forecast year counter | 0 for actuals; otherwise current year counter minus the year counter of the first forecast month + 1 |
| 55 | Quarter start flag | 1 when `MOD(k-1,3)=0`, else 0 |
| 56 | Quarter end flag | 1 when `MOD(k,3)=0`, else 0 |

Do not make first/last flags depend on blank cells outside the timeline. The forecast year counter advances at financial-year boundaries, not after every 12 forecast months. Quarter start/end flags mark actual fiscal-quarter boundaries; a partial final quarter does not create an artificial quarter-end flag. A partial final year does not create an artificial financial-year-end flag.

Row 59 contains each quarter's final represented monthly position; row 65 contains each year's. Use these to align downstream reporting to the monthly model.

## Shared lists and names

C73 is “Model currency”; D73 is the configured currency input, named `model.currency` at workbook scope. This is a model unit label, not an FX-conversion mechanism.

C75:D75 contains “Calendar month” and “Month number”; C76:D87 lists January–December and 1–12. Initially include no employee, vendor, forecast-driver or account lists. Add a list when its consuming module is built, updating relevant names, validation sources and grouping. Preserve existing codes when editing a model with consumers.

## Horizon changes

At N=48, monthly ranges end BU, quarterly AO and annual AC. Derive these boundaries for other horizons. Resize affected calendar/flag blocks and all dependent templates/modules, including formulas, names, groups, formatting and hidden-column boundaries. Shortening a populated model requires assessing data that would be removed; a request to edit a date does not authorize discarding it.
