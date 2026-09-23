# Monthly

Read with [shared guidance](shared-guidance.md). Monthly is the template for sheets using the monthly calendar. Tab/section colour: #00B0F0. Freeze Z6. Start the timeline in Z and use N columns from [Controls](controls.md); for 48 months these are Z:BU, followed by spacer BV and hidden BW:XFD.

## Header and flags

Apply the shared header/status convention. The following formulas link to Controls and fill right through the last monthly column:

| Row | X label | Z anchor |
|---|---|---|
| 1 | Date | `=Controls!Z20` |
| 2 | Date | `=Controls!Z21` |
| 3 | Label | `=Controls!Z22` |
| 4 | Counter | `=Controls!Z23` |
| 5 | Year | `=Controls!Z24` |

A7 is MODEL FLAGS; C9 is “Monthly flags”. Rows 10:22 link, in order, to Controls rows 44:56. Put matching labels in C and units in X: 1/0 for binary flags, # for other numeric flags. For example, Z10 is `=Controls!Z44`, filled across; subsequent rows follow the source row sequence. Rows 21 and 22 contain Quarter start flag and Quarter end flag, linking to Controls rows 55 and 56. Do not recalculate flags independently.

Group rows 8:22. Row 23 is blank; new module sections start at row 24 or later. The base template has no business calculation sections.

Copy Monthly for a new monthly sheet. Keep its timeline columns aligned with all other monthly sheets. If Controls anchors differ in an existing workbook, use its mapped source ranges rather than these new-template addresses.
