# Inputs — Input_sheet

Read with [shared guidance](shared-guidance.md). `Input_sheet` is a blank template for input sheets without a timeline. Tab/section colour: #FFD966. Freeze F6.

Rows 1:5 have black fill across the visible template width. All header contents are blank except C2, which displays `=model.check.message` in white bold text. There are no timeline formulas, unit labels or model flags.

Keep the shared A:E widths. Set every column from F through the final visible column to E's width, initially 11.56. For the default 48-month model, this is F:BV, with BW:XFD hidden. BV is full-width on Inputs, not a narrow spacer. For other horizons, use the monthly template's final visible column as the initial width boundary; later input modules can specify their own extent.

The body is empty and has no inherited row groups or collapsed spacer-column groups. Leave row 6 blank; the first input section may start at row 7. Create input ranges, validations, units and detail groups when a specific module is added, following shared formatting.

Copy Input_sheet for new input sheets. The base-workbook builder constructs this blank layout directly, without inheriting Monthly's calendar, flags or groups.
