# Excel model instructions

Reusable instructions for Codex, including models of 30 MB or more. Read this document and the relevant sheet specification: [Controls](controls.md), [Monthly](monthly.md), [Quarterly](quarterly.md), [Annual](annual.md), [Inputs](inputs.md), [Checks](checks.md). Load other sheet specifications only when dependencies require them. Reuse instructions already read in the task; reread when they change or relevant context is missing.

## Scope and authority

Keep specifications in `docs/`, generated example workbooks in `templates/`, production scripts in `runner/`, and development tests in `tests/`. Use `.build/` only for local results and retired scripts. See [runner usage](runner.md) for new-workbook creation and the development/release boundary. Resolve each output map's workbook path relative to that map; its specification names refer to `docs/` in the recorded release package.

These specifications define new templates. When editing an existing model, preserve its layout and logic unless a change is requested; use its workbook-specific map to locate the affected area. Do not silently migrate it to the new template layout. User instructions override these defaults. Actual workbook state establishes what currently exists; the specifications establish the intended result.

The base workbook contains Controls, Monthly, Quarterly, Annual, Input_sheet and Checks. Controls owns all three calendars; the monthly, quarterly and annual templates link to it. The workbook contains flags and shared lists, but no business assumptions beyond its configuration and no specific model checks.

For a new workbook, establish the start date, actuals cutoff, financial year end and currency from the task; do not inherit PCI dates or USD silently. Default to 48 months. The model starts on the first day of its financial year and covers at least one full financial year. All periods are whole months. If the horizon ends mid-quarter/year, include a final partial reporting period and identify it as partial.

For a new model using this six-sheet standard, invoke the [local runner](runner.md). It constructs an empty Excel workbook directly from the tested implementation of these specifications, with the requested configuration and timeline lengths. Existing workbooks and their maps are outputs, not build inputs. Codex updates the specifications and builder together when the standard changes; routine runs execute the builder without reinterpreting the prose. Add and verify explicit builder support for other layouts before distributing them.

For new sheets within a model, copy the appropriate sheet template; preserve its calendar, flags, widths, panes and groups unless the new module specifies an exception. Template changes do not automatically update existing copies.

## Efficient editing and verification

Speed and correctness are joint requirements. Use native Excel automation directly for local workbook creation, editing and calculation when available. This project-specific tool choice takes precedence over generic spreadsheet-skill defaults. Use another library or engine only when Excel is unavailable or a required capability needs it; preserve workbook features and keep the additional work bounded.

Reuse a suitable existing build/edit script that follows this tool policy, adapting its target and parameters to the request. Do not rerun a full builder for a narrow edit or let saved script defaults reset unrelated workbook state. Write only the additional automation the task needs.

Use the explicit target path throughout a task. A new-model request permits a new file copied from the approved template; when editing an existing model, create a new version only when requested. Do not overwrite a newer workbook from an older base to recover a failed edit.

Use a compact map to locate the affected ranges, then inspect enough current state to establish that the intended edit is safe. A mapped value change needs little inspection; a row insertion or horizon change requires dependent ranges and workbook objects. Avoid exporting all sheets or reading every formula to answer a narrow question.

Batch edits, calculation and verification in one task-owned Excel session where practical. Read/write bounded blocks using rectangular arrays and `Value2`, rather than one COM call per cell. Apply styles to actual content blocks, avoiding whole-sheet formatting. Temporarily suppress calculation, events and screen updates where appropriate, restoring their original settings on exit. Avoid unnecessary activation, clipboard work and inflated used ranges. These practices follow [Microsoft's performance guidance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions).

Use explicit worksheet objects; resolve a copied worksheet before renaming it. Preserve macros, names, validation and other features present in the target. Use `Formula2`/`Formula2R1C1` where supported and appropriate; preserve existing implicit-intersection or array behaviour deliberately. See [Microsoft's Formula2 documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2).

Verify proportionately by default:

| Change | Verification |
|---|---|
| New design or broad redesign | Required sheet structure, calculations and visual layout on every sheet |
| Values or formatting | Changed range and relevant output/display |
| Repeated formulas | Anchor, representative copied formulas, boundaries and affected outputs |
| Rows, columns, sheets or horizon | Changed block, affected references, names, validations, groups, panes and range boundaries |
| Calendar/flag logic | Period alignment and actual/forecast transitions, including applicable edge cases |

For later edits, visually review changed areas and affected views; expand review when the change or a failure warrants it. Reusing an approved template does not require reviewing unchanged layouts again. Reimporting the workbook into another engine requires a specific capability or verification need; avoid routine cross-engine rendering and recalculation. These are build-time verification activities, not instructions to populate the Checks sheet.

Recalculate after a batch of formula/value changes using a scope that includes affected dependencies. Formatting-only work normally needs no recalculation. Calendar changes may affect the entire model. Do not force a full calculation or dependency rebuild for every edit, and do not treat a local recalculation as proof that downstream outputs are current. See [Microsoft's calculation guidance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance).

Save at completion, with intermediate saves only when warranted. Avoid routine repeated save/close/reopen cycles. Check execution and save results; if an edit fails or may be partially applied, inspect that state before retrying. Close/release only resources owned by the task. Never discard the user's unsaved work. Report verification limits, including unavailable recalculation; script success alone does not establish correctness.

## Modelling conventions

Use short, auditable formulas, bounded references and explicit assumption cells. Avoid unnecessary volatile formulas, full-column calculations and repeated lookups. Calculate shared drivers once per block. Use only the detail needed for the model's purpose.

Repeated formulas must copy correctly across periods and down equivalent detail rows. Expose complex calculations in understandable steps. Total additive detail with a simple SUM; do not sum ratios, flags or balances across periods without the relevant business rule. Monetary unit labels reference `=model.currency`.

## Shared formatting

Font: Verdana 10. Standard body row height: 18 points; increase only when content needs additional lines. Labels left aligned, numeric outputs right aligned. No merged cells. Rows 1:5 have black fill and white text. Every sheet has `C2 =model.check.message`, white and bold; C1 and C3:C5 remain blank. This is the sole content exception to the blank Inputs header.

Sections have bold uppercase titles in A. Subsections have uppercase titles in B, leaving A blank/unfilled. Bands match the tab colour and extend through the last content/timeline column, excluding the right spacer. Leave one blank row before each title. Group detail under each subsection, or under a section with no subsections; include its total. Collapsed outlines leave one visible blank row between titles. Do not group section-level rows above the first subsection.

Calculation detail has no borders; totals are bold with a normal solid top border across the calculation width. Update grouping when extending a block. Calendar and flag blocks do not receive arbitrary total rows.

| Cell type | Appearance |
|---|---|
| Manual inputs and list values | Blue #1950FF font, yellow #FFFF66 fill |
| Dropdowns | Blue #1950FF font, pink #FCD8E0 fill |
| Formulas and ordinary labels | Black font; header white overrides this |
| Warnings and visible assumption notes | Red #C00000 font |
| Irrelevant input/calculation cells | No solid background; grey #808080 light downward diagonal pattern |

Input/list borders are hairline, continuous, automatic colour on inside and outside edges. List titles are bold black; separate lists with one blank row. Do not apply irrelevant-cell patterns to period headings.

Time-based sheets use X for units and Z as the first timeline column. Write literal unit labels such as `1/0` as text to prevent Excel converting them to dates; monetary unit labels remain formulas referencing `model.currency`. Default widths:

| Columns | Width |
|---|---:|
| A:B | 0.94 |
| C | 29.11 |
| D:E | 11.56 |
| F:W | 0.94 |
| X | 9.67 |
| Y | 0.94 |
| Timeline | 11.56 |
| Following spacer | 1.00 |

Hide columns after the spacer. Group F:W on time-based sheets; keep content-bearing columns visible. Widen columns used for new content with bounded, reviewed widths (normally at least 12). Do not broadly autofit a large workbook. Format only occupied template rows/blocks; column visibility and widths may cover the required column spans.

Use these canonical number formats rather than repeating them in maps:

| Type | Excel number format |
|---|---|
| Date | `[$-en-US]mmm/d/yy;[$-en-US]mmm/d/yy;"-";@` |
| Integer / whole monetary amount / flag / counter | `#,##0;[Red]-#,##0;"-"` |
| Decimal | `#,##0.0;[Red]-#,##0.0;"-"` |
| Percentage | `0.0%;[Red]-0.0%;"-"` |
| Financial-year label | `0` |

Zero flags and counters display as dashes; their stored values remain numeric 0/1 or counts. Financial-year labels remain ungrouped (2025). Retain an existing format during unrelated edits.

## Compact workbook map

Each runner-created workbook receives its own `<name>.map.json` beside the output, referencing the matching release specifications. The builder generates this map; it does not read an existing workbook map. The legacy `docs/workbook-map.json`, if retained, describes the earlier v1.1 example only. For other workbook work, maintain one map beside its documentation. It is a locator and record of the actual workbook, not another copy of these specifications. Record the workbook actually built or inspected; never assume an existing workbook matches these generic coordinates.

Record only:

- Workbook path/version and the scope/date of the most recent structural verification.
- Actual sheet names, specification references, key anchors, timeline extents, section/block boundaries and sheet-specific exceptions.
- Key source/target dependencies, defined names, validation sources and nonstandard groups/panes/hidden ranges.
- Anchor formulas and fill extents where they are needed to maintain workbook-specific logic.

Standard styling is inherited by reference. Do not store used-range noise, repeated formulas, generic business rationale, or historical change narratives. Update only affected map entries after changes. Do not label the whole map verified when only one area was checked. Validate JSON syntax, without imposing full workbook reconciliation on each edit.
