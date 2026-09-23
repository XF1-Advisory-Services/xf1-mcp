# Documentation review — 22 September 2026

The revised instruction set starts at [shared-guidance.md](shared-guidance.md) and includes [Controls](controls.md), [Monthly](monthly.md), [Quarterly](quarterly.md), [Annual](annual.md), [Inputs](inputs.md) and [Checks](checks.md). This review note explains the changes; Codex does not need to load it for routine model work.

## Agreed requirements

- Audience: primarily Codex. Originals remain untouched; publish here.
- Efficient operation on models of 30 MB or more.
- Default horizon 48 months, configurable; start at financial-year start and cover at least one complete financial year.
- Speed and correctness both matter. The current tool-selection, reuse and verification policy is maintained in [shared guidance](shared-guidance.md#efficient-editing-and-verification).
- Reusable Markdown specifications plus a compact workbook-specific map, without duplicated rule sets.
- Controls initially contains currency and calendar months only; other lists arrive with their modules.
- Checks is included but contains no specific checks. All model-sheet headers display its status, including Inputs. Annual was added to the scope on 23 September 2026.

## What changed

| Original approach | Revised approach |
|---|---|
| Four shared guides plus five prose/map pairs | One shared guide and six sheet specifications, including the subsequently requested Annual sheet |
| Layout, styles and copy instructions repeated across prose and JSON | Shared rules once; sheet exceptions locally; actual workbook anchors in its map |
| PCI dates, currency, account lists, driver codes and version-selection convention | Model configuration and module-specific lists supplied for the actual task |
| Map-first policy often prohibiting inspection and verification | Map locates the edit; bounded inspection and proportionate verification establish correctness |
| Mandatory one-off COM procedures and a fixed failure-count stop rule | Direct native Excel automation; reuse suitable scripts and recover based on actual workbook state |
| Long PowerShell example and repeated implementation cautions | Short conditional automation guidance, with Microsoft references |
| Approximate days/30 actual-month count and ambiguous “column number” | Exact monthly position, explicitly one-based within the timeline |
| Year counter mixes fiscal-year labels and calendar start year | Position-based fiscal-year counter, starting at one |
| Boundary flags depend on cells outside the timeline | Defined no-actuals and all-actuals behaviour |
| Fixed 48-column examples can be mistaken for limits | Period counts and all dependent boundaries derive from N |
| Empty checks can produce an OK message | “Checks not configured”; no check formulas added |
| Check-message requirement conflicts with blank Inputs header | C2 is the explicit exception |

Simplification rests on removing duplication, project history and unnecessary prescribed mechanics. It does not assume that a newer model eliminates Excel feature-preservation, calculation or verification requirements. No comparative GPT-version benchmark was performed.

## New-template design choices

The documents make whole-month cutoffs explicit and allow partial final quarters/years. A model with no actual months uses the month-end before its start as the cutoff. Existing calendars are not silently converted.

New Controls uses D73 for currency and C76:D87 for calendar months, shortening the legacy list area. New Quarterly omits unused flag placeholder rows. Checks retains familiar C40 status anchoring and blank room for later checks, without the PCI balance-sheet tolerance or reporting links. These are new-template defaults; existing workbook coordinates are preserved unless migration is requested.

Minor number-format variants are consolidated; flags and counters use the standard integer format, including dashes for zero. Section bands end at the last data column; the Inputs header spans its visible width. The Annual sheet, subsequently added at the user's request, links to Controls and follows Quarterly's formatting.

## Sources and validation

Reviewed these 14 files in `C:/Users/VavrinecKryzanek/OneDrive - X-F1 Advisory Services Private Limited/XF1/202606 PCI Excel/docs`:

- `controls_sheet_build.md`, `controls_sheet_map.json`
- `monthly_build.md`, `monthly_map.json`
- `quarterly_build.md`, `quarterly_map.json`
- `input_sheet_build.md`, `input_sheet_map.json`
- `checks_build.md`, `checks_map.json`
- `formatting_guide.md`, `working_style.md`, `model_agent_operating_guide.md`, `excel_com_best_practices.md`

At initial publication, the six revised instruction documents totalled approximately **3,400 words**, versus **5,925 words of original Markdown** (about **43% shorter**). Including the five original JSON maps, the source set totalled 9,395 words; that is a different comparison because workbook maps still add workbook-specific content. These initial counts exclude this review note and subsequent operating-policy refinements.

At the initial documentation review, local links, calendar definitions, template references, scope, header placement and blank-check behaviour were checked for consistency; no workbook was built or tested at that stage. Subsequent workbook verification is recorded in its map. The large-model guidance remains a design requirement, not a measured runtime guarantee.
