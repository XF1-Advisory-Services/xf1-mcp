# Fresh inputs for each new workbook

Every request to generate a NEW Excel template or file using XF1 starts with an empty input set and a fresh unique request ID, even in the same conversation. Only values explicitly supplied for that request, answers to its clarifications, or reuse explicitly requested by the user belong to it. Do not import earlier assumptions, examples, saved defaults or previous request files. If the user asks for reuse ambiguously, clarify exactly which values they mean.

Collect six required values: StartDate, LastActualDate, FinancialYearEndMonth, Currency, Months and OutputPath. There is no implicit horizon for MCP builds. A complete request needs no redundant questions. ActualLabel and ForecastLabel are optional presentation settings, defaulting to A and F; do not carry custom labels over from another request.

Call `check_workbook_request` with a fresh `requestId` and a `supplied` object containing only the present fields and their `{requestId,source}` metadata. Sources are `current_request`, `clarification` or `explicit_reuse`. Do not send actual values, paths or conversation text. Show ALL returned missing questions together in their numbered order. Save that exact number-to-field mapping with the pending request ID. A reply such as `1: 2025-01-01; 2: 2026-08-31` answers that pending list. It is not another build request. Retain answers already supplied for this request; recheck and ask together for any still missing. If multiple requests are pending or a numbered answer is ambiguous, ask which request/list it answers.

After all six values are explicit, save a NEW local JSON file. This example illustrates syntax only; never use its values without the user's instruction:

```json
{
  "schemaVersion": 1,
  "requestId": "unique-id-for-this-request",
  "inputs": {
    "StartDate": {"value":"2025-01-01","requestId":"unique-id-for-this-request","source":"current_request"},
    "LastActualDate": {"value":"2026-08-31","requestId":"unique-id-for-this-request","source":"clarification"},
    "FinancialYearEndMonth": {"value":12,"requestId":"unique-id-for-this-request","source":"clarification"},
    "Currency": {"value":"USD","requestId":"unique-id-for-this-request","source":"explicit_reuse"},
    "Months": {"value":48,"requestId":"unique-id-for-this-request","source":"clarification"},
    "OutputPath": {"value":"C:\\Models\\New-model.xlsx","requestId":"unique-id-for-this-request","source":"current_request"}
  }
}
```

Call `get_current_release`, verify the approved bridge hash, then invoke the bridge with `-McpUrl` and `-RequestPath`. Bridge 1.1.0 requires this request file and refuses legacy scalar build arguments. It rejects missing inputs together, stale/missing provenance, invalid source types and invalid core values before network, cache or Excel work. `-ValidateRequestOnly` validates locally without contacting MCP or opening Excel. `-ResolveOnly` discovers/verifies the package without authorizing a workbook build.

Ordinary follow-up edits to an existing workbook keep that workbook's assumptions and context. Do not run this collection/build flow or create a replacement file unless the user requests a new workbook. The preflight's `edit_existing` operation returns no new-input questions and never grants build readiness.

The service is stateless and cannot inspect the conversation. Provenance is asserted by the client; neither a fresh-looking ID nor a source label proves what the user said. The bridge rejects inconsistent records but cannot detect a client copying old values while falsely labelling them current, or replaying an entire prior request file. Direct development runner use is outside this MCP gate and retains its legacy API. Instructions plus boundary validation improve compliance; they do not provide complete control over an LLM's context.
