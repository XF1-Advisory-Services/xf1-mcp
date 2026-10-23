import { createMcpHandler } from 'mcp-handler';
import { z } from 'zod';
import { checkWorkbookRequest, requestSchema } from './request.js';

export const instructions = `XF1 creates NEW six-sheet workbooks using desktop Excel on the user's Windows laptop. For EVERY new Excel template/file request, including another build in the same thread, create a fresh unique requestId and start with EMPTY inputs. Include only values explicitly supplied for this request, its numbered clarifications, or explicitly requested reuse. Never infer values from earlier thread assumptions, an earlier request file, saved workbook, examples or defaults. Call check_workbook_request with presence/provenance metadata only (no values or paths). Required inputs: StartDate, LastActualDate, FinancialYearEndMonth, Currency, Months, OutputPath. Ask ALL missing questions together as the returned numbered list, preserving the number-to-field mapping for this request. Numbered replies continue that request; they do not start another request. Explicit reuse must identify the values being reused; ask if ambiguous. Ordinary follow-up edits to an existing workbook retain their context and do not use the new-workbook workflow. After preflight is ready, save a local JSON object {schemaVersion:1,requestId,inputs:{Field:{value,requestId,source}}}; source is current_request, clarification or explicit_reuse. Include all six required fields, optionally ActualLabel/ForecastLabel; only those presentation labels default to A/F. Before EVERY build call get_current_release; if it fails, stop. Verify the returned bridge SHA-256 and invoke Invoke-XF1Build.ps1 with -McpUrl and -RequestPath. The bridge validates request provenance, completeness and values before network, cache or Excel operations, then obtains fresh approval, verifies the exact package and runs the tested builder locally. Initial setup requires Windows, PowerShell 7+, usable desktop Excel with Formula2 and local execution permission. Keep the selected version fixed throughout the build. Never reconstruct Excel logic from prose, upload workbook data, pass credentials or use a stale cache. The service and bridge cannot prove what the user said: provenance is a client assertion, not authentication or control over LLM context. Codex is the tested client; Claude compatibility is untested.`;

export function createHandler(resolveRelease) {
  const handler = createMcpHandler((server) => {
    server.registerTool('check_workbook_request', {
      title: 'Collect fresh inputs for a new XF1 workbook',
      description: 'Call for every NEW Excel file/template, even in the same thread. Start empty; submit only field presence and current-request provenance. Returns ALL missing inputs as numbered questions. Never send values, paths or conversation text. Ordinary edits retain existing context.',
      inputSchema: requestSchema,
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    }, async (input) => {
      try {
        const result = checkWorkbookRequest(input);
        return { content: [{ type: 'text', text: JSON.stringify(result) }], structuredContent: result };
      } catch (error) {
        return { isError: true, content: [{ type: 'text', text: error.message }] };
      }
    });
    server.registerTool('get_current_release', {
      title: 'Get the approved XF1 workbook release',
      description: 'Required before each NEW local XF1 workbook. First collect fresh inputs with check_workbook_request and ask all missing inputs as numbered questions; never reuse prior assumptions implicitly. Returns approved release and local execution instructions; release discovery alone does not establish input readiness.',
      inputSchema: z.object({}).strict(),
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    }, async () => {
      try {
        const result = { release: await resolveRelease(), instructions };
        return { content: [{ type: 'text', text: JSON.stringify(result) }], structuredContent: result };
      } catch (error) {
        return { isError: true, content: [{ type: 'text', text: error.message }] };
      }
    });
  }, { serverInfo: { name: 'xf1-mcp', version: '0.1.0' }, instructions });

  return async (request) => {
    // No web UI needs cross-origin access. This is not user authentication.
    const origin = request.headers.get('origin');
    if (origin && origin !== new URL(request.url).origin) {
      return new Response('Cross-origin requests are not supported.', { status: 403 });
    }
    const response = await handler(request);
    response.headers.set('Cache-Control', 'no-store');
    response.headers.set('CDN-Cache-Control', 'no-store');
    response.headers.set('Vercel-CDN-Cache-Control', 'no-store');
    return response;
  };
}
