import { createMcpHandler } from 'mcp-handler';
import { z } from 'zod';

export const instructions = `XF1 creates NEW six-sheet workbooks using desktop Excel on the user's Windows laptop. It cannot create files remotely or upgrade an existing model. Before EVERY build call get_current_release; if that fails, stop. Use only the returned approved version and verify hashes. Keep that version fixed throughout the build. Initial setup requires Windows, PowerShell 7+, usable desktop Excel with Formula2, and a local agent allowed to run PowerShell. Download the returned bridge URL once, verify its SHA-256 against bridge.sha256, then run the local Invoke-XF1Build.ps1 with -McpUrl and the user's StartDate, LastActualDate, FinancialYearEndMonth, Currency and OutputPath; Months defaults to 48. The bridge makes its own fresh MCP check, verifies or downloads the exact package, and calls the existing runner. If approval changes between discovery and execution the bridge uses the fresh selection. Never reconstruct Excel logic from prose, upload the workbook, pass credentials, or silently use an old cache. Codex is the tested client; Claude compatibility is expected but untested.`;

export function createHandler(resolveRelease) {
  const handler = createMcpHandler((server) => {
    server.registerTool('get_current_release', {
      title: 'Get the approved XF1 workbook release',
      description: 'Required before each NEW local XF1 workbook. Returns verified release identity, asset references and local execution instructions. No Excel execution occurs on this server.',
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
