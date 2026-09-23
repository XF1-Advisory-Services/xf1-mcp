import selection from '../approved-release.json' with { type: 'json' };
import { createHandler } from '../lib/mcp.js';
import { resolveApprovedRelease } from '../lib/release.js';

const handler = createHandler(() => resolveApprovedRelease(selection));
export { handler as GET, handler as POST, handler as DELETE };
