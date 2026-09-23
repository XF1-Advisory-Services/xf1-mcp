import { POST } from './api/mcp.js';
import { serve } from './lib/http.js';
const server = serve(POST, Number(process.env.PORT || 3000));
server.on('listening', () => console.log(`XF1 MCP: http://127.0.0.1:${server.address().port}/api/mcp`));
