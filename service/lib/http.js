import { createServer } from 'node:http';

// Local test adapter only. Vercel calls the Web Request handler directly.
export function serve(handler, port = 0) {
  const server = createServer(async (req, res) => {
    try {
      if (req.url !== '/api/mcp') { res.writeHead(404).end(); return; }
      const chunks = [];
      let size = 0;
      for await (const chunk of req) {
        size += chunk.length;
        if (size > 64 * 1024) { res.writeHead(413).end(); return; }
        chunks.push(chunk);
      }
      const url = `http://127.0.0.1:${server.address().port}${req.url}`;
      const body = ['GET', 'HEAD'].includes(req.method) ? undefined : Buffer.concat(chunks);
      const response = await handler(new Request(url, { method: req.method, headers: req.headers, body }));
      res.writeHead(response.status, Object.fromEntries(response.headers));
      res.end(Buffer.from(await response.arrayBuffer()));
    } catch { res.writeHead(500).end('Local request failed.'); }
  });
  server.listen(port, '127.0.0.1');
  return server;
}
