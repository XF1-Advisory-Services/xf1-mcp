import test from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { once } from 'node:events';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { readFile, writeFile, mkdir, appendFile, stat } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { createHandler } from '../lib/mcp.js';
import { serve } from '../lib/http.js';
import { resolveApprovedRelease, validateRelease } from '../lib/release.js';
const unapprovedHandler = createHandler(() => resolveApprovedRelease({ schemaVersion: 1, release: null }));

const exec = promisify(execFile);
const root = fileURLToPath(new URL('../../', import.meta.url));
const work = path.join(root, '.build', 'distribution', new Date().toISOString().replaceAll(/[:.]/g, '-'));
const bridge = path.join(root, 'distribution', 'Invoke-XF1Build.ps1');
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const clone = (value) => structuredClone(value);
const readJson = async (file) => JSON.parse((await readFile(file, 'utf8')).replace(/^\uFEFF/, ''));
const powershell = (args) => exec('pwsh', ['-NoProfile', '-File', ...args], { maxBuffer: 2_000_000, timeout: 120_000 });
let failures = 0;
const evidence = [];

test('XF1 distribution contract and local bridge', async (t) => {
  await mkdir(work, { recursive: true });
  for (const version of ['1.2.0', '1.2.1']) {
    await powershell([path.join(root, 'distribution', 'Prepare-XF1Release.ps1'), '-Version', version, '-OutputDirectory', path.join(work, version)]);
  }
  const candidate = await readJson(path.join(work, '1.2.0', 'release-candidate.json'));
  const approved = { ...candidate, approval: 'approved' };
  const zip = await readFile(path.join(work, '1.2.0', 'xf1-1.2.0.zip'));
  const originalManifestHash = hash(await readFile(path.join(root, 'release-manifest.json')));
  const originalApprovalHash = hash(await readFile(path.join(root, 'service', 'approved-release.json')));
  let downloads = 0;
  let artifactBytes = zip;
  let onDownload = () => {};
  const assets = createServer(async (req, res) => {
    downloads++;
    await onDownload();
    res.writeHead(200, { 'content-type': 'application/zip' });
    res.end(artifactBytes);
  });
  assets.listen(0, '127.0.0.1'); await once(assets, 'listening');
  const assetBase = `http://127.0.0.1:${assets.address().port}`;
  const local = { ...approved, package: { ...approved.package, url: `${assetBase}/package.zip` }, bridge: { ...approved.bridge, url: `${assetBase}/Invoke-XF1Build.ps1` } };
  let current = local;
  let calls = 0;
  let unavailable = false;
  const server = serve(createHandler(() => {
    calls++;
    if (unavailable) throw new Error('No approved XF1 release is configured.');
    return clone(current);
  }));
  await once(server, 'listening');
  const endpoint = `http://127.0.0.1:${server.address().port}/api/mcp`;
  const cache = path.join(work, 'cache');
  const invoke = async (extra = [], cachePath = cache) => {
    const result = await powershell([bridge, '-McpUrl', endpoint, '-AllowLoopbackForTest', '-CacheRoot', cachePath, ...extra]);
    return JSON.parse(result.stdout);
  };
  const resolve = (cachePath = cache) => invoke(['-ResolveOnly'], cachePath);
  const check = async (name, fn) => t.test(name, async () => {
    try { await fn(); evidence.push({ name, passed: true }); }
    catch (error) { failures++; evidence.push({ name, passed: false, error: error.message }); throw error; }
  });
  try {
    await check('release selection rejects null and unapproved candidates', () => {
      assert.throws(() => resolveApprovedRelease({ schemaVersion: 1, release: null }), /No approved/);
      assert.throws(() => validateRelease(candidate), /invalid/);
      assert.equal(validateRelease(approved).version, '1.2.0');
    });
    await check('release selection rejects mutable URLs, wrong hashes and unsafe versions', () => {
      for (const mutate of [r => r.package.url += '?latest=1', r => r.package.sha256 = 'bad', r => r.version = '../other', r => r.package.bytes = 8_000_000, r => r.runner = 'other.ps1', r => r.bridge.version = '2.0.0']) {
        const value = clone(approved); mutate(value); assert.throws(() => validateRelease(value));
      }
    });
    await check('production MCP initializes and lists exactly the release tool', async () => {
      const init = await unapprovedHandler(new Request(endpoint, { method: 'POST', headers: { 'content-type': 'application/json', accept: 'application/json, text/event-stream' }, body: JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1' } } }) }));
      assert.equal(init.status, 200);
      assert.match(await init.text(), /xf1-mcp/);
      const listed = await rpc(unapprovedHandler, 'tools/list');
      assert.equal(listed.result.tools.length, 1);
      assert.equal(listed.result.tools[0].name, 'get_current_release');
    });
    await check('production endpoint refuses workbook approval until explicitly configured', async () => {
      const data = await rpc(unapprovedHandler, 'tools/call', { name: 'get_current_release', arguments: {} });
      assert.equal(data.result.isError, true);
      assert.match(data.result.content[0].text, /No approved/);
    });
    await check('production response is not cacheable and rejects unrelated web origins', async () => {
      const response = await unapprovedHandler(new Request(endpoint, { method: 'POST', headers: { origin: 'https://other.example' } }));
      assert.equal(response.status, 403);
      const response2 = await unapprovedHandler(new Request(endpoint, { method: 'POST', headers: { 'content-type': 'application/json', accept: 'application/json, text/event-stream', 'MCP-Protocol-Version': '2025-06-18' }, body: JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/list' }) }));
      assert.equal(response2.headers.get('cache-control'), 'no-store');
    });
    await check('first resolution downloads and verifies a workbook-free package', async () => {
      const result = await resolve(); assert.equal(result.downloaded, true); assert.equal(downloads, 1);
      assert.equal(result.releaseVersion, '1.2.0');
      const manifest = await readJson(path.join(result.packageDirectory, 'release-manifest.json'));
      assert.equal(manifest.assets.length, (await readJson(path.join(root, 'release-manifest.json'))).assets.length);
      assert.ok(manifest.assets.every(a => !/xlsx|workbook-map|\.build/.test(a.path)));
    });
    await check('exact cached version is reused after a fresh MCP approval check', async () => {
      const before = calls;
      const result = await resolve(); assert.equal(result.downloaded, false); assert.equal(downloads, 1); assert.equal(calls, before + 1);
    });
    await check('unavailable approved selection cannot fall back to a warm cache', async () => {
      unavailable = true;
      try { await assert.rejects(resolve(), /No approved/); } finally { unavailable = false; }
      assert.equal(downloads, 1);
    });
    await check('bridge refuses an altered bootstrap script hash', async () => {
      current = clone(local); current.bridge.sha256 = '0'.repeat(64);
      try { await assert.rejects(resolve(), /hash mismatch/); } finally { current = local; }
    });
    await check('a corrupt package download is rejected and not installed', async () => {
      current = clone(local); current.package.sha256 = '0'.repeat(64);
      try { await assert.rejects(resolve(path.join(work, 'bad-download')), /hash mismatch/); } finally { current = local; }
      await assert.rejects(stat(path.join(work, 'bad-download', 'xf1-base-six-sheet', '1.2.0', 'package')));
    });
    await check('tampered cache files fail without being silently redownloaded', async () => {
      const location = path.join(cache, 'xf1-base-six-sheet', '1.2.0', 'package', 'runner', 'Layout.ps1');
      const original = await readFile(location); const before = downloads;
      await appendFile(location, '\n# modified');
      try { await assert.rejects(resolve(), /hash mismatch/); assert.equal(downloads, before); }
      finally { await writeFile(location, original); }
    });
    await check('package identity remains fixed even if current approval changes during download', async () => {
      onDownload = () => { current = { ...local, version: '1.2.1' }; };
      try { const result = await resolve(path.join(work, 'pinned')); assert.equal(result.releaseVersion, '1.2.0'); }
      finally { onDownload = () => {}; current = local; }
    });
    await check('a new approved version gets a separate cache directory', async () => {
      const next = await readJson(path.join(work, '1.2.1', 'release-candidate.json'));
      current = { ...next, approval: 'approved', package: { ...next.package, url: local.package.url }, bridge: local.bridge };
      artifactBytes = await readFile(path.join(work, '1.2.1', 'xf1-1.2.1.zip'));
      try {
        const result = await resolve(); assert.equal(result.releaseVersion, '1.2.1'); assert.equal(result.downloaded, true);
        assert.ok(await stat(path.join(cache, 'xf1-base-six-sheet', '1.2.0', 'package')));
      } finally { current = local; artifactBytes = zip; }
    });
    await check('wrong manifest identity cannot reuse an existing version directory', async () => {
      current = clone(local); current.manifestSha256 = '0'.repeat(64);
      try { await assert.rejects(resolve(), /hash mismatch/); } finally { current = local; }
    });
    await check('concurrent resolution of the same version is locked until verification completes', async () => {
      let releaseDownload, signalDownload;
      const started = new Promise(resolve => { signalDownload = resolve; });
      const gate = new Promise(resolve => { releaseDownload = resolve; });
      onDownload = async () => { signalDownload(); await gate; };
      const location = path.join(work, 'locked');
      const first = resolve(location);
      try {
        await started;
        await assert.rejects(resolve(location), /in use by another local build/);
      } finally { releaseDownload(); onDownload = () => {}; }
      assert.equal((await first).downloaded, true);
    });
    await check('ZIP traversal is rejected before extraction', async () => {
      const evil = path.join(work, 'traversal.zip');
      await powershell([path.join(root, 'tests', 'New-XF1BadZip.ps1'), '-OutputPath', evil]);
      artifactBytes = await readFile(evil);
      current = clone(local); current.package.sha256 = hash(artifactBytes); current.package.bytes = artifactBytes.length;
      try { await assert.rejects(resolve(path.join(work, 'bad-path')), /Unapproved package path/); }
      finally { current = local; artifactBytes = zip; }
    });
    await check('test mode cannot be used with a remote MCP endpoint', async () => {
      await assert.rejects(powershell([bridge, '-McpUrl', 'https://example.com/api/mcp', '-AllowLoopbackForTest', '-ResolveOnly']), /requires a loopback/);
    });
    await check('plain HTTP is forbidden without explicit loopback test mode', async () => {
      await assert.rejects(powershell([bridge, '-McpUrl', endpoint, '-ResolveOnly']), /must use HTTPS/);
    });
    if (process.env.XF1_TEST_EXCEL === '1') {
      await check('native Excel end-to-end: MCP, verified cache, existing runner, workbook and sidecars', async () => {
        const output = path.join(work, 'native-build.xlsx');
        const result = await invoke(['-StartDate', '2025-01-01', '-LastActualDate', '2026-08-31', '-FinancialYearEndMonth', '12', '-Currency', 'USD', '-Months', '48', '-OutputPath', output]);
        assert.equal(result.downloaded, false);
        assert.equal(result.build.releaseVersion, '1.2.0');
        assert.equal(result.build.generation, 'from-scratch');
        assert.equal(result.build.verification.formulaErrors, 0);
        assert.equal(result.build.verification.sheets.length, 6);
        assert.equal(result.build.manifestSHA256.toLowerCase(), candidate.manifestSha256.toLowerCase());
        assert.ok((await stat(output)).size > 0);
        const map = await readJson(output.replace('.xlsx', '.map.json'));
        assert.equal(map.configuration.months, 48);
        assert.equal((await readJson(output.replace('.xlsx', '.build.json'))).releaseVersion, '1.2.0');
        await assert.rejects(invoke(['-StartDate', '2025-01-01', '-LastActualDate', '2026-08-31', '-FinancialYearEndMonth', '12', '-Currency', 'USD', '-OutputPath', output]), /Output already exists/);
      });
    }
    await check('connection failure stops even with a verified cache', async () => {
      await new Promise(resolve => server.close(resolve));
      await assert.rejects(resolve());
    });
    await check('local tests do not change the development manifest or production approval', async () => {
      assert.equal(hash(await readFile(path.join(root, 'release-manifest.json'))), originalManifestHash);
      assert.equal(hash(await readFile(path.join(root, 'service', 'approved-release.json'))), originalApprovalHash);
    });
  } finally {
    server.closeAllConnections(); server.close();
    assets.closeAllConnections(); await new Promise(resolve => assets.close(resolve));
    await writeFile(path.join(work, 'results.json'), JSON.stringify({ completedAt: new Date().toISOString(), nativeExcel: process.env.XF1_TEST_EXCEL === '1', failures, evidence }, null, 2));
    console.log(`Evidence: ${work}`);
  }
});

async function rpc(handler, method, params = {}) {
  const response = await handler(new Request('http://127.0.0.1/api/mcp', { method: 'POST', headers: { 'content-type': 'application/json', accept: 'application/json, text/event-stream', 'MCP-Protocol-Version': '2025-06-18' }, body: JSON.stringify({ jsonrpc: '2.0', id: 7, method, params }) }));
  assert.equal(response.status, 200);
  const text = await response.text();
  return JSON.parse(text.startsWith('{') ? text : text.split('\n').find(line => line.startsWith('data:')).slice(5));
}
