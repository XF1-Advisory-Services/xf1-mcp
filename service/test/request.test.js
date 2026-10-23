import test from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { mkdtemp, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { checkWorkbookRequest, questions } from '../lib/request.js';
import { createHandler } from '../lib/mcp.js';

const fields = Object.keys(questions);
const metadata = (requestId, source = 'current_request', selected = fields) => Object.fromEntries(selected.map(field => [field, { requestId, source }]));
const values = { StartDate: '2025-01-01', LastActualDate: '2025-12-31', FinancialYearEndMonth: 12, Currency: 'EUR', Months: 24, OutputPath: path.join(os.tmpdir(), 'xf1-request-test.xlsx') };
const record = (requestId, source = 'current_request') => ({ schemaVersion: 1, requestId, inputs: Object.fromEntries(fields.map(field => [field, { requestId, source, value: values[field] }])) });

test('partial request asks all and only missing inputs in a numbered list', () => {
  const result = checkWorkbookRequest({ requestId: 'one', supplied: metadata('one', 'current_request', ['Currency']) });
  assert.equal(result.ready, false);
  assert.deepEqual(result.questions.map(q => q.field), fields.filter(f => f !== 'Currency'));
  assert.deepEqual(result.questions.map(q => q.number), [1, 2, 3, 4, 5]);
  assert.equal(result.clarification.split('\n').length, 5);
});
test('complete current request is ready without redundant questions', () => {
  const result = checkWorkbookRequest({ requestId: 'one', supplied: metadata('one') });
  assert.equal(result.ready, true); assert.deepEqual(result.questions, []);
});
test('same-thread second request starts empty and stale provenance is rejected', () => {
  checkWorkbookRequest({ requestId: 'one', supplied: metadata('one') });
  const result = checkWorkbookRequest({ requestId: 'two' });
  assert.equal(result.questions.length, 6); assert.equal(result.ready, false);
  assert.throws(() => checkWorkbookRequest({ requestId: 'two', supplied: metadata('one') }), /different request/);
});
test('numbered clarifications bind to the pending mapping, including partial answers', () => {
  const supplied = metadata('one', 'current_request', ['Currency']);
  const pending = checkWorkbookRequest({ requestId: 'one', supplied });
  // Client resolves the user's "1: ...; 3: ..." using this request's saved mapping.
  for (const number of [1, 3]) supplied[pending.questions.find(q => q.number === number).field] = { requestId: 'one', source: 'clarification' };
  const remaining = checkWorkbookRequest({ requestId: 'one', supplied });
  assert.deepEqual(remaining.questions.map(q => q.field), ['LastActualDate', 'Months', 'OutputPath']);
  for (const q of remaining.questions) supplied[q.field] = { requestId: 'one', source: 'clarification' };
  assert.equal(checkWorkbookRequest({ requestId: 'one', supplied }).ready, true);
});
test('explicit reuse belongs to the new request; implicit/history sources are forbidden', () => {
  assert.equal(checkWorkbookRequest({ requestId: 'two', supplied: metadata('two', 'explicit_reuse') }).ready, true);
  assert.throws(() => checkWorkbookRequest({ requestId: 'two', supplied: metadata('two', 'history') }));
});
test('ordinary workbook edits do not reset or recollect assumptions', () => {
  const result = checkWorkbookRequest({ requestId: 'edit', operation: 'edit_existing' });
  assert.equal(result.status, 'existing_workbook_edit'); assert.equal(result.ready, false); assert.deepEqual(result.questions, []);
});
test('MCP boundary exposes preflight and rejects values and stale request inputs', async () => {
  const handler = createHandler(() => { throw new Error('Not needed for input collection'); });
  const call = async (method, params = {}) => {
    const response = await handler(new Request('http://127.0.0.1/api/mcp', { method: 'POST', headers: { 'content-type': 'application/json', accept: 'application/json, text/event-stream', 'MCP-Protocol-Version': '2025-06-18' }, body: JSON.stringify({ jsonrpc: '2.0', id: 1, method, params }) }));
    const body = await response.text(); return JSON.parse(body.startsWith('{') ? body : body.split('\n').find(s => s.startsWith('data:')).slice(5));
  };
  assert.deepEqual((await call('tools/list')).result.tools.map(t => t.name).sort(), ['check_workbook_request', 'get_current_release']);
  const missing = await call('tools/call', { name: 'check_workbook_request', arguments: { requestId: 'one' } });
  assert.equal(missing.result.structuredContent.questions.length, 6);
  const stale = await call('tools/call', { name: 'check_workbook_request', arguments: { requestId: 'two', supplied: metadata('one') } });
  assert.equal(stale.result.isError, true);
  const invalid = await call('tools/call', { name: 'check_workbook_request', arguments: { requestId: 'one', supplied: { Currency: { requestId: 'one', source: 'current_request', value: 'EUR' } } } });
  assert.ok(invalid.error || invalid.result?.isError);
});
test('bridge enforces completeness and provenance before any network or Excel action', async (t) => {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'xf1-request-'));
  const bridge = fileURLToPath(new URL('../../distribution/Invoke-XF1Build.ps1', import.meta.url));
  const exec = promisify(execFile);
  let number = 0;
  const validate = async request => {
    const file = path.join(directory, `${++number}.json`); await writeFile(file, JSON.stringify(request));
    return exec('pwsh', ['-NoProfile', '-File', bridge, '-McpUrl', 'https://invalid.example/api/mcp', '-RequestPath', file, '-ValidateRequestOnly'], { timeout: 20000 });
  };
  await t.test('complete, clarified and explicitly reused request values pass', async () => {
    for (const source of ['current_request', 'clarification', 'explicit_reuse']) assert.equal(JSON.parse((await validate(record(source, source))).stdout).ready, true);
  });
  await t.test('second request omitted inputs do not inherit prior values or 48 months', async () => {
    const request = record('two'); delete request.inputs.Months; delete request.inputs.Currency;
    await assert.rejects(validate(request), error => /1\. Which currency/.test(error.stderr) && /2\. How many months/.test(error.stderr));
  });
  await t.test('stale provenance, implicit reuse and invalid values stop', async () => {
    const stale = record('two'); stale.inputs.Currency.requestId = 'one'; await assert.rejects(validate(stale), /stale provenance/);
    const implicit = record('two', 'history'); await assert.rejects(validate(implicit), /provenance/);
    const bad = record('one'); bad.inputs.Months.value = 1; await assert.rejects(validate(bad), /Months must/);
    const date = record('one'); date.inputs.StartDate.value = '2025-02-01'; await assert.rejects(validate(date), /first day/);
  });
  await t.test('legacy scalar calls cannot bypass the request-file gate', async () => {
    await assert.rejects(exec('pwsh', ['-NoProfile', '-File', bridge, '-McpUrl', 'https://invalid.example/api/mcp', '-StartDate', '2025-01-01']), /parameter.*StartDate/is);
  });
});
