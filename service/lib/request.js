import { z } from 'zod';

export const questions = {
  StartDate: 'What is the model start date (YYYY-MM-DD, first day of its financial year)?',
  LastActualDate: 'What is the last actuals date (YYYY-MM-DD, month end; for no actuals use the day before the start date)?',
  FinancialYearEndMonth: 'Which month ends the financial year (1-12)?',
  Currency: 'Which currency should this workbook use (for example USD or EUR)?',
  Months: 'How many months should this workbook cover (at least 12)?',
  OutputPath: 'What is the new workbook\'s full Windows output path, including its .xlsx filename?',
};
const id = z.string().trim().min(1).max(128);
const provenance = z.object({
  requestId: id,
  source: z.enum(['current_request', 'clarification', 'explicit_reuse']),
}).strict();
export const requestSchema = z.object({
  requestId: id.describe('Fresh unique ID for this NEW workbook request; retain only while collecting its answers.'),
  operation: z.enum(['new_workbook', 'edit_existing']).default('new_workbook'),
  supplied: z.object(Object.fromEntries(Object.keys(questions).map(field => [field, provenance.optional()]))).strict().default({}),
}).strict();

export function checkWorkbookRequest(input) {
  const { requestId, operation, supplied } = requestSchema.parse(input);
  if (operation === 'edit_existing') {
    return { requestId, status: 'existing_workbook_edit', ready: false, questions: [], instructions: 'Keep the existing workbook context. Do not run the new-workbook bridge or reset assumptions for an ordinary edit.' };
  }
  const stale = Object.keys(supplied).filter(field => supplied[field].requestId !== requestId);
  if (stale.length) throw new Error(`Answers belong to a different request: ${stale.join(', ')}. Start empty for each new workbook; explicit reuse must be attributed to the current request.`);
  const missing = Object.keys(questions).filter(field => !supplied[field]);
  const numbered = missing.map((field, index) => ({ number: index + 1, field, question: questions[field] }));
  return {
    requestId, status: missing.length ? 'needs_input' : 'ready', ready: missing.length === 0,
    questions: numbered,
    clarification: numbered.map(item => `${item.number}. ${item.question}`).join('\n'),
    instructions: missing.length
      ? 'Ask ALL questions together exactly as numbered. Keep this number-to-field mapping with this requestId. Apply numbered answers only to this pending request; retain answered fields and ask together for any still missing. Never infer reuse or fill missing inputs from earlier requests or defaults.'
      : 'Save the current request values and their provenance in a local request JSON file and invoke the approved bridge with -RequestPath. This preflight checks provenance metadata only; the local bridge validates actual values. Do not send values, paths or conversation text to this server.',
  };
}
