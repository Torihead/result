// Clear queue sheet for fresh generation
import * as dotenv from 'dotenv';
dotenv.config();

import { loadConfig } from '../src/config';
import { initSheetsClient, updateRange } from '../src/sheets/client';

async function main() {
  const appConfig = loadConfig();
  await initSheetsClient(appConfig);
  
  // Clear queue by keeping only headers
  const headers = ['bot_id', 'queue_id', 'scheduled_date', 'scheduled_time', 'category', 'draft_text', 'status', 'guard_result', 'output_json', 'created_at'];
  await updateRange('queue!A1:J1', [headers]);
  
  // Clear data rows (overwrite with empty)
  await updateRange('queue!A2:J100', Array(99).fill(['', '', '', '', '', '', '', '', '', '']));
  
  console.log('✅ Queue cleared!');
}

main().catch(console.error);
