// ============================================================
// Queue Sheet Operations
// ============================================================

import { readRange, appendRows, objectToRow, rowToObject } from './client';
import { QueueItem } from '../utils/types';
import { getJSTTimestamp } from '../utils/date';

const SHEET_NAME = 'queue';

// Column headers for writing (default order)
const HEADERS = [
  'bot_id',
  'queue_id',
  'scheduled_date',
  'scheduled_time',
  'category',
  'draft_text',
  'status',
  'guard_result',
  'output_json',
  'created_at',
];

/**
 * Read existing queue items for a bot
 */
export async function readQueueItems(botId: string): Promise<QueueItem[]> {
  const range = `${SHEET_NAME}!A:Z`; // Read all columns
  const rows = await readRange(range);
  
  if (rows.length < 2) {
    return [];
  }
  
  // Get headers from first row (dynamic)
  const headers = rows[0];
  const botIdIndex = headers.indexOf('bot_id');
  
  // Skip header row
  const dataRows = rows.slice(1);
  
  // Filter by bot_id (find bot_id column dynamically)
  const filtered = dataRows.filter((row) => {
    if (botIdIndex >= 0) {
      return row[botIdIndex] === botId;
    }
    return row[0] === botId; // Fallback to first column
  });
  
  // Convert to objects using dynamic headers
  const items: QueueItem[] = filtered.map((row) => {
    const obj = rowToObject<Record<string, string>>(headers, row);
    return {
      bot_id: obj.bot_id || '',
      queue_id: obj.queue_id || '',
      scheduled_date: obj.scheduled_date || '',
      scheduled_time: obj.scheduled_time || '',
      category: obj.category || '',
      draft_text: obj.draft_text || '',
      status: (obj.status as QueueItem['status']) || 'draft',
      guard_result: obj.guard_result || '',
      output_json: obj.output_json || '',
      created_at: obj.created_at || '',
    };
  });
  
  return items;
}

/**
 * Get existing queue IDs for a bot
 */
export async function getExistingQueueIds(botId: string): Promise<Set<string>> {
  const items = await readQueueItems(botId);
  return new Set(items.map((item) => item.queue_id));
}

/**
 * Check if a slot is already filled
 */
export async function isSlotFilled(
  botId: string,
  date: string,
  time: string
): Promise<boolean> {
  const items = await readQueueItems(botId);
  return items.some(
    (item) => item.scheduled_date === date && item.scheduled_time === time
  );
}

/**
 * Add queue items to the sheet
 */
export async function addQueueItems(items: QueueItem[]): Promise<void> {
  if (items.length === 0) return;
  
  const rows = items.map((item) => objectToRow(HEADERS, {
    ...item,
    created_at: item.created_at || getJSTTimestamp(),
  }));
  
  await appendRows(SHEET_NAME, rows);
  console.log(`[Queue] Added ${items.length} items to queue`);
}

/**
 * Create a new queue item
 */
export function createQueueItem(params: {
  botId: string;
  queueId: string;
  scheduledDate: string;
  scheduledTime: string;
  category: string;
  draftText: string;
  status: QueueItem['status'];
  guardResult: string;
  outputJson: string;
}): QueueItem {
  return {
    bot_id: params.botId,
    queue_id: params.queueId,
    scheduled_date: params.scheduledDate,
    scheduled_time: params.scheduledTime,
    category: params.category,
    draft_text: params.draftText,
    status: params.status,
    guard_result: params.guardResult,
    output_json: params.outputJson,
    created_at: getJSTTimestamp(),
  };
}
