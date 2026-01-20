// ============================================================
// Config Sheet Operations
// ============================================================

import { readRange, rowToObject } from './client';
import { BotConfig } from '../utils/types';

const SHEET_NAME = 'config';

/**
 * Read bot configuration from config sheet
 * Sheet format: bot_id, key, value, note
 */
export async function readBotConfig(botId: string): Promise<BotConfig> {
  const range = `${SHEET_NAME}!A:D`;
  const rows = await readRange(range);
  
  if (rows.length < 2) {
    throw new Error(`No config found for bot_id: ${botId}`);
  }
  
  // Skip header row
  const dataRows = rows.slice(1);
  
  // Filter by bot_id and build config object
  const config: BotConfig = { bot_id: botId };
  
  for (const row of dataRows) {
    const [rowBotId, key, value] = row;
    if (rowBotId === botId && key) {
      // Try to parse as number if applicable
      const numValue = parseFloat(value);
      if (!isNaN(numValue) && key !== 'ng_words' && key !== 'policy_constraints') {
        config[key] = numValue;
      } else {
        config[key] = value || '';
      }
    }
  }
  
  // Validate required fields
  if (Object.keys(config).length <= 1) {
    throw new Error(`No config entries found for bot_id: ${botId}`);
  }
  
  return config;
}

/**
 * Get all bot IDs from config
 */
export async function getAllBotIds(): Promise<string[]> {
  const range = `${SHEET_NAME}!A:A`;
  const rows = await readRange(range);
  
  if (rows.length < 2) {
    return [];
  }
  
  // Skip header, get unique bot_ids
  const botIds = new Set<string>();
  rows.slice(1).forEach((row) => {
    if (row[0]) {
      botIds.add(row[0]);
    }
  });
  
  return Array.from(botIds);
}
