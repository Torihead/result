// ============================================================
// History Sheet Operations
// ============================================================

import { readRange, rowToObject } from './client';
import { HistoryPost } from '../utils/types';

const SHEET_NAME = 'history';
const MAX_ROWS = 30;

/**
 * Read recent history posts for a bot (latest 30)
 */
export async function readHistoryPosts(botId: string): Promise<HistoryPost[]> {
  const range = `${SHEET_NAME}!A:Z`; // Read all columns
  const rows = await readRange(range);
  
  if (rows.length < 2) {
    console.log(`[History] No history found for bot_id: ${botId}`);
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
  
  // Take latest 30 (assuming rows are ordered by posted_at desc or newest at bottom)
  const latest = filtered.slice(-MAX_ROWS);
  
  // Convert to objects using dynamic headers
  const posts: HistoryPost[] = latest.map((row) => {
    const obj = rowToObject<Record<string, string>>(headers, row);
    return {
      bot_id: obj.bot_id || '',
      posted_at: obj.posted_at || '',
      text: obj.text || '',
      category: obj.category || '',
      likes: parseInt(obj.likes, 10) || 0,
      impressions: parseInt(obj.impressions, 10) || 0,
      note: obj.note || '',
    };
  });
  
  console.log(`[History] Loaded ${posts.length} history posts for bot_id: ${botId}`);
  return posts;
}

/**
 * Get recent text samples for repetition checking
 */
export function getRecentTexts(posts: HistoryPost[]): string[] {
  return posts.map((p) => p.text).filter(Boolean);
}

/**
 * Check if text is too similar to recent posts
 */
export function isSimilarToRecent(
  text: string,
  recentTexts: string[],
  threshold: number = 0.7
): boolean {
  const normalizedText = normalizeText(text);
  
  for (const recent of recentTexts) {
    const normalizedRecent = normalizeText(recent);
    const similarity = calculateSimilarity(normalizedText, normalizedRecent);
    if (similarity >= threshold) {
      return true;
    }
  }
  
  return false;
}

/**
 * Normalize text for comparison
 */
function normalizeText(text: string): string {
  return text
    .toLowerCase()
    .replace(/[^\w\s\u3040-\u309f\u30a0-\u30ff\u4e00-\u9faf]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Simple Jaccard similarity for text comparison
 */
function calculateSimilarity(text1: string, text2: string): number {
  const words1 = new Set(text1.split(' '));
  const words2 = new Set(text2.split(' '));
  
  const intersection = new Set([...words1].filter((w) => words2.has(w)));
  const union = new Set([...words1, ...words2]);
  
  if (union.size === 0) return 0;
  return intersection.size / union.size;
}
