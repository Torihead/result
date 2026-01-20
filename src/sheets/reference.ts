// ============================================================
// Reference Posts Sheet Operations
// ============================================================

import { readRange, rowToObject } from './client';
import { ReferencePost } from '../utils/types';

const SHEET_NAME = 'reference_posts';
const MAX_ROWS = 30;

/**
 * Read reference posts for a bot (latest 30)
 */
export async function readReferencePosts(botId: string): Promise<ReferencePost[]> {
  const range = `${SHEET_NAME}!A:Z`; // Read all columns
  const rows = await readRange(range);
  
  if (rows.length < 2) {
    console.log(`[Reference] No reference posts found for bot_id: ${botId}`);
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
  
  // Take latest 30 (assuming rows are ordered by added_at desc or newest at bottom)
  const latest = filtered.slice(-MAX_ROWS);
  
  // Convert to objects using dynamic headers
  const posts: ReferencePost[] = latest.map((row) => {
    const obj = rowToObject<Record<string, string>>(headers, row);
    return {
      bot_id: obj.bot_id || '',
      ref_id: obj.ref_id || '',
      url: obj.url || '',
      text: obj.text || '',
      category: obj.category || '',
      likes: parseInt(obj.likes, 10) || 0,
      retweets: parseInt(obj.retweets, 10) || 0,
      replies: parseInt(obj.replies, 10) || 0,
      impressions: parseInt(obj.impressions, 10) || 0,
      engagement_rate: parseFloat(obj.engagement_rate) || 0,
      note: obj.note || '',
      added_at: obj.added_at || '',
    };
  });
  
  console.log(`[Reference] Loaded ${posts.length} reference posts for bot_id: ${botId}`);
  return posts;
}

/**
 * Get high-performing reference posts (by engagement rate)
 */
export function getTopPerformingPosts(
  posts: ReferencePost[],
  limit: number = 10
): ReferencePost[] {
  return [...posts]
    .sort((a, b) => b.engagement_rate - a.engagement_rate)
    .slice(0, limit);
}

/**
 * Get categories from reference posts
 */
export function getCategories(posts: ReferencePost[]): string[] {
  const categories = new Set<string>();
  posts.forEach((post) => {
    if (post.category) {
      categories.add(post.category);
    }
  });
  return Array.from(categories);
}
