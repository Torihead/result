// ============================================================
// Date Utilities (JST-focused)
// ============================================================

/**
 * Get current date/time in JST
 */
export function getJSTNow(): Date {
  const now = new Date();
  // Convert to JST (UTC+9)
  const jstOffset = 9 * 60 * 60 * 1000;
  const utc = now.getTime() + now.getTimezoneOffset() * 60 * 1000;
  return new Date(utc + jstOffset);
}

/**
 * Format date as YYYY-MM-DD
 */
export function formatDate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Format date as YYYYMMDD (for queue_id)
 */
export function formatDateCompact(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

/**
 * Format time as HH:MM
 */
export function formatTime(date: Date): string {
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${hours}:${minutes}`;
}

/**
 * Format time as HHMM (for queue_id)
 */
export function formatTimeCompact(time: string): string {
  return time.replace(':', '');
}

/**
 * Get today and tomorrow in JST as YYYY-MM-DD
 */
export function getTargetDates(): string[] {
  const today = getJSTNow();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  return [formatDate(today), formatDate(tomorrow)];
}

/**
 * Default posting times
 */
export const DEFAULT_TIMES = ['08:10', '12:40', '20:30'];

/**
 * Get scheduled times based on posts_per_day
 */
export function getScheduledTimes(postsPerDay: number): string[] {
  const count = Math.min(Math.max(postsPerDay, 1), 3);
  return DEFAULT_TIMES.slice(0, count);
}

/**
 * Generate queue_id
 * Format: ${bot_id}_${YYYYMMDD}_${HHMM}_${index}
 */
export function generateQueueId(
  botId: string,
  date: string,
  time: string,
  index: number
): string {
  const dateCompact = date.replace(/-/g, '');
  const timeCompact = formatTimeCompact(time);
  return `${botId}_${dateCompact}_${timeCompact}_${index}`;
}

/**
 * Get ISO timestamp in JST
 */
export function getJSTTimestamp(): string {
  const jst = getJSTNow();
  return jst.toISOString();
}
