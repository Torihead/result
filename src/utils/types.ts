// ============================================================
// Type Definitions for X-Post Generator
// ============================================================

// Config from Google Sheets
export interface BotConfig {
  bot_id: string;
  persona?: string;
  tone?: string;
  topics?: string;
  ng_words?: string;
  policy_constraints?: string;
  posts_per_day?: number;
  main_hashtag?: string;
  [key: string]: string | number | undefined;
}

// Reference post from Sheets
export interface ReferencePost {
  bot_id: string;
  ref_id: string;
  url: string;
  text: string;
  category: string;
  likes: number;
  retweets: number;
  replies: number;
  impressions: number;
  engagement_rate: number;
  note: string;
  added_at: string;
}

// History post from Sheets
export interface HistoryPost {
  bot_id: string;
  posted_at: string;
  text: string;
  category: string;
  likes: number;
  impressions: number;
  note: string;
}

// Queue item to Sheets
export interface QueueItem {
  bot_id: string;
  queue_id: string;
  scheduled_date: string;
  scheduled_time: string;
  category: string;
  draft_text: string;
  status: 'draft' | 'approved' | 'rejected' | 'posted';
  guard_result: string;
  output_json: string;
  created_at: string;
}

// ============================================================
// Agent Output Types
// ============================================================

// Planner output
export interface PlanItem {
  category: string;
  intent: string;
  hook: string;
  idea: string;
}

export type PlannerOutput = PlanItem[];

// Writer output
export interface WriterOutput {
  draft_text: string;
  category: string;
  char_count: number;
  hashtags: string[];
}

// Guard output
export interface GuardOutput {
  decision: 'approved' | 'rewrite' | 'reject';
  final_text: string;
  reason: string;
  risk_flags: string[];
}

// ============================================================
// LLM Adapter Types
// ============================================================

export interface LLMRequest {
  system: string;
  user: string;
  jsonSchemaHint?: string;
}

export interface LLMAdapter {
  callLLM(request: LLMRequest): Promise<string>;
}

// ============================================================
// Pipeline Context
// ============================================================

export interface PipelineContext {
  botId: string;
  config: BotConfig;
  referencePosts: ReferencePost[];
  historyPosts: HistoryPost[];
  targetDates: string[];
  scheduledTimes: string[];
}

export interface GeneratedDraft {
  planItem: PlanItem;
  writerOutput: WriterOutput;
  guardOutput: GuardOutput;
  queueItem: QueueItem;
}
