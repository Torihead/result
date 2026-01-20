// ============================================================
// Planner Agent
// ============================================================

import { LLMAdapter, PlannerOutput, PipelineContext, PlanItem } from '../utils/types';
import { parseJSONWithRetry, SCHEMA_HINTS } from '../llm/json-parser';
import { getTopPerformingPosts, getCategories } from '../sheets/reference';
import { getRecentTexts } from '../sheets/history';

/**
 * Planner Agent: Generates post ideas with category, intent, and hook
 */
export async function runPlanner(
  llm: LLMAdapter,
  context: PipelineContext,
  count: number
): Promise<PlannerOutput> {
  const { config, referencePosts, historyPosts } = context;
  
  // Get top performing reference posts for inspiration
  const topPosts = getTopPerformingPosts(referencePosts, 10);
  const categories = getCategories(referencePosts);
  const recentTexts = getRecentTexts(historyPosts);
  
  // Build reference examples
  const referenceExamples = topPosts
    .map((p, i) => `${i + 1}. [${p.category}] ${p.text}`)
    .join('\n');
  
  // Build recent post summaries for avoidance
  const recentSummary = recentTexts
    .slice(0, 10)
    .map((t, i) => `${i + 1}. ${t.substring(0, 50)}...`)
    .join('\n');
  
  const systemPrompt = `You are a social media content planner for X (Twitter).
Your task is to generate ${count} unique post ideas.

Bot persona: ${config.persona || 'Engaging social media personality'}
Tone: ${config.tone || 'Friendly and informative'}
Topics: ${config.topics || 'General interest'}

Available categories: ${categories.length > 0 ? categories.join(', ') : 'General, Tips, Opinion, Question, Observation'}

Rules:
- Each idea must be distinct and avoid repetition with recent posts
- Each idea should have viral potential
- Consider different times of day for audience engagement
- Include variety in categories
- Output ONLY a valid JSON array, no markdown or explanation`;

  const userPrompt = `Generate ${count} X post ideas.

Reference high-performing posts for inspiration:
${referenceExamples || 'No reference posts available'}

Recent posts to AVOID similar topics:
${recentSummary || 'No recent posts'}

Output a JSON array with exactly ${count} items. Each item must have:
- category: string (topic category)
- intent: string (what you want to achieve with this post)
- hook: string (attention-grabbing opening)
- idea: string (core idea/message)`;

  console.log(`[Planner] Generating ${count} post ideas...`);
  
  const response = await llm.callLLM({
    system: systemPrompt,
    user: userPrompt,
    jsonSchemaHint: SCHEMA_HINTS.planner,
  });
  
  const plan = await parseJSONWithRetry<PlannerOutput>(
    response,
    llm,
    SCHEMA_HINTS.planner
  );
  
  // Validate plan structure
  if (!Array.isArray(plan)) {
    throw new Error('Planner output is not an array');
  }
  
  for (const item of plan) {
    if (!item.category || !item.intent || !item.hook || !item.idea) {
      throw new Error('Invalid plan item structure');
    }
  }
  
  console.log(`[Planner] Generated ${plan.length} ideas`);
  return plan;
}
