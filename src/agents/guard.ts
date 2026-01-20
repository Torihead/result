// ============================================================
// Guard Agent
// ============================================================

import { LLMAdapter, GuardOutput, PipelineContext } from '../utils/types';
import { parseJSONWithRetry, SCHEMA_HINTS } from '../llm/json-parser';
import { getRecentTexts, isSimilarToRecent } from '../sheets/history';

const MAX_CHARS = 280;

/**
 * Guard Agent: Checks NG words, policy, repetition, and tone
 */
export async function runGuard(
  llm: LLMAdapter,
  context: PipelineContext,
  draftText: string,
  category: string
): Promise<GuardOutput> {
  const { config, historyPosts } = context;
  
  // Get NG words and policy constraints from config
  const ngWords = config.ng_words
    ? String(config.ng_words).split(',').map((w) => w.trim())
    : [];
  const policyConstraints = config.policy_constraints || '';
  
  // Get recent texts for repetition check
  const recentTexts = getRecentTexts(historyPosts);
  
  // Pre-check: local NG word check
  const localNgCheck = checkNgWords(draftText, ngWords);
  
  // Pre-check: local repetition check
  const localRepetitionCheck = isSimilarToRecent(draftText, recentTexts, 0.7);
  
  // Pre-check: character limit
  const charLimitOk = draftText.length <= MAX_CHARS;
  
  const systemPrompt = `You are a content safety guard for X (Twitter) posts.
Your task is to review a draft tweet and ensure it meets quality and safety standards.

Bot persona: ${config.persona || 'Engaging social media personality'}
Expected tone: ${config.tone || 'Friendly and informative'}

NG Words (must not contain): ${ngWords.join(', ') || 'None specified'}

Policy constraints:
${policyConstraints || 'Standard social media best practices'}

Pre-check results (already verified):
- NG word detected locally: ${localNgCheck.found ? `YES - "${localNgCheck.word}"` : 'NO'}
- Too similar to recent posts: ${localRepetitionCheck ? 'YES' : 'NO'}
- Within ${MAX_CHARS} char limit: ${charLimitOk ? 'YES' : `NO (${draftText.length} chars)`}

Rules:
- If any pre-check failed, you MUST reject or request rewrite
- Check for inappropriate content, spam signals, or policy violations
- Verify the tone matches the persona
- Ensure the post is engaging and professional
- Output ONLY a valid JSON object, no markdown or explanation`;

  const userPrompt = `Review this draft tweet:

"${draftText}"

Category: ${category}

Recent posts (for repetition context):
${recentTexts.slice(0, 5).map((t) => `- ${t.substring(0, 80)}...`).join('\n') || 'No recent posts'}

Output a JSON object with:
- decision: "approved" (safe to post) | "rewrite" (needs changes) | "reject" (cannot be fixed)
- final_text: string (original if approved, improved version if rewrite, empty if reject)
- reason: string (explanation of decision)
- risk_flags: array of strings (any issues found, empty if none)`;

  console.log('[Guard] Reviewing draft...');
  
  const response = await llm.callLLM({
    system: systemPrompt,
    user: userPrompt,
    jsonSchemaHint: SCHEMA_HINTS.guard,
  });
  
  const output = await parseJSONWithRetry<GuardOutput>(
    response,
    llm,
    SCHEMA_HINTS.guard
  );
  
  // Validate output
  if (!['approved', 'rewrite', 'reject'].includes(output.decision)) {
    throw new Error(`Invalid guard decision: ${output.decision}`);
  }
  
  // Ensure risk_flags is array
  if (!Array.isArray(output.risk_flags)) {
    output.risk_flags = [];
  }
  
  // Force rejection if local checks failed and LLM missed it
  if (localNgCheck.found && output.decision === 'approved') {
    output.decision = 'reject';
    output.reason = `Contains NG word: "${localNgCheck.word}"`;
    output.risk_flags.push('ng_word_detected');
    output.final_text = '';
  }
  
  if (localRepetitionCheck && output.decision === 'approved') {
    output.decision = 'rewrite';
    output.reason = 'Too similar to recent posts';
    output.risk_flags.push('repetition_detected');
  }
  
  if (!charLimitOk && output.decision === 'approved') {
    output.decision = 'rewrite';
    output.reason = `Exceeds ${MAX_CHARS} character limit`;
    output.risk_flags.push('char_limit_exceeded');
  }
  
  console.log(`[Guard] Decision: ${output.decision}`);
  if (output.risk_flags.length > 0) {
    console.log(`[Guard] Risk flags: ${output.risk_flags.join(', ')}`);
  }
  
  return output;
}

/**
 * Local NG word check
 */
function checkNgWords(
  text: string,
  ngWords: string[]
): { found: boolean; word?: string } {
  const lowerText = text.toLowerCase();
  
  for (const word of ngWords) {
    if (word && lowerText.includes(word.toLowerCase())) {
      return { found: true, word };
    }
  }
  
  return { found: false };
}
