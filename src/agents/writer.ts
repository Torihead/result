// ============================================================
// Writer Agent
// ============================================================

import { LLMAdapter, PlanItem, WriterOutput, PipelineContext } from '../utils/types';
import { parseJSONWithRetry, SCHEMA_HINTS } from '../llm/json-parser';

const MAX_CHARS = 140;

/**
 * Writer Agent: Converts plan items into 140-char drafts with hashtags
 */
export async function runWriter(
  llm: LLMAdapter,
  context: PipelineContext,
  planItem: PlanItem
): Promise<WriterOutput> {
  const { config, referencePosts } = context;
  
  // Get reference examples for this category
  const categoryRefs = referencePosts
    .filter((p) => p.category === planItem.category)
    .slice(0, 3);
  
  const referenceExamples = categoryRefs
    .map((p) => `- ${p.text}`)
    .join('\n');
  
  // Get main hashtag from config (default: none)
  const mainHashtag = config.main_hashtag || '';
  
  const systemPrompt = `You are a skilled X (Twitter) copywriter.
Your task is to write a compelling tweet based on the given idea.

Bot persona: ${config.persona || 'Engaging social media personality'}
Tone: ${config.tone || 'Friendly and informative'}

Rules:
- Maximum ${MAX_CHARS} characters for the main text (CRITICAL!)
- Make it engaging and shareable
- Include a hook at the beginning
- Be concise but impactful
- Use appropriate emojis sparingly if it fits the tone
- Suggest relevant hashtags that match the content
- Output ONLY a valid JSON object, no markdown or explanation`;

  const userPrompt = `Write a tweet based on this plan:

Category: ${planItem.category}
Intent: ${planItem.intent}
Hook: ${planItem.hook}
Idea: ${planItem.idea}

${referenceExamples ? `Reference examples in this category:\n${referenceExamples}` : ''}

${mainHashtag ? `Main hashtag to always include: ${mainHashtag}` : ''}

Output a JSON object with:
- draft_text: string (the tweet text WITHOUT hashtags, MUST be under ${MAX_CHARS} characters)
- category: string (same as input category)
- char_count: number (character count of draft_text only)
- hashtags: string[] (2-4 relevant hashtags${mainHashtag ? `, always include "${mainHashtag}" first` : ''})`;

  console.log(`[Writer] Writing draft for: ${planItem.idea.substring(0, 50)}...`);
  
  const response = await llm.callLLM({
    system: systemPrompt,
    user: userPrompt,
    jsonSchemaHint: SCHEMA_HINTS.writer,
  });
  
  const output = await parseJSONWithRetry<WriterOutput>(
    response,
    llm,
    SCHEMA_HINTS.writer
  );
  
  // Validate output
  if (!output.draft_text) {
    throw new Error('Writer output missing draft_text');
  }
  
  // Ensure hashtags array exists
  if (!output.hashtags) {
    output.hashtags = mainHashtag ? [mainHashtag] : [];
  }
  
  // Ensure main hashtag is included first
  if (mainHashtag && !output.hashtags.includes(mainHashtag)) {
    output.hashtags.unshift(mainHashtag);
  }
  
  // Format hashtags (ensure they start with #)
  output.hashtags = output.hashtags.map(tag => 
    tag.startsWith('#') ? tag : `#${tag}`
  ).slice(0, 4); // Max 4 hashtags
  
  // Ensure char count is accurate (text only, not hashtags)
  output.char_count = output.draft_text.length;
  
  // Warn if over limit (but don't fail)
  if (output.char_count > MAX_CHARS) {
    console.warn(`[Writer] Warning: Draft exceeds ${MAX_CHARS} chars (${output.char_count})`);
  }
  
  console.log(`[Writer] Draft complete: ${output.char_count} chars, hashtags: ${output.hashtags.join(' ')}`);
  return output;
}

/**
 * Rewrite a draft based on guard feedback
 */
export async function rewriteDraft(
  llm: LLMAdapter,
  context: PipelineContext,
  originalDraft: string,
  feedback: string,
  riskFlags: string[],
  originalHashtags: string[] = []
): Promise<WriterOutput> {
  const { config } = context;
  const mainHashtag = config.main_hashtag || '';
  
  const systemPrompt = `You are a skilled X (Twitter) copywriter.
Your task is to rewrite a tweet based on feedback.

Bot persona: ${config.persona || 'Engaging social media personality'}
Tone: ${config.tone || 'Friendly and informative'}

Rules:
- Maximum ${MAX_CHARS} characters for the main text (CRITICAL!)
- Address all the issues mentioned in the feedback
- Maintain the original intent and message
- Suggest relevant hashtags that match the content
- Output ONLY a valid JSON object, no markdown or explanation`;

  const userPrompt = `Rewrite this tweet based on feedback:

Original draft:
${originalDraft}

Original hashtags: ${originalHashtags.join(' ') || 'None'}

Issues to fix:
- ${feedback}
- Risk flags: ${riskFlags.join(', ') || 'None'}

${mainHashtag ? `Main hashtag to always include: ${mainHashtag}` : ''}

Output a JSON object with:
- draft_text: string (the rewritten tweet WITHOUT hashtags, MUST be under ${MAX_CHARS} characters)
- category: string (keep the same category)
- char_count: number (character count of draft_text only)
- hashtags: string[] (2-4 relevant hashtags${mainHashtag ? `, always include "${mainHashtag}" first` : ''})`;

  console.log('[Writer] Rewriting draft based on feedback...');
  
  const response = await llm.callLLM({
    system: systemPrompt,
    user: userPrompt,
    jsonSchemaHint: SCHEMA_HINTS.writer,
  });
  
  const output = await parseJSONWithRetry<WriterOutput>(
    response,
    llm,
    SCHEMA_HINTS.writer
  );
  
  // Ensure hashtags array exists
  if (!output.hashtags) {
    output.hashtags = mainHashtag ? [mainHashtag] : [];
  }
  
  // Ensure main hashtag is included first
  if (mainHashtag && !output.hashtags.includes(mainHashtag)) {
    output.hashtags.unshift(mainHashtag);
  }
  
  // Format hashtags
  output.hashtags = output.hashtags.map(tag => 
    tag.startsWith('#') ? tag : `#${tag}`
  ).slice(0, 4);
  
  output.char_count = output.draft_text.length;
  
  console.log(`[Writer] Rewrite complete: ${output.char_count} chars, hashtags: ${output.hashtags.join(' ')}`);
  return output;
}
