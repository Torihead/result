// ============================================================
// JSON Parser with Retry Logic
// ============================================================

import { LLMAdapter } from '../utils/types';

/**
 * Extract JSON from LLM response (handles markdown code blocks)
 */
function extractJSON(text: string): string {
  // Try to extract JSON from markdown code block
  const codeBlockMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (codeBlockMatch) {
    return codeBlockMatch[1].trim();
  }
  
  // Try to find JSON array or object
  const jsonMatch = text.match(/(\[[\s\S]*\]|\{[\s\S]*\})/);
  if (jsonMatch) {
    return jsonMatch[1].trim();
  }
  
  return text.trim();
}

/**
 * Parse JSON with validation
 */
export function parseJSON<T>(text: string): T {
  const jsonStr = extractJSON(text);
  return JSON.parse(jsonStr) as T;
}

/**
 * Parse JSON with retry on failure
 * If parsing fails, asks LLM to fix the JSON
 */
export async function parseJSONWithRetry<T>(
  text: string,
  llm: LLMAdapter,
  schemaHint: string,
  maxRetries: number = 1
): Promise<T> {
  // First attempt
  try {
    return parseJSON<T>(text);
  } catch (firstError) {
    if (maxRetries <= 0) {
      throw firstError;
    }
    
    console.log('[JSON Parser] Parse failed, attempting fix...');
    
    // Ask LLM to fix the JSON
    const fixPrompt = `The following text should be valid JSON but failed to parse.
Error: ${firstError instanceof Error ? firstError.message : 'Unknown error'}

Original text:
${text}

Expected format:
${schemaHint}

Please output ONLY the corrected valid JSON, no explanation or markdown.`;

    const fixedResponse = await llm.callLLM({
      system: 'You are a JSON fixer. Output ONLY valid JSON, nothing else.',
      user: fixPrompt,
    });
    
    // Second attempt
    try {
      return parseJSON<T>(fixedResponse);
    } catch (secondError) {
      console.error('[JSON Parser] Fix attempt also failed');
      throw new Error(
        `JSON parse failed after retry. Original: ${text.substring(0, 200)}...`
      );
    }
  }
}

/**
 * Schema hints for each agent output
 */
export const SCHEMA_HINTS = {
  planner: `[
  {
    "category": "string (topic category)",
    "intent": "string (what you want to achieve)",
    "hook": "string (attention grabber)",
    "idea": "string (core idea)"
  }
]`,
  
  writer: `{
  "draft_text": "string (the tweet, max 280 chars)",
  "category": "string (category from plan)",
  "char_count": number
}`,
  
  guard: `{
  "decision": "approved" | "rewrite" | "reject",
  "final_text": "string (final or rewritten text)",
  "reason": "string (explanation)",
  "risk_flags": ["string"]
}`,
};
