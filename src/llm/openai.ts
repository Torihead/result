// ============================================================
// OpenAI LLM Adapter
// ============================================================

import OpenAI from 'openai';
import { LLMAdapter, LLMRequest } from '../utils/types';

export function createOpenAIAdapter(apiKey: string, model: string): LLMAdapter {
  const client = new OpenAI({ apiKey });
  
  return {
    async callLLM(request: LLMRequest): Promise<string> {
      const messages: OpenAI.Chat.ChatCompletionMessageParam[] = [
        { role: 'system', content: request.system },
        { role: 'user', content: request.user },
      ];
      
      // Add JSON schema hint if provided
      if (request.jsonSchemaHint) {
        messages[0].content += `\n\nExpected JSON format:\n${request.jsonSchemaHint}`;
      }
      
      const response = await client.chat.completions.create({
        model,
        messages,
        temperature: 0.7,
        max_tokens: 2000,
      });
      
      const content = response.choices[0]?.message?.content;
      if (!content) {
        throw new Error('Empty response from OpenAI');
      }
      
      return content.trim();
    },
  };
}
