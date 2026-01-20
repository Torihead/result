// ============================================================
// Anthropic LLM Adapter
// ============================================================

import { LLMAdapter, LLMRequest } from '../utils/types';

interface AnthropicMessage {
  role: 'user' | 'assistant';
  content: string;
}

interface AnthropicResponse {
  content: Array<{ type: string; text: string }>;
}

export function createAnthropicAdapter(apiKey: string, model: string): LLMAdapter {
  return {
    async callLLM(request: LLMRequest): Promise<string> {
      let systemPrompt = request.system;
      
      // Add JSON schema hint if provided
      if (request.jsonSchemaHint) {
        systemPrompt += `\n\nExpected JSON format:\n${request.jsonSchemaHint}`;
      }
      
      const messages: AnthropicMessage[] = [
        { role: 'user', content: request.user },
      ];
      
      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
        },
        body: JSON.stringify({
          model,
          max_tokens: 2000,
          system: systemPrompt,
          messages,
        }),
      });
      
      if (!response.ok) {
        const error = await response.text();
        throw new Error(`Anthropic API error: ${response.status} - ${error}`);
      }
      
      const data = (await response.json()) as AnthropicResponse;
      const content = data.content[0]?.text;
      
      if (!content) {
        throw new Error('Empty response from Anthropic');
      }
      
      return content.trim();
    },
  };
}
