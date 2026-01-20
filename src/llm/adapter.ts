// ============================================================
// LLM Adapter Interface
// ============================================================

import { LLMAdapter, LLMRequest } from '../utils/types';
import { AppConfig } from '../config';
import { createOpenAIAdapter } from './openai';
import { createAnthropicAdapter } from './anthropic';

/**
 * Create LLM adapter based on configuration
 */
export function createLLMAdapter(config: AppConfig): LLMAdapter {
  switch (config.llmProvider) {
    case 'openai':
      if (!config.openaiApiKey) {
        throw new Error('OpenAI API key is required');
      }
      return createOpenAIAdapter(config.openaiApiKey, config.openaiModel);
    
    case 'anthropic':
      if (!config.anthropicApiKey) {
        throw new Error('Anthropic API key is required');
      }
      return createAnthropicAdapter(config.anthropicApiKey, config.anthropicModel);
    
    default:
      throw new Error(`Unknown LLM provider: ${config.llmProvider}`);
  }
}

export { LLMAdapter, LLMRequest };
