// ============================================================
// Environment Configuration
// ============================================================

import * as dotenv from 'dotenv';
import * as fs from 'fs';
import * as path from 'path';

// Load environment variables
dotenv.config();

export interface AppConfig {
  // Google Sheets
  googleSheetsId: string;
  googleServiceAccountEmail: string;
  googlePrivateKey: string;
  
  // LLM
  llmProvider: 'openai' | 'anthropic';
  openaiApiKey?: string;
  openaiModel: string;
  anthropicApiKey?: string;
  anthropicModel: string;
  
  // Bot
  botId: string;
  postsPerDay: number;
}

interface GoogleServiceAccountCredentials {
  client_email: string;
  private_key: string;
}

function getEnvOrThrow(key: string): string {
  const value = process.env[key];
  if (!value) {
    throw new Error(`Missing required environment variable: ${key}`);
  }
  return value;
}

function getEnvOrDefault(key: string, defaultValue: string): string {
  return process.env[key] || defaultValue;
}

function loadGoogleCredentials(): { email: string; privateKey: string } {
  // First, try JSON file path
  const credentialsPath = process.env.GOOGLE_SHEETS_CREDENTIALS_PATH;
  if (credentialsPath) {
    const fullPath = path.resolve(credentialsPath);
    if (fs.existsSync(fullPath)) {
      const credentialsJson = fs.readFileSync(fullPath, 'utf-8');
      const credentials: GoogleServiceAccountCredentials = JSON.parse(credentialsJson);
      return {
        email: credentials.client_email,
        privateKey: credentials.private_key,
      };
    }
  }
  
  // Fallback to direct environment variables
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const privateKey = process.env.GOOGLE_PRIVATE_KEY;
  
  if (email && privateKey) {
    return {
      email,
      privateKey: privateKey.replace(/\\n/g, '\n'),
    };
  }
  
  throw new Error(
    'Google credentials not found. Please set either:\n' +
    '  - GOOGLE_SHEETS_CREDENTIALS_PATH (path to service account JSON file)\n' +
    '  - Or both GOOGLE_SERVICE_ACCOUNT_EMAIL and GOOGLE_PRIVATE_KEY'
  );
}

export function loadConfig(botIdOverride?: string): AppConfig {
  const llmProvider = getEnvOrDefault('LLM_PROVIDER', 'openai') as 'openai' | 'anthropic';
  
  // Validate LLM API key based on provider
  if (llmProvider === 'openai' && !process.env.OPENAI_API_KEY) {
    throw new Error('OPENAI_API_KEY is required when LLM_PROVIDER is openai');
  }
  if (llmProvider === 'anthropic' && !process.env.ANTHROPIC_API_KEY) {
    throw new Error('ANTHROPIC_API_KEY is required when LLM_PROVIDER is anthropic');
  }
  
  const googleCredentials = loadGoogleCredentials();
  
  return {
    googleSheetsId: getEnvOrThrow('GOOGLE_SHEETS_ID'),
    googleServiceAccountEmail: googleCredentials.email,
    googlePrivateKey: googleCredentials.privateKey,
    
    llmProvider,
    openaiApiKey: process.env.OPENAI_API_KEY,
    openaiModel: getEnvOrDefault('OPENAI_MODEL', 'gpt-4o'),
    anthropicApiKey: process.env.ANTHROPIC_API_KEY,
    anthropicModel: getEnvOrDefault('ANTHROPIC_MODEL', 'claude-3-5-sonnet-20241022'),
    
    botId: botIdOverride || getEnvOrDefault('BOT_ID', 'default_bot'),
    postsPerDay: parseInt(getEnvOrDefault('POSTS_PER_DAY', '3'), 10),
  };
}
