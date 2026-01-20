// ============================================================
// X-Post Generator - Main Entry Point
// ============================================================

import { loadConfig } from './config';
import { initSheetsClient, initializeSheets } from './sheets/client';
import { readBotConfig } from './sheets/config';
import { readReferencePosts } from './sheets/reference';
import { readHistoryPosts } from './sheets/history';
import { createLLMAdapter } from './llm/adapter';
import { runPipeline, displaySummary } from './pipeline/runner';
import { PipelineContext } from './utils/types';
import { getTargetDates, getScheduledTimes } from './utils/date';

async function main(): Promise<void> {
  console.log('🚀 X-Post Generator Starting...\n');
  
  try {
    // Get bot_id from CLI args or env
    const botIdArg = process.argv[2];
    
    // Load app configuration
    const appConfig = loadConfig(botIdArg);
    console.log(`📦 Loaded configuration for bot: ${appConfig.botId}`);
    console.log(`🤖 LLM Provider: ${appConfig.llmProvider}`);
    
    // Initialize Google Sheets client
    console.log('📊 Connecting to Google Sheets...');
    await initSheetsClient(appConfig);
    
    // Initialize sheets if they don't exist
    console.log('📋 Checking/creating required sheets...');
    await initializeSheets(appConfig.botId);
    
    // Load bot-specific configuration from sheet
    console.log('⚙️ Loading bot configuration...');
    const botConfig = await readBotConfig(appConfig.botId);
    console.log(`   Persona: ${botConfig.persona || 'Default'}`);
    console.log(`   Tone: ${botConfig.tone || 'Default'}`);
    console.log(`   Topics: ${botConfig.topics || 'General'}`);
    
    // Load reference posts
    console.log('📚 Loading reference posts...');
    const referencePosts = await readReferencePosts(appConfig.botId);
    
    // Load history posts
    console.log('📜 Loading post history...');
    const historyPosts = await readHistoryPosts(appConfig.botId);
    
    // Create LLM adapter
    const llm = createLLMAdapter(appConfig);
    
    // Build pipeline context
    const postsPerDay = botConfig.posts_per_day || appConfig.postsPerDay;
    const context: PipelineContext = {
      botId: appConfig.botId,
      config: { ...botConfig, posts_per_day: postsPerDay },
      referencePosts,
      historyPosts,
      targetDates: getTargetDates(),
      scheduledTimes: getScheduledTimes(postsPerDay),
    };
    
    // Run the pipeline
    const results = await runPipeline(llm, context);
    
    // Display summary
    displaySummary(results);
    
    console.log('✨ Generation complete!');
    
  } catch (error) {
    console.error('\n❌ Error:', error);
    process.exit(1);
  }
}

// Run main
main();
