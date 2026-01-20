// ============================================================
// Pipeline Runner - Orchestrates the multi-agent pipeline
// ============================================================

import { 
  LLMAdapter, 
  PipelineContext, 
  QueueItem, 
  GeneratedDraft,
  PlanItem,
} from '../utils/types';
import { generateQueueId, getTargetDates, getScheduledTimes } from '../utils/date';
import { runPlanner } from '../agents/planner';
import { runWriter, rewriteDraft } from '../agents/writer';
import { runGuard } from '../agents/guard';
import { getExistingQueueIds, addQueueItems, createQueueItem } from '../sheets/queue';

const MAX_REWRITE_ATTEMPTS = 2;

/**
 * Run the full pipeline: Planner -> Writer -> Guard -> Queue
 */
export async function runPipeline(
  llm: LLMAdapter,
  context: PipelineContext
): Promise<GeneratedDraft[]> {
  const { botId, config } = context;
  
  // Get target dates (today and tomorrow)
  const targetDates = getTargetDates();
  const postsPerDay = config.posts_per_day || 3;
  const scheduledTimes = getScheduledTimes(postsPerDay);
  
  console.log('\n========================================');
  console.log(`[Pipeline] Starting for bot: ${botId}`);
  console.log(`[Pipeline] Target dates: ${targetDates.join(', ')}`);
  console.log(`[Pipeline] Scheduled times: ${scheduledTimes.join(', ')}`);
  console.log('========================================\n');
  
  // Get existing queue IDs to avoid duplicates
  const existingIds = await getExistingQueueIds(botId);
  console.log(`[Pipeline] Existing queue items: ${existingIds.size}`);
  
  // Build list of slots to fill
  const slotsToFill: Array<{ date: string; time: string; index: number }> = [];
  
  for (const date of targetDates) {
    for (let i = 0; i < scheduledTimes.length; i++) {
      const time = scheduledTimes[i];
      const queueId = generateQueueId(botId, date, time, i);
      
      if (!existingIds.has(queueId)) {
        slotsToFill.push({ date, time, index: i });
      }
    }
  }
  
  if (slotsToFill.length === 0) {
    console.log('[Pipeline] All slots are already filled. Nothing to do.');
    return [];
  }
  
  console.log(`[Pipeline] Slots to fill: ${slotsToFill.length}`);
  
  // Run Planner to get ideas
  const planCount = slotsToFill.length;
  const plan = await runPlanner(llm, context, planCount);
  
  // Process each plan item through Writer and Guard
  const results: GeneratedDraft[] = [];
  const queueItems: QueueItem[] = [];
  
  for (let i = 0; i < Math.min(plan.length, slotsToFill.length); i++) {
    const planItem = plan[i];
    const slot = slotsToFill[i];
    const queueId = generateQueueId(botId, slot.date, slot.time, slot.index);
    
    console.log(`\n--- Processing slot ${i + 1}/${slotsToFill.length} ---`);
    console.log(`[Pipeline] Queue ID: ${queueId}`);
    console.log(`[Pipeline] Idea: ${planItem.idea}`);
    
    try {
      // Run Writer
      let writerOutput = await runWriter(llm, context, planItem);
      
      // Run Guard
      let guardOutput = await runGuard(
        llm, 
        context, 
        writerOutput.draft_text, 
        writerOutput.category
      );
      
      // Handle rewrite requests
      let rewriteAttempts = 0;
      while (
        guardOutput.decision === 'rewrite' && 
        rewriteAttempts < MAX_REWRITE_ATTEMPTS
      ) {
        console.log(`[Pipeline] Rewrite attempt ${rewriteAttempts + 1}/${MAX_REWRITE_ATTEMPTS}`);
        
        writerOutput = await rewriteDraft(
          llm,
          context,
          writerOutput.draft_text,
          guardOutput.reason,
          guardOutput.risk_flags
        );
        
        guardOutput = await runGuard(
          llm,
          context,
          writerOutput.draft_text,
          writerOutput.category
        );
        
        rewriteAttempts++;
      }
      
      // Determine final status
      let status: QueueItem['status'] = 'draft';
      if (guardOutput.decision === 'reject') {
        status = 'rejected';
      } else if (guardOutput.decision === 'approved') {
        status = 'draft'; // Still draft until manually reviewed
      }
      
      // Create output JSON
      const outputJson = JSON.stringify({
        plan: planItem,
        writer: writerOutput,
        guard: guardOutput,
      });
      
      // Create queue item
      const queueItem = createQueueItem({
        botId,
        queueId,
        scheduledDate: slot.date,
        scheduledTime: slot.time,
        category: planItem.category,
        draftText: guardOutput.final_text || writerOutput.draft_text,
        status,
        guardResult: JSON.stringify(guardOutput),
        outputJson,
      });
      
      queueItems.push(queueItem);
      results.push({
        planItem,
        writerOutput,
        guardOutput,
        queueItem,
      });
      
      console.log(`[Pipeline] Slot ${i + 1} completed with status: ${status}`);
      
    } catch (error) {
      console.error(`[Pipeline] Error processing slot ${i + 1}:`, error);
      
      // Create rejected queue item for error case
      const queueItem = createQueueItem({
        botId,
        queueId,
        scheduledDate: slot.date,
        scheduledTime: slot.time,
        category: planItem.category,
        draftText: '',
        status: 'rejected',
        guardResult: JSON.stringify({ 
          decision: 'reject', 
          reason: `Error: ${error instanceof Error ? error.message : 'Unknown'}`,
          risk_flags: ['processing_error'],
        }),
        outputJson: JSON.stringify({ error: String(error) }),
      });
      
      queueItems.push(queueItem);
    }
  }
  
  // Save all queue items to sheet
  if (queueItems.length > 0) {
    await addQueueItems(queueItems);
  }
  
  console.log('\n========================================');
  console.log(`[Pipeline] Completed. Generated ${results.length} drafts.`);
  console.log('========================================\n');
  
  return results;
}

/**
 * Display summary of generated drafts
 */
export function displaySummary(results: GeneratedDraft[]): void {
  console.log('\n📊 Generation Summary:');
  console.log('─'.repeat(50));
  
  const approved = results.filter((r) => r.guardOutput.decision === 'approved');
  const rejected = results.filter((r) => r.guardOutput.decision === 'reject');
  const rewritten = results.filter((r) => r.guardOutput.decision === 'rewrite');
  
  console.log(`✅ Approved: ${approved.length}`);
  console.log(`❌ Rejected: ${rejected.length}`);
  console.log(`📝 Needs review: ${rewritten.length}`);
  
  console.log('\n📝 Generated Drafts:');
  console.log('─'.repeat(50));
  
  for (const result of results) {
    const status = result.guardOutput.decision === 'approved' ? '✅' : 
                   result.guardOutput.decision === 'reject' ? '❌' : '⚠️';
    
    console.log(`\n${status} [${result.queueItem.scheduled_date} ${result.queueItem.scheduled_time}]`);
    console.log(`   Category: ${result.planItem.category}`);
    console.log(`   Text: ${result.queueItem.draft_text.substring(0, 100)}...`);
    
    if (result.guardOutput.risk_flags.length > 0) {
      console.log(`   Flags: ${result.guardOutput.risk_flags.join(', ')}`);
    }
  }
  
  console.log('\n' + '─'.repeat(50));
}
