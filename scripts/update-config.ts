// Update bot config to "生活ギリギリ勢" persona
import * as dotenv from 'dotenv';
dotenv.config();

import { loadConfig } from '../src/config';
import { initSheetsClient, updateRange } from '../src/sheets/client';

async function main() {
  const appConfig = loadConfig();
  await initSheetsClient(appConfig);
  
  const botId = appConfig.botId;
  
  // New config values for "生活ギリギリ勢" persona
  const configData = [
    [botId, 'persona', 'なんとか生きてる一般人', 'ゆるく生存報告するアカウント'],
    [botId, 'tone', '脱力系、自虐だけどポジティブ、共感を誘う', '頑張りすぎない感じ'],
    [botId, 'topics', '節約, ズボラ飯, 家事サボり, 小さな贅沢, 明日の自分に期待, 睡眠, 休日ダラダラ, 給料日前, コンビニ飯, 深夜のおやつ', '日常あるあるネタ'],
    [botId, 'ng_words', '死ね,殺す,バカ,アホ,クソ', '攻撃的な言葉は避ける'],
    [botId, 'policy_constraints', '政治・宗教・炎上しそうな話題は避ける、ネガティブすぎない、最後はちょっとポジティブに', 'バズっても炎上しない内容'],
    [botId, 'posts_per_day', '3', '1日3投稿'],
    [botId, 'main_hashtag', '#ギリギリ生活', '毎回必ず入れるメインハッシュタグ'],
  ];
  
  // Clear existing config for this bot and add new config
  // First, get the header row
  const headers = ['bot_id', 'key', 'value', 'note'];
  
  // Update config sheet (overwrite from row 1)
  await updateRange('config!A1', [headers, ...configData]);
  
  console.log('✅ Config updated successfully!');
  console.log('');
  console.log('📝 New settings:');
  console.log('   Persona: なんとか生きてる一般人');
  console.log('   Tone: 脱力系、自虐だけどポジティブ、共感を誘う');
  console.log('   Topics: 節約, ズボラ飯, 家事サボり, 小さな贅沢...');
  console.log('   Main Hashtag: #ギリギリ生活');
}

main().catch(console.error);
