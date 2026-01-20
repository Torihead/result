// ============================================================
// Google Sheets Client
// ============================================================

import { google, sheets_v4 } from 'googleapis';
import { AppConfig } from '../config';

let sheetsClient: sheets_v4.Sheets | null = null;
let spreadsheetId: string = '';

/**
 * Initialize Google Sheets client
 */
export async function initSheetsClient(config: AppConfig): Promise<sheets_v4.Sheets> {
  if (sheetsClient) {
    return sheetsClient;
  }
  
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: config.googleServiceAccountEmail,
      private_key: config.googlePrivateKey,
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  
  sheetsClient = google.sheets({ version: 'v4', auth });
  spreadsheetId = config.googleSheetsId;
  
  return sheetsClient;
}

/**
 * Get spreadsheet ID
 */
export function getSpreadsheetId(): string {
  return spreadsheetId;
}

/**
 * Read values from a sheet range
 */
export async function readRange(range: string): Promise<string[][]> {
  if (!sheetsClient) {
    throw new Error('Sheets client not initialized');
  }
  
  const response = await sheetsClient.spreadsheets.values.get({
    spreadsheetId,
    range,
  });
  
  return (response.data.values as string[][]) || [];
}

/**
 * Append rows to a sheet
 */
export async function appendRows(
  sheetName: string,
  rows: (string | number)[][]
): Promise<void> {
  if (!sheetsClient) {
    throw new Error('Sheets client not initialized');
  }
  
  await sheetsClient.spreadsheets.values.append({
    spreadsheetId,
    range: `${sheetName}!A:Z`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: rows,
    },
  });
}

/**
 * Update a specific range
 */
export async function updateRange(
  range: string,
  values: (string | number)[][]
): Promise<void> {
  if (!sheetsClient) {
    throw new Error('Sheets client not initialized');
  }
  
  await sheetsClient.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values,
    },
  });
}

/**
 * Convert row array to object using headers
 */
export function rowToObject<T extends Record<string, unknown>>(
  headers: string[],
  row: string[]
): T {
  const obj: Record<string, string> = {};
  headers.forEach((header, index) => {
    obj[header] = row[index] || '';
  });
  return obj as T;
}

/**
 * Convert object to row array using headers
 */
export function objectToRow(
  headers: string[],
  obj: Record<string, unknown>
): (string | number)[] {
  return headers.map((header) => {
    const value = obj[header];
    if (value === undefined || value === null) return '';
    if (typeof value === 'number') return value;
    return String(value);
  });
}

/**
 * Check if a sheet exists
 */
export async function sheetExists(sheetName: string): Promise<boolean> {
  if (!sheetsClient) {
    throw new Error('Sheets client not initialized');
  }
  
  const response = await sheetsClient.spreadsheets.get({
    spreadsheetId,
  });
  
  const sheets = response.data.sheets || [];
  return sheets.some((sheet) => sheet.properties?.title === sheetName);
}

/**
 * Create a new sheet with headers
 */
export async function createSheet(
  sheetName: string,
  headers: string[]
): Promise<void> {
  if (!sheetsClient) {
    throw new Error('Sheets client not initialized');
  }
  
  // Add new sheet
  await sheetsClient.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          addSheet: {
            properties: {
              title: sheetName,
            },
          },
        },
      ],
    },
  });
  
  // Add headers
  await updateRange(`${sheetName}!A1`, [headers]);
}

/**
 * Initialize required sheets if they don't exist
 */
export async function initializeSheets(botId: string): Promise<void> {
  const sheetsConfig = [
    {
      name: 'config',
      headers: ['bot_id', 'key', 'value', 'note'],
      initialData: [
        [botId, 'persona', 'テック系インフルエンサー', 'ボットのペルソナ'],
        [botId, 'tone', 'フレンドリーで情報豊富', '投稿のトーン'],
        [botId, 'topics', 'AI, プログラミング, 生産性', '対象トピック'],
        [botId, 'ng_words', '死ね,殺す,バカ', '禁止ワード（カンマ区切り）'],
        [botId, 'policy_constraints', '政治的な内容は避ける', 'ポリシー制約'],
        [botId, 'posts_per_day', '3', '1日の投稿数（1-3）'],
        [botId, 'main_hashtag', '#MyHashtag', '毎回入れるメインハッシュタグ'],
      ],
    },
    {
      name: 'reference_posts',
      headers: ['bot_id', 'ref_id', 'url', 'text', 'category', 'likes', 'retweets', 'replies', 'impressions', 'engagement_rate', 'note', 'added_at'],
      initialData: [
        [botId, 'ref_001', 'https://x.com/example', 'サンプル投稿テキストです。参考にしてください。', 'Tips', '100', '50', '10', '5000', '0.02', 'サンプル', new Date().toISOString().split('T')[0]],
      ],
    },
    {
      name: 'history',
      headers: ['bot_id', 'posted_at', 'text', 'category', 'likes', 'impressions', 'note'],
      initialData: [],
    },
    {
      name: 'queue',
      headers: ['bot_id', 'queue_id', 'scheduled_date', 'scheduled_time', 'category', 'draft_text', 'status', 'guard_result', 'output_json', 'created_at'],
      initialData: [],
    },
  ];
  
  for (const sheet of sheetsConfig) {
    const exists = await sheetExists(sheet.name);
    if (!exists) {
      console.log(`  Creating sheet: ${sheet.name}...`);
      await createSheet(sheet.name, sheet.headers);
      
      if (sheet.initialData.length > 0) {
        await appendRows(sheet.name, sheet.initialData);
      }
    } else {
      // Update headers for existing sheets
      await updateRange(`${sheet.name}!A1`, [sheet.headers]);
    }
  }
  
  // Create guide sheet if it doesn't exist
  const guideExists = await sheetExists('📖_使い方ガイド');
  if (!guideExists) {
    console.log('  Creating sheet: 📖_使い方ガイド...');
    await createGuideSheet();
  }
}

/**
 * Create a guide sheet with instructions for all sheets
 */
async function createGuideSheet(): Promise<void> {
  const sheetName = '📖_使い方ガイド';
  
  // Create the sheet
  await sheetsClient!.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          addSheet: {
            properties: {
              title: sheetName,
              index: 0, // First sheet
            },
          },
        },
      ],
    },
  });
  
  const guideContent = [
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['📖 X-Post Generator 使い方ガイド'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['このスプレッドシートは、X（Twitter）投稿を自動生成するツールのデータベースです。'],
    ['各シートの役割と入力方法を以下に説明します。'],
    [''],
    ['───────────────────────────────────────────────────────────────────────────────'],
    ['📌 config シート - ボットの基本設定'],
    ['───────────────────────────────────────────────────────────────────────────────'],
    [''],
    ['【列の説明】'],
    ['  bot_id    : ボットの識別子（.envのBOT_IDと一致させる）'],
    ['  key       : 設定項目名（下記参照）'],
    ['  value     : 設定値'],
    ['  note      : メモ（任意）'],
    [''],
    ['【設定項目（key）一覧】'],
    ['  persona           : ボットのキャラクター設定（例：テック系インフルエンサー、料理研究家）'],
    ['  tone              : 投稿のトーン（例：フレンドリー、プロフェッショナル、カジュアル）'],
    ['  topics            : 扱うトピック（カンマ区切り、例：AI, プログラミング, 生産性）'],
    ['  ng_words          : 禁止ワード（カンマ区切り、例：死ね,殺す,バカ）'],
    ['  policy_constraints: ポリシー制約（例：政治的な内容は避ける）'],
    ['  posts_per_day     : 1日の投稿数（1〜3）'],
    ['  main_hashtag      : ★毎回必ず入れるメインハッシュタグ（例：#ギリギリ生活）'],
    [''],
    ['───────────────────────────────────────────────────────────────────────────────'],
    ['📚 reference_posts シート - 参考投稿（お手本）'],
    ['───────────────────────────────────────────────────────────────────────────────'],
    [''],
    ['【目的】'],
    ['  過去にバズった投稿や、理想的な文体の投稿を登録しておくと、'],
    ['  Writerエージェントが同じスタイルで新しい投稿を生成します。'],
    [''],
    ['【列の説明】'],
    ['  bot_id          : ボットID'],
    ['  ref_id          : 参照ID（任意、例：ref_001）'],
    ['  url             : 元の投稿URL（任意）'],
    ['  text            : ★投稿本文（必須）- これがお手本として使われます'],
    ['  category        : ★カテゴリ（必須）- 下記参照'],
    ['  likes           : いいね数（参考値、任意）'],
    ['  retweets        : リツイート数（参考値、任意）'],
    ['  replies         : コメント/リプライ数（参考値、任意）'],
    ['  impressions     : インプレッション数（参考値、任意）'],
    ['  engagement_rate : エンゲージメント率（参考値、任意）'],
    ['  note            : メモ（任意）'],
    ['  added_at        : 追加日（任意）'],
    [''],
    ['【カテゴリ一覧】※統一して使用してください'],
    ['  Tips      : ノウハウ・ハウツー系（〜する方法、〜のコツ）'],
    ['  Insight   : 気づき・考察系（〜だと気づいた、〜について思うこと）'],
    ['  Question  : 質問・投げかけ系（みんなはどう思う？、〜ってどうしてる？）'],
    ['  News      : ニュース・情報共有系（〜がリリース、〜が話題）'],
    ['  Personal  : 個人的なエピソード系（今日〜した、〜を試してみた）'],
    ['  Promotion : 宣伝・告知系（新サービス、イベント告知）'],
    ['  Thread    : スレッド用（長文を分割する場合）'],
    [''],
    ['【おすすめ】'],
    ['  ・各カテゴリに2〜5件程度登録すると効果的'],
    ['  ・自分の過去のバズ投稿を優先的に登録'],
    ['  ・参考にしたい他アカウントの投稿も可（文体参考用）'],
    [''],
    ['───────────────────────────────────────────────────────────────────────────────'],
    ['📜 history シート - 投稿履歴'],
    ['───────────────────────────────────────────────────────────────────────────────'],
    [''],
    ['【目的】'],
    ['  過去に投稿した内容を記録し、重複を避けるために使用されます。'],
    ['  ※手動で入力するか、投稿後に記録してください。'],
    [''],
    ['【列の説明】'],
    ['  bot_id      : ボットID'],
    ['  posted_at   : 投稿日時（例：2026-01-19 08:10）'],
    ['  text        : 投稿本文'],
    ['  category    : カテゴリ'],
    ['  likes       : いいね数（投稿後に記録、任意）'],
    ['  impressions : インプレッション数（投稿後に記録、任意）'],
    ['  note        : メモ（任意）'],
    [''],
    ['───────────────────────────────────────────────────────────────────────────────'],
    ['📋 queue シート - 投稿キュー（生成されたドラフト）'],
    ['───────────────────────────────────────────────────────────────────────────────'],
    [''],
    ['【目的】'],
    ['  ツールが生成したドラフトが保存されます。'],
    ['  ※基本的に自動入力されます。'],
    [''],
    ['【列の説明】'],
    ['  bot_id         : ボットID'],
    ['  queue_id       : キューID（自動生成）'],
    ['  scheduled_date : 予定日（例：2026-01-19）'],
    ['  scheduled_time : 予定時刻（例：08:10）'],
    ['  category       : カテゴリ'],
    ['  draft_text     : ドラフト本文'],
    ['  status         : ステータス（draft/approved/rejected/posted）'],
    ['  guard_result   : Guardエージェントの判定結果'],
    ['  output_json    : 生成時の詳細データ（JSON）'],
    ['  created_at     : 作成日時'],
    [''],
    ['【ステータスの意味】'],
    ['  draft    : 下書き（確認待ち）'],
    ['  approved : 承認済み（投稿可能）'],
    ['  rejected : 却下（問題あり、要修正）'],
    ['  posted   : 投稿済み'],
    [''],
    ['───────────────────────────────────────────────────────────────────────────────'],
    ['💡 ヒント'],
    ['───────────────────────────────────────────────────────────────────────────────'],
    [''],
    ['・複数のボットを運用する場合は、bot_idを変えて同じシートに追加できます'],
    ['・reference_postsを充実させると、生成される投稿の質が向上します'],
    ['・カテゴリ名は統一することで、参照が正しく機能します'],
    ['・生成されたドラフトは必ず確認してから投稿してください'],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
  ];
  
  await updateRange(`${sheetName}!A1`, guideContent.map(row => [row[0] || '']));
}