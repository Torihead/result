# X-Post Generator

Google Sheetsをデータベースとして使用するX（Twitter）投稿アイデア/ドラフト生成ツール。
マルチエージェントパイプラインで高品質な投稿を自動生成します。

## 機能

- **マルチボット対応**: `bot_id`による複数ボットの管理
- **マルチエージェントパイプライン**:
  1. **Planner**: カテゴリ、意図、フックを含むアイデアを生成
  2. **Writer**: 140文字以内のドラフトを作成
  3. **Guard**: NGワード、ポリシー違反、重複、トーンをチェック
- **Google Sheetsによるデータ管理**: 設定、参照投稿、履歴、キューを一元管理
- **LLMアダプター**: OpenAI/Anthropic対応（簡単に切り替え可能）
- **JSONパース自動修復**: 失敗時に1回リトライ

## セットアップ

### 1. Google Sheets準備

スプレッドシートを作成し、以下の4シートを用意してください。

#### config シート
| bot_id | key | value | note |
|--------|-----|-------|------|
| my_bot | persona | テック系インフルエンサー | ボットのペルソナ |
| my_bot | tone | フレンドリーで情報豊富 | 投稿のトーン |
| my_bot | topics | AI, プログラミング, 生産性 | 対象トピック |
| my_bot | ng_words | 死ね,殺す,バカ | 禁止ワード（カンマ区切り） |
| my_bot | policy_constraints | 政治的な内容は避ける | ポリシー制約 |
| my_bot | posts_per_day | 3 | 1日の投稿数（1-3） |

#### reference_posts シート
| bot_id | ref_id | url | text | category | likes | impressions | engagement_rate | note | added_at |
|--------|--------|-----|------|----------|-------|-------------|-----------------|------|----------|
| my_bot | ref_001 | https://... | 参考投稿テキスト | Tips | 1500 | 50000 | 0.03 | 高パフォーマンス投稿 | 2024-01-01 |

#### history シート
| bot_id | posted_at | text | category | likes | impressions | note |
|--------|-----------|------|----------|-------|-------------|------|
| my_bot | 2024-01-15 08:10 | 過去の投稿テキスト | Tips | 200 | 5000 | |

#### queue シート
| bot_id | queue_id | scheduled_date | scheduled_time | category | draft_text | status | guard_result | output_json | created_at |
|--------|----------|----------------|----------------|----------|------------|--------|--------------|-------------|------------|
| (自動生成) | | | | | | | | | |

### 2. Google Service Account設定

1. [Google Cloud Console](https://console.cloud.google.com/)でプロジェクトを作成
2. Google Sheets APIを有効化
3. サービスアカウントを作成し、JSONキーをダウンロード
4. スプレッドシートをサービスアカウントのメールアドレスと共有（編集権限）

### 3. インストール

```bash
cd x-post-generator
npm install
```

### 4. 環境変数設定

`.env.example`をコピーして`.env`を作成し、以下を設定:

```bash
cp .env.example .env
```

```env
# Google Sheets
GOOGLE_SHEETS_ID=your_spreadsheet_id_here
GOOGLE_SERVICE_ACCOUNT_EMAIL=your-service-account@your-project.iam.gserviceaccount.com
GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"

# LLM Provider
LLM_PROVIDER=openai
OPENAI_API_KEY=sk-your-api-key

# Bot
BOT_ID=my_bot
POSTS_PER_DAY=3
```

## 使い方

### 基本実行

```bash
# 開発モード（ts-node）
npm run dev

# または特定のbot_idを指定
npm run dev -- my_bot

# ビルド後の実行
npm run build
npm start
```

### 出力例

```
🚀 X-Post Generator Starting...

📦 Loaded configuration for bot: my_bot
🤖 LLM Provider: openai
📊 Connecting to Google Sheets...
⚙️ Loading bot configuration...
   Persona: テック系インフルエンサー
   Tone: フレンドリーで情報豊富
   Topics: AI, プログラミング, 生産性
📚 Loading reference posts...
📜 Loading post history...

========================================
[Pipeline] Starting for bot: my_bot
[Pipeline] Target dates: 2024-01-19, 2024-01-20
[Pipeline] Scheduled times: 08:10, 12:40, 20:30
========================================

[Planner] Generating 6 post ideas...
[Planner] Generated 6 ideas

--- Processing slot 1/6 ---
[Writer] Writing draft for: AIの最新トレンド...
[Writer] Draft complete: 128 chars
[Guard] Reviewing draft...
[Guard] Decision: approved

...

📊 Generation Summary:
──────────────────────────────────────────────────
✅ Approved: 5
❌ Rejected: 0
📝 Needs review: 1

✨ Generation complete!
```

## スケジューリング

デフォルトの投稿時間（JST）:
- 08:10 - 朝の通勤時間帯
- 12:40 - 昼休み
- 20:30 - 夜のリラックスタイム

`posts_per_day`の設定により、最初のN個の時間帯が使用されます。

## キューID形式

```
{bot_id}_{YYYYMMDD}_{HHMM}_{index}
例: my_bot_20240119_0810_0
```

## アーキテクチャ

```
src/
├── index.ts           # メインエントリーポイント
├── config.ts          # 環境変数設定
├── sheets/            # Google Sheets操作
│   ├── client.ts      # Sheetsクライアント
│   ├── config.ts      # configシート
│   ├── reference.ts   # reference_postsシート
│   ├── history.ts     # historyシート
│   └── queue.ts       # queueシート
├── llm/               # LLMアダプター
│   ├── adapter.ts     # アダプターファクトリ
│   ├── openai.ts      # OpenAI実装
│   ├── anthropic.ts   # Anthropic実装
│   └── json-parser.ts # JSONパース＆リトライ
├── agents/            # エージェント
│   ├── planner.ts     # Plannerエージェント
│   ├── writer.ts      # Writerエージェント
│   └── guard.ts       # Guardエージェント
├── pipeline/          # パイプライン
│   └── runner.ts      # パイプライン実行
└── utils/             # ユーティリティ
    ├── types.ts       # 型定義
    └── date.ts        # 日付ユーティリティ
```

## エージェント出力形式

### Planner出力
```json
[
  {
    "category": "Tips",
    "intent": "フォロワーに価値を提供",
    "hook": "知ってた？",
    "idea": "AI活用の意外な方法"
  }
]
```

### Writer出力
```json
{
  "draft_text": "知ってた？AIを使えば毎日の...",
  "category": "Tips",
  "char_count": 128
}
```

### Guard出力
```json
{
  "decision": "approved",
  "final_text": "知ってた？AIを使えば毎日の...",
  "reason": "トーンが適切で、NGワードなし",
  "risk_flags": []
}
```

## LLMプロバイダーの切り替え

`.env`で`LLM_PROVIDER`を変更するだけ:

```env
# OpenAI使用
LLM_PROVIDER=openai
OPENAI_API_KEY=sk-...
OPENAI_MODEL=gpt-4o

# Anthropic使用
LLM_PROVIDER=anthropic
ANTHROPIC_API_KEY=sk-ant-...
ANTHROPIC_MODEL=claude-3-5-sonnet-20241022
```

## 定期実行（cron）

```bash
# 毎日6時に実行（JST）
0 6 * * * cd /path/to/x-post-generator && npm start >> /var/log/x-post-generator.log 2>&1
```

## 注意事項

- API使用料が発生します（OpenAI/Anthropic）
- Google Sheets APIには1日あたりのリクエスト制限があります
- 生成されたドラフトは`status=draft`で保存されます。投稿前に確認してください。
- 本番環境では適切なエラーハンドリングとリトライロジックを追加してください。

## ライセンス

MIT
