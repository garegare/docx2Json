# ロードマップ #12 実装計画
# AI・ワークフロー連携によるセマンティック・トレーサビリティ基盤

作成日: 2026-03-09
Issue: #19
対象ブランチ: feature/roadmap-12-ai-workflow

---

## 背景・目的

現行の `--features ai`（Rust から直接 Anthropic API を叩く実装）を廃止し、
**Rust は「パース＋バリデーション」に特化、AI オーケストレーションは外部ワークフロー（n8n 等）に委ねる**
アーキテクチャに置き換える。

これにより:
- AI の不安定なレスポンスや形式異常を外部で吸収できる
- Human-in-the-loop（キーワードリストのレビュー）が自然に組み込める
- Rust バイナリは純粋なデータ変換・整合性チェックに集中できる

---

## アーキテクチャ設計

```
docx2json parse      docx2json            外部ワークフロー (n8n 等)
──────────────       extract-candidates   ──────────────────────────
.docx / .xlsx   →   → JSONL              →  AI (keyword 候補抽出)
                                             ↓ User: keywords.json 確定
document.json   →   inject-tags          →  AI (タグ付与)
                    ← (validated JSON)   ←  inject-tags を呼び出し
                    summarize
                    → tags_summary.json
```

### Rust 側の責務

| コマンド | 入力 | 出力 | 役割 |
|---------|------|------|------|
| `parse` | .docx / .xlsx | document.json | 構造化パース（既存） |
| `extract-candidates` | document.json | candidates.jsonl | LLM 向け候補抽出テキスト生成 |
| `inject-tags` | document.json + タグ | updated document.json | タグ注入 ＋ 整合性バリデーション |
| `summarize` | document.json（複数可） | tags_summary.json | タグ使用統計の集計 |

### 外部ワークフロー（n8n 等）の責務

- AI API の呼び出し・リトライ・形式異常補正
- セクションごとのループ処理
- ブートストラップ（初回 keywords.json 生成）
- 通常運用（keywords.json に基づくタグ付与）

---

## JSON フォーマット変更

### Section（変更点のみ）

```json
{
  "id": "a3f2c1d4e5b6",
  "context_path": ["第1章 導入", "1.1 背景"],
  "heading": "1.1 背景",
  "body_text": "...",
  "assets": [],
  "children": [],
  "metadata": {
    "ai_tags": ["認証", "セキュリティ"]
  }
}
```

- **`id`**: 文書タイトル + context_path を連結した文字列の FNV-1a 16進数ハッシュ（実行間で安定・追加クレート不要）
- **`metadata.ai_tags`**: AI が付与したタグの配列（初期値は空配列）

後方互換性: `#[serde(default)]` により、既存 JSON を読み込んでも `id = ""`, `ai_tags = []` となる。

### keywords.json

```json
{
  "version": "1.0",
  "keywords": ["認証", "API設計", "セキュリティ", "パフォーマンス"],
  "created_at": "2026-03-09T00:00:00Z"
}
```

### tags_summary.json

```json
{
  "generated_at": "2026-03-09T12:34:56Z",
  "total_sections": 42,
  "tagged_sections": 35,
  "untagged_sections": 7,
  "tag_counts": {
    "認証": 12,
    "API設計": 8,
    "セキュリティ": 6
  },
  "top_tags": ["認証", "API設計", "セキュリティ"]
}
```

### candidates.jsonl（JSONL 形式・1行1セクション）

```jsonl
{"id":"a3f2c1","context_path":["第1章"],"heading":"1.1 背景","body_text":"..."}
{"id":"b4e8a2","context_path":["第2章"],"heading":"2.1 設計方針","body_text":"..."}
```

---

## CLI 設計

### サブコマンド構造（後方互換性維持）

```
docx2json [SUBCOMMAND] [OPTIONS]
         └── parse            （省略時のデフォルト）
         └── extract-candidates
         └── inject-tags
         └── summarize
```

サブコマンドを省略した場合は既存の `parse` として動作する（後方互換）。

### `parse`（既存、変更最小限）

```bash
docx2json parse --input ./docs --output ./out [既存オプション]
# 後方互換: サブコマンドなしでも動作
docx2json --input ./docs --output ./out
```

変更: 出力 JSON の各 Section に `id` と `metadata` フィールドを追加。

### `extract-candidates`

```bash
docx2json extract-candidates \
  --input ./out/spec.json \
  --output ./candidates.jsonl \
  [--max-body-chars 2000]    # LLM トークン節約のため body_text を切り詰め
```

- `--input`: `parse` が出力した document.json
- `--output`: JSONL ファイル（1行1セクション、assets・children を除いたコンパクト形式）
- `--max-body-chars`: body_text の最大文字数（デフォルト: 0 = 制限なし）

### `inject-tags`

```bash
docx2json inject-tags \
  --input ./out/spec.json \
  --section-id a3f2c1d4e5b6 \
  --tags '["認証", "API設計"]' \
  --keywords ./keywords.json \
  --output ./out/spec.json

# 初回（keywords.json なし）: --init でバリデーションをスキップ
docx2json inject-tags \
  --input ./out/spec.json \
  --section-id a3f2c1d4e5b6 \
  --tags '["認証候補", "新規タグ"]' \
  --init \
  --output ./out/spec.json
```

- `--tags`: JSON 配列文字列
- `--keywords`: バリデーション用マスターリスト（省略時 or `--init` 時はスキップ）
- バリデーション: keywords に存在しないタグを警告 + 排除
- **リターンコード**: 排除されたタグがある場合でも `0` を返す（ワークフロー継続のため）

### `summarize`

```bash
docx2json summarize \
  --input ./out/ \
  --output ./tags_summary.json

# 単一ファイルも可
docx2json summarize \
  --input ./out/spec.json \
  --output ./tags_summary.json
```

- `--input`: document.json 単体 or JSON を含むディレクトリ
- 複数ファイルのタグ統計を横断集計

---

## 実装フェーズ

### Phase 1: モデル変更（`id` + `metadata` 追加）

**変更ファイル**: `src/models.rs`

```rust
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SectionMetadata {
    #[serde(default)]
    pub ai_tags: Vec<String>,
}

impl Default for SectionMetadata {
    fn default() -> Self { Self { ai_tags: Vec::new() } }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Section {
    #[serde(default)]
    pub id: String,                   // FNV-1a hash（後から fill_section_id() で付与）
    pub context_path: Vec<String>,
    pub heading: String,
    pub body_text: String,
    pub assets: Vec<Asset>,
    pub children: Vec<Section>,
    #[serde(default)]
    pub metadata: SectionMetadata,    // 初期値: { ai_tags: [] }
}
```

**変更ファイル**: `src/parser/mod.rs`（`fill_context_path` の後に `fill_section_id` を呼ぶ）

```rust
// Section ID の生成: 文書タイトル + context_path を連結した FNV-1a ハッシュ
fn fill_section_id(sections: &mut Vec<Section>, title: &str) { ... }
fn fnv1a_hex(s: &str) -> String { ... }
```

### Phase 2: CLI サブコマンド化

**変更ファイル**: `src/main.rs`

- clap の `#[command(subcommand)]` を使ったサブコマンド構造に移行
- サブコマンドなし → 既存 parse 動作にフォールバック（`--input` がある場合）
- `ai.rs` の `--ai` フラグ・`--features ai` ビルドを削除

```rust
#[derive(Subcommand)]
enum Commands {
    Parse(ParseArgs),
    ExtractCandidates(ExtractCandidatesArgs),
    InjectTags(InjectTagsArgs),
    Summarize(SummarizeArgs),
}
```

### Phase 3: `extract-candidates` コマンド

**新規ファイル**: `src/commands/extract_candidates.rs`

- document.json を読み込み
- `Section` を再帰的に走査
- 各 Section を `{ id, context_path, heading, body_text }` に変換（assets/children 除外）
- `--max-body-chars` で body_text を切り詰め
- JSONL として出力

### Phase 4: `inject-tags` コマンド

**新規ファイル**: `src/commands/inject_tags.rs`

1. document.json を読み込み
2. `--section-id` で対象 Section を再帰検索
3. `--keywords` が指定されていれば `keywords.json` を読み込み
4. `--init` でなければタグをバリデーション（無効タグを eprintln + 排除）
5. `section.metadata.ai_tags` を更新
6. 更新済み document.json を出力

### Phase 5: `summarize` コマンド

**新規ファイル**: `src/commands/summarize.rs`

1. `--input` がファイルなら単一ファイル、ディレクトリなら `.json` ファイルを全収集
2. 各 Section を再帰走査して `ai_tags` を集計
3. `total_sections`, `tagged_sections`, `tag_counts` を計算
4. `tags_summary.json` を出力

### Phase 6: `ai.rs` の削除とクリーンアップ

- `src/ai.rs` を削除（または空の stub に）
- `Cargo.toml` の `ureq` optional dependency を削除
- `main.rs` から `--ai` フラグ・`ai::transform()` 呼び出しを削除
- `--features ai` ビルドのサポートを終了

---

## 後方互換性

| 観点 | 対応 |
|------|------|
| 既存の parse 動作 | サブコマンドなし → parse にフォールバック |
| 既存の JSON 読み込み | `id` / `metadata` に `#[serde(default)]` → 旧 JSON でも動作 |
| `--features ai` | Phase 6 で削除（破壊的変更だが issue の意図通り） |
| `--ai` フラグ | Phase 6 で削除（同上） |

---

## ファイル構成（変更後）

```
src/
├── main.rs                   ← サブコマンド分岐追加
├── models.rs                 ← id / SectionMetadata 追加
├── config.rs                 ← 変更なし
├── output.rs                 ← 変更なし
├── splitter.rs               ← 変更なし
├── ai.rs                     ← Phase 6 で削除
├── parser/
│   ├── mod.rs                ← fill_section_id 追加
│   ├── docx.rs               ← 変更なし
│   ├── xlsx.rs               ← 変更なし
│   └── xlsx_advanced.rs      ← 変更なし
└── commands/
    ├── mod.rs                ← 新規
    ├── extract_candidates.rs ← 新規
    ├── inject_tags.rs        ← 新規
    └── summarize.rs          ← 新規
```

---

## 実装順序の推奨

1. Phase 1（モデル変更）→ ビルド確認 → コミット
2. Phase 2（CLI 構造）→ 既存 parse 動作確認 → コミット
3. Phase 3（extract-candidates）→ コミット
4. Phase 4（inject-tags）→ コミット
5. Phase 5（summarize）→ コミット
6. Phase 6（ai.rs 削除）→ コミット → PR 作成

---

## n8n ワークフロー利用例

```bash
# Step 1: ドキュメントをパース
docx2json parse --input /data/spec.docx --output /data/output.json

# Step 2: AI 向けテキストを生成
docx2json extract-candidates --input /data/output.json --output /data/candidates.jsonl

# Step 3: [外部] AI でキーワード候補を抽出 → keywords.json を確定（Human-in-the-loop）

# Step 4: セクションごとにタグを注入（ワークフローがループで呼び出す）
docx2json inject-tags \
  --input /data/output.json \
  --section-id a3f2c1d4e5b6 \
  --tags '["認証", "API設計"]' \
  --keywords /data/keywords.json \
  --output /data/output.json

# Step 5: プロジェクト全体のタグ統計を集計
docx2json summarize --input /data/ --output /data/tags_summary.json
```
