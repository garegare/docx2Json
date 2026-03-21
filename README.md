# Docx/Xlsx/Pptx to AI-Ready JSON Converter (Rust)

このツールは、Microsoft Office形式（.docx, .xlsx, .pptx）のドキュメントを解析し、LLM（大規模言語モデル）へのインプットに最適化された構造化JSONへ変換する、Rust製の高パフォーマンス・コンバーターです。

## 🎯 プロジェクトの目的
AIによる文書解析（RAGや要約）の精度を最大化するため、単なるテキスト抽出ではなく、文書の**階層構造（見出し・段落の関係）**を維持したまま、ノイズを除去したクリーンなJSONを生成します。

## 🚀 実装済み機能

| 機能 | 状態 | 説明 |
| :--- | :---: | :--- |
| **高速XMLストリームパース** | ✅ | `quick-xml` を使用。低メモリ消費で高速処理。 |
| **変更履歴の自動確定抽出** | ✅ | `w:del`（削除）を無視、`w:ins`（挿入）を採用し最新状態を取得。 |
| **再帰的セクション構造** | ✅ | 見出しレベルを検知し、ネストされたJSON構造を構築。 |
| **アセット統合（画像）** | ✅ | 画像をBase64エンコードして `assets` 配列に紐付け。`Start`/`Empty` 両イベント対応済み。 |
| **画像メタデータ抽出** | ✅ | `wp:docPr` の `descr`（代替テキスト）/ `name` を `Asset.title` に格納。AIが画像の意味を把握可能に。 |
| **Lazy Serialization（省メモリ）** | ✅ | 画像はバイナリで保持し JSON 出力時のみ Base64 変換。並列処理時のメモリ使用量を抑制。 |
| **テーブルのMarkdown変換** | ✅ | `w:tbl` を走査し Markdown表形式で `body_text` に統合。セル内改行を `<br>` に変換し表構造を保護。 |
| **エラーハンドリング強化** | ✅ | `anyhow` 導入。コンテキスト付きエラーチェーンを階層表示。破損ファイルがあっても並列処理を継続。 |
| **一括バッチ処理** | ✅ | `rayon` による並列処理で複数ファイルを高速変換。 |
| **設定ファイルによる見出し制御** | ✅ | `docx2json.json` を入力Dir・カレントDir・バイナリDirの順で自動探索。全角スタイル名（`見出し１` 等）も正規化して認識。 |
| **箇条書きのMarkdown化** | ✅ | `w:numPr` を検知し `- ` / `1. ` 形式に変換。インデントレベルに応じてネストを表現。 |
| **RAG最適化（context_path）** | ✅ | 各セクションにルートから自身への見出しパスリスト（`context_path`）を付与。チャンク分割後も文書内の位置をAIが把握可能。 |
| **セクション単位のチャンク分割** | ✅ | `--split <level>` で指定した深さでセクションを分割し、セクションごとに個別 JSON を出力（RAG 向け）。 |
| **画像リサイズ・圧縮** | ✅ | `--image-max-px <N>` で長辺を N px にリサイズし JPEG 再エンコード。`--image-quality <Q>` で品質調整。 |
| **OMML数式のLaTeX変換** | ✅ | `m:oMath` を検知し `$...$` 形式で埋め込み。分数・上付き・下付き・平方根を LaTeX に変換。 |
| **見出し検出の柔軟化** | ✅ | `heading_styles` キーに `"prefix:<str>"` / `"regex:<pattern>"` 記法で前方一致・正規表現マッチを指定可能。 |
| **進捗表示（CLI）** | ✅ | `indicatif` によるアニメーションプログレスバー。経過時間・処理中ファイル名をリアルタイム表示。 |
| **XLSXパース本実装** | ✅ | シート名を `heading`、内容をMarkdownテーブルとして `body_text` に格納。Shared Strings・リッチテキスト対応。行数上限（`xlsx_max_rows`）超過時はヘッダーを保持したまま子Sectionに分割。 |
| **神エクセル対応** | ✅ | セル結合解決（`<mergeCell>` 展開）、書式ベース見出し判定（太字・背景色で Section を自動生成）、浮遊テキストボックス抽出（`xl/drawings/drawing*.xml`）。設定で有効化。 |
| **セクション ID 付与** | ✅ | 各セクションに FNV-1a 64bit ハッシュの安定 ID（`id`）を付与。文書更新後も同一セクションを同定可能。 |
| **AI タグ候補テキスト抽出** | ✅ | `extract-candidates` サブコマンド。LLM 向けに JSONL 形式でセクションを出力。`--max-body-chars` でトークン節約。 |
| **AI タグ注入・バリデーション** | ✅ | `inject-tags` サブコマンド。セクション ID 指定でタグを注入し `keywords.json` でバリデーション。 |
| **タグ使用統計集計** | ✅ | `summarize` サブコマンド。複数ドキュメントのタグ使用頻度を横断集計して `tags_summary.json` を生成。 |
| **PPTXパース** | ✅ | スライド単位で `Section` 化。テキストボックスを Y 座標順に結合し、スライドノートを `[ノート]` として body_text 末尾に付加。画像を `assets` に格納。 |
| **`elements` 構造化配列** | ✅ | 段落・テーブル・画像参照を順序保持した `elements` 配列で管理。RAG チャンク分割時のセマンティック欠落を防止。 |
| **意味的役割（SemanticRole）** | ✅ | スタイル名から `Warning` / `Note` / `Tip` / `CodeBlock` / `Quote` 等を自動推定し `elements[].metadata.role` に付与。日本語スタイル名にも対応。 |
| **カスタム SemanticRole マッピング** | ✅ | `docx.semantic_role_styles` でスタイル名 → SemanticRole を設定ファイルから外部注入可能。 |
| **出力フィールド制御** | ✅ | `output.include_body_text`（デフォルト: `false`）と `output.include_base64`（デフォルト: `true`）で JSON サイズを最適化。 |
| **見出しスタイル検査** | ✅ | `inspect-styles` サブコマンド。DOCX の `word/styles.xml` を走査し、見出しスタイル一覧と `docx2json.json` 用 `heading_styles` 設定スニペットを JSON で出力。 |
| **AsciiDoc 変換出力** | ✅ | `to-asciidoc` サブコマンド。`parse` で生成した `document.json` を AsciiDoc 形式（`.adoc`）に変換。セル結合（rowspan / colspan）を正確に再現し、複雑な表も構造を維持して出力。 |

## 🛠 技術スタック
| カテゴリ | ライブラリ | 選定理由 |
| :--- | :--- | :--- |
| **Core** | `Rust` | メモリ安全、高速、クロスコンパイルの容易性 |
| **Parsing** | `zip`, `quick-xml` | Officeの実体(ZIP+XML)を高速・省メモリで処理 |
| **Parallel** | `rayon` | 複数ファイルの並列処理によるスループット向上 |
| **Serialization**| `serde`, `serde_json` | 厳密な型定義に基づいた安全なJSON生成 |
| **CLI** | `clap` | 型安全なCLI引数パース。サブコマンド構造に対応 |
| **Progress** | `indicatif` | アニメーションプログレスバーによるリアルタイム進捗表示 |
| **Regex** | `regex` | 見出しスタイル名の正規表現マッチング |

## 🔄 処理フロー

### parse（ドキュメント変換）
1. **Scan**: ディレクトリ内の `.docx`, `.xlsx` をリストアップ。
2. **Parse**: XMLを走査し、見出しレベルに基づいてセクションを分割。変更履歴を最新状態で確定。
3. **Internal Structure**: 段落テキストを結合し、画像データを抽出してメモリ上に保持。
4. **ID 付与**: 各セクションに FNV-1a 64bit ハッシュの安定 ID を付与。
5. **Output**: `document.json` を出力。

### セマンティック・タグ連携フロー（外部ワークフローとの連携）

```
docx2json parse         docx2json               外部ワークフロー (n8n 等)
──────────────          ──────────────────────  ──────────────────────────
.docx / .xlsx   →   →  extract-candidates      →  AI (keyword 候補抽出)
                        → candidates.jsonl          ↓
                                                User: keywords.json 確定
                                                    ↓
document.json   →   →  inject-tags             ←  AI (タグ付与)
                        ← document.json
                        (metadata.ai_tags 更新済み)

                        summarize
                        → tags_summary.json
```

## 📦 ビルド & 実行

```bash
# 開発ビルド
cargo build

# parse サブコマンド（明示）
cargo run -- parse --input ./docs --output ./out

# 単一ファイル（.docx / .xlsx / .pptx いずれも対応）
cargo run -- parse --input ./doc.docx --output ./out
cargo run -- parse --input ./slides.pptx --output ./out

# 設定ファイルを明示指定
cargo run -- parse --input ./docs --config ./my-config.json

# チャンク分割（最上位セクション単位で個別 JSON を出力）
cargo run -- parse --input ./docs --output ./out --split 1

# 画像リサイズ（長辺 1024px 以下に縮小、JPEG 品質 80%）
cargo run -- parse --input ./docs --output ./out --image-max-px 1024

# XLSXの大きな表を100行ずつ子Sectionに分割
cargo run -- parse --input ./sheets --output ./out --xlsx-max-rows 100

# 神エクセル対応（設定ファイルで xlsx.heading.enabled: true を指定）
cargo run -- parse --input ./god-excel.xlsx --output ./out --config ./my-config.json

# 実効設定を確認（設定ファイル + CLI 引数の適用後の JSON を出力）
cargo run -- parse --input ./docs --dump-config

# 設定ファイルの雛形を生成
cargo run -- parse --dump-config > docx2json.json
```

### `parse` サブコマンド CLI オプション一覧

| オプション | 短縮形 | デフォルト | 設定ファイル対応 | 説明 |
| :--- | :---: | :--- | :---: | :--- |
| `--input <PATH>` | `-i` | `.`（カレント） | — | 入力ファイルまたはディレクトリ |
| `--output <PATH>` | `-o` | 入力と同じ場所 | — | 出力ディレクトリ |
| `--config <PATH>` | — | 自動探索 | — | 設定ファイル（`docx2json.json`）のパス |
| `--split <LEVEL>` | — | 無効 | — | セクション分割の深さ（RAG 向け個別 JSON 出力） |
| `--image-max-px <N>` | — | `0`（無効） | `image.max_px` | 画像の最大辺長（px）。設定ファイルより優先 |
| `--image-quality <Q>` | — | `80` | `image.quality` | JPEG 品質（1〜100）。設定ファイルより優先 |
| `--xlsx-max-rows <N>` | — | `0`（無効） | `xlsx.max_rows` | XLSX シートの最大行数。設定ファイルより優先 |
| `--dump-config` | — | — | — | 実効設定（設定ファイル + CLI 引数の適用後）を JSON で出力して終了 |

### `inspect-styles` — DOCX 見出しスタイル検査

カスタムスタイルを使っている DOCX をパースする前に、どのスタイル名が見出しとして定義されているかを確認できます。
出力された `docx.heading_styles` は `docx2json.json` にそのまま貼り付けて使用できます。

```bash
# 標準出力に JSON を表示（stderr に検出サマリー）
docx2json inspect-styles --input ./sample.docx

# ファイルに書き出す
docx2json inspect-styles --input ./sample.docx --output ./heading_config.json
```

**出力例（stdout）:**
```json
{
  "docx": {
    "heading_styles": {
      "1": 1,
      "heading 1": 1,
      "2": 2,
      "heading 2": 2,
      "3": 3,
      "heading 3": 3
    }
  },
  "styles": [
    { "style_id": "1", "name": "heading 1", "level": 1 },
    { "style_id": "2", "name": "heading 2", "level": 2 },
    { "style_id": "3", "name": "heading 3", "level": 3 }
  ]
}
```

**出力フィールド:**

| フィールド | 説明 |
| :--- | :--- |
| `docx.heading_styles` | `docx2json.json` に貼り付けられる設定スニペット。`styleId` と表示名の両方をキーとして収録 |
| `styles[].style_id` | XML 内部の `w:styleId`（`w:pStyle` で参照される識別子） |
| `styles[].name` | Word 上の表示スタイル名（`w:name w:val`） |
| `styles[].level` | 見出しレベル（1〜9）。`w:outlineLvl` の値 + 1 |

> **ヒント:** `inspect-styles` で検出されなかった場合は、そのファイルが Word 標準の見出しスタイルを使っていない可能性があります。
> `ppr_underline_as_heading` や `run_underline_as_heading` の設定を検討してください。

### `to-asciidoc` — AsciiDoc 変換

`parse` で生成した `document.json` を AsciiDoc 形式（`.adoc`）に変換します。
セル結合（rowspan / colspan）を含む複雑な表も AsciiDoc のスパン記法で正確に出力します。

```bash
# 基本（出力先省略時は入力 JSON と同じディレクトリに .adoc を生成）
docx2json to-asciidoc --input ./output.json

# 出力先を明示指定
docx2json to-asciidoc --input ./output.json --output ./doc.adoc
```

**AsciiDoc 出力の主な特徴:**

| 機能 | 説明 |
| :--- | :--- |
| セクション見出し | `context_path` の深さに応じた `==` 〜 `======` レベルで出力 |
| 表（colspan / rowspan） | `2+\|`（colspan）、`.3+\|`（rowspan）、`2.3+\|`（両方）の AsciiDoc スパン記法を使用 |
| 空列・幽霊列の除去 | マージ境界だけに存在する空列を自動検出・除去し列数を最適化 |
| 空行の除去 | rowspan 端数で生じる全空行（中間・末尾とも）を自動除去 |
| 意味的役割 | `NOTE:` / `WARNING:` / `TIP:` などの AsciiDoc admonition に変換 |
| セル内改行 | `\n` → ` +\n`（AsciiDoc ハードラインブレーク）で改行を保持 |
| 特殊文字エスケープ | セル内の `\|` と `\[` を自動エスケープ |

### AI・ワークフロー連携コマンド

```bash
# Step 1: ドキュメントをパース（セクション ID も付与される）
docx2json parse --input ./docs/spec.docx --output ./output.json

# Step 2: LLM 向けテキストを JSONL 形式で抽出
docx2json extract-candidates \
  --input ./output.json \
  --output ./candidates.jsonl \
  --max-body-chars 2000   # LLM トークン節約のため body_text を切り詰め

# Step 3: [外部] AI でキーワード候補を抽出 → keywords.json を確定（Human-in-the-loop）

# Step 4: セクションごとにタグを注入（外部ワークフローがループで呼び出す）
docx2json inject-tags \
  --input ./output.json \
  --section-id fac9d8c798625bae \
  --tags '["認証", "API設計"]' \
  --keywords ./keywords.json \
  --output ./output.json

# 初回（keywords.json 未作成時）: --init でバリデーションをスキップ
docx2json inject-tags \
  --input ./output.json \
  --section-id fac9d8c798625bae \
  --tags '["認証候補", "新規タグ"]' \
  --init \
  --output ./output.json

# Step 5: プロジェクト全体のタグ統計を集計
docx2json summarize \
  --input ./output/          # ディレクトリ指定で複数ファイルを横断集計
  --output ./tags_summary.json

# 単一ファイルも可
docx2json summarize \
  --input ./output.json \
  --output ./tags_summary.json
```

## ⚙️ 設定ファイル（`docx2json.json`）

入力ディレクトリに `docx2json.json` を置くことでパース動作をカスタマイズできます。
設定ファイルが存在しない場合はデフォルト設定が使用されます。

設定は **`docx`・`image`・`xlsx`・`output`** の4つのセクションに分かれています。
各セクションは省略可能で、省略した場合はデフォルト値が使用されます。

```json
{
  "docx": {
    "heading_styles": {
      "Heading1": 1,
      "Heading2": 2,
      "Heading3": 3,
      "見出し1": 1,
      "見出し2": 2,
      "見出し3": 3,
      "prefix:第": 1,
      "regex:^\\d+\\.": 2
    },
    "ppr_underline_as_heading": true,
    "run_underline_as_heading": false,
    "semantic_role_styles": {
      "MyCustomAlert": "warning",
      "SpecialNote": "note"
    }
  },
  "image": {
    "max_px": 1024,
    "quality": 80
  },
  "xlsx": {
    "max_rows": 100
  },
  "output": {
    "include_body_text": false,
    "include_base64": true
  }
}
```

### 設定セクション一覧

#### `docx` — DOCX パーサー設定

| キー | デフォルト | 説明 |
| :--- | :--- | :--- |
| `heading_styles` | 標準スタイル名セット | `スタイル名: レベル` のマッピング。Heading1〜3・見出し1〜3を既定で認識。キー記法は下表参照。 |
| `ppr_underline_as_heading` | `true` | 段落デフォルト書式（`w:pPr > w:rPr`）の下線を見出し（level 1）として扱う。 |
| `run_underline_as_heading` | `false` | ランレベル（`w:r > w:rPr`）の下線を見出し（level 1）として扱う。Wordの「見出し」スタイルを使わず直接書式で見出しを表現した文書向け。 |
| `semantic_role_styles` | `{}`（空） | スタイル名 → SemanticRole のカスタムマッピング。組み込みルールより優先される。値は `"note"` / `"warning"` / `"tip"` / `"code_block"` / `"quote"` / `"bullet_list"` / `"ordered_list"` のいずれか。 |

#### `image` — 画像処理設定

| キー | デフォルト | CLI 引数 | 説明 |
| :--- | :--- | :--- | :--- |
| `max_px` | `0`（無効） | `--image-max-px` | 画像の最大辺長（px）。超過する画像をリサイズし JPEG 再エンコード。CLI 引数が優先。 |
| `quality` | `80` | `--image-quality` | JPEG 再エンコード品質（1〜100）。`max_px > 0` のときのみ有効。CLI 引数が優先。 |

#### `xlsx` — XLSX パーサー設定

| キー | デフォルト | CLI 引数 | 説明 |
| :--- | :--- | :--- | :--- |
| `max_rows` | `0`（無効） | `--xlsx-max-rows` | XLSXシートの最大データ行数。超過した場合ヘッダー行を引き継いだ子Sectionに分割。CLI 引数が優先。 |
| `heading` | `null`（無効） | — | 神エクセル対応の書式ベース見出し判定設定。`enabled: true` で有効化。詳細は下表参照。 |

### `xlsx.heading` — 神エクセル対応設定（#10）

セル結合の展開と浮遊テキストボックス抽出は `xlsx.heading` の設定に関わらず**常に実行**されます。
書式ベースの見出し判定のみ `enabled: true` が必要です。

```json
{
  "xlsx": {
    "max_rows": 100,
    "heading": {
      "enabled": true,
      "detect_bold": true,
      "detect_fill": true,
      "heading_font_size_threshold": 0.0,
      "heading_cell_ratio": 0.5
    }
  }
}
```

| キー | デフォルト | 説明 |
| :--- | :--- | :--- |
| `enabled` | `false` | `true` にすると書式ベース見出し判定モード（`xlsx_advanced` パーサー）に切り替え。`false` は従来モード（後方互換） |
| `detect_bold` | `true` | 太字セルを見出し条件に含める |
| `detect_fill` | `true` | 非白・非透明の背景色付きセルを見出し条件に含める |
| `heading_font_size_threshold` | `0.0`（無効） | 見出し判定の最小フォントサイズ（pt）。0.0 で無効 |
| `heading_cell_ratio` | `0.5` | 行内で見出し書式セルが占める割合の閾値。これ以上なら見出し行と判定 |

**動作モードの違い:**

| モード | 使用パーサー | セル結合 | 書式見出し | Drawingテキスト |
| :--- | :--- | :---: | :---: | :---: |
| `xlsx.heading` なし / `enabled: false` | `xlsx.rs`（従来） | ❌ | ❌ | ❌ |
| `enabled: true` | `xlsx_advanced.rs` | ✅ | ✅ | ✅ |

#### `output` — JSON 出力設定

| キー | デフォルト | 説明 |
| :--- | :--- | :--- |
| `include_body_text` | `false` | `true` にすると後方互換用フラットテキスト `body_text` を出力。新規利用では `elements` を参照推奨。 |
| `include_base64` | `true` | `false` にすると画像 Base64 データ（`assets[].data`）を省略。JSON サイズを大幅に削減できる。 |

### `docx.heading_styles` キー記法

| 記法 | 例 | マッチ条件 |
| :--- | :--- | :--- |
| 通常文字列 | `"Heading1": 1` | スタイル名と**完全一致**（全角→半角正規化後） |
| `"prefix:<文字列>"` | `"prefix:第": 1` | スタイル名が指定文字列で**前方一致** |
| `"regex:<パターン>"` | `"regex:^\\d+\\.": 2` | スタイル名が**正規表現**にマッチ |

優先順位: **完全一致 > 前方一致 > 正規表現**（同一カテゴリ内は設定ファイルの順序）

### 見出し検出の優先順位

1. `heading_styles` に一致する `w:pStyle`（最優先）
2. `w:pPr > w:rPr > w:u`（段落デフォルト下線） — `ppr_underline_as_heading: true` 時
3. `w:r > w:rPr > w:u`（ランレベル下線） — `run_underline_as_heading: true` 時

## 📂 出力フォーマット仕様

### `document.json`（`parse` コマンドの出力）

```json
{
  "title": "ドキュメントタイトル",
  "sections": [
    {
      "id": "fac9d8c798625bae",
      "context_path": ["第1章 導入"],
      "heading": "第1章 導入",
      "elements": [
        {
          "type": "paragraph",
          "text": "セクション内の段落テキスト。",
          "metadata": { "role": "note", "alignment": "left" }
        },
        {
          "type": "table",
          "rows": [["項目", "値"], ["A", "1"]],
          "metadata": {}
        },
        {
          "type": "asset_ref",
          "asset_id": "rId5",
          "metadata": {}
        }
      ],
      "assets": [
        {
          "type": "image",
          "id": "rId5",
          "title": "図1 構成図",
          "data": "iVBORw0KGgoAAAANSUhEUgAA..."
        }
      ],
      "children": [
        {
          "id": "581702679c106bbd",
          "context_path": ["第1章 導入", "1.1 背景"],
          "heading": "1.1 背景",
          "elements": [
            {
              "type": "paragraph",
              "text": "サブセクションのテキスト。",
              "metadata": {}
            }
          ],
          "assets": [],
          "children": [],
          "metadata": {
            "ai_tags": ["認証", "API設計"]
          }
        }
      ],
      "metadata": {
        "ai_tags": []
      }
    }
  ]
}
```

- **`id`**: 文書タイトル + context_path を連結した FNV-1a 64bit ハッシュ（16文字 16進数）。実行間で安定。
- **`elements`**: 段落・テーブル・画像参照を出現順で保持する配列。`type` フィールドで種別を識別。
  - `paragraph`: テキスト段落。`metadata.role` に SemanticRole（`warning` / `note` / `tip` / `code_block` / `quote` 等）が付与される場合がある。
  - `table`: テーブル。`rows` は `string[][]` の2次元配列。
  - `asset_ref`: 画像参照。`asset_id` が `assets[].id` に対応。
- **`body_text`**: 後方互換用フラットテキスト。`output.include_body_text: true` を設定した場合のみ出力（デフォルト: 出力しない）。
- **`assets[].data`**: Base64 エンコードされた画像データ。`output.include_base64: false` で省略可能（デフォルト: 出力する）。
- **`metadata.ai_tags`**: `inject-tags` で注入された AI タグ。初期値は空配列。

### `candidates.jsonl`（`extract-candidates` コマンドの出力）

各セクションを1行1JSON（JSONL形式）で出力。LLM への入力に最適化。

```jsonl
{"id":"fac9d8c798625bae","context_path":["第1章 導入"],"heading":"第1章 導入","body_text":"..."}
{"id":"581702679c106bbd","context_path":["第1章 導入","1.1 背景"],"heading":"1.1 背景","body_text":"..."}
```

### `keywords.json`（外部ワークフローが生成・Human-in-the-loop で確定）

```json
{
  "version": "1.0",
  "keywords": ["認証", "API設計", "セキュリティ", "パフォーマンス"],
  "created_at": "2026-03-09T00:00:00Z"
}
```

### `tags_summary.json`（`summarize` コマンドの出力）

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

## 📁 ファイル構成

```
src/
├── main.rs              # CLIエントリー（サブコマンド分岐・並列処理・進捗表示）
├── models.rs            # データ構造（Document / Section / SectionMetadata / Asset）
├── config.rs            # 設定ファイルの読み込みと管理（XlsxHeadingConfig を含む）
├── output.rs            # JSON書き出し
├── splitter.rs          # --split によるチャンク分割出力
├── parser/
│   ├── mod.rs           # ファイル種別ルーティング + fill_context_path / fill_section_id
│   ├── docx.rs          # DOCXパーサー（elements / SemanticRole 対応）
│   ├── pptx.rs          # PPTXパーサー（スライド単位 Section 化）
│   ├── xlsx.rs          # XLSXパーサー（従来モード・後方互換）
│   └── xlsx_advanced.rs # 神エクセル対応パーサー（セル結合・書式見出し・Drawing）
└── commands/
    ├── mod.rs               # サブコマンドモジュール宣言
    ├── extract_candidates.rs # LLM 向け候補テキスト抽出（→ JSONL）
    ├── inject_tags.rs        # AI タグ注入 + keywords.json バリデーション
    ├── inspect_styles.rs     # DOCX 見出しスタイル検査（→ heading_styles 設定スニペット）
    ├── summarize.rs          # タグ使用統計横断集計（→ tags_summary.json）
    └── to_asciidoc.rs        # document.json → AsciiDoc 変換（セル結合・SemanticRole 対応）
```
