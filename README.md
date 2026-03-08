# Docx/Xlsx to AI-Ready JSON Converter (Rust)

このツールは、Microsoft Office形式（.docx, .xlsx）のドキュメントを解析し、LLM（大規模言語モデル）へのインプットに最適化された構造化JSONへ変換する、Rust製の高パフォーマンス・コンバーターです。

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
| **高速XMLストリームパース（Low Memory）** | ✅ | `quick-xml` を使用。低メモリ消費で高速処理。画像はバイナリで保持し JSON 出力時のみ Base64 変換（Lazy Serialization）。 |
| **テーブルのMarkdown変換** | ✅ | `w:tbl` を走査し Markdown表形式で `body_text` に統合。セル内改行を `<br>` に変換し表構造を保護。空セルは ` ` で補填。 |
| **エラーハンドリング強化** | ✅ | `anyhow` 導入。コンテキスト付きエラーチェーンを階層表示。破損ファイルがあっても並列処理を継続。 |
| **一括バッチ処理** | ✅ | `rayon` による並列処理で複数ファイルを高速変換。 |
| **設定ファイルによる見出し制御** | ✅ | `docx2json.json` を入力Dir・カレントDir・バイナリDirの順で自動探索。全角スタイル名（`見出し１` 等）も正規化して認識。 |
| **箇条書きのMarkdown化** | ✅ | `w:numPr` を検知し `- ` / `1. ` 形式に変換。`word/numbering.xml` を参照し番号付き・箇条書きを判別。インデントレベル（`w:ilvl`）に応じてネストを表現。 |
| **RAG最適化（context_path）** | ✅ | 各セクションにルートから自身への見出しパスリスト（`context_path`）を付与。チャンク分割後も文書内の位置をAIが把握可能。 |
| **セクション単位のチャンク分割** | ✅ | `--split <level>` で指定した深さ（1=最上位）でセクションを分割し、セクションごとに個別 JSON を出力。子セクションの本文は Markdown 見出しとして統合。 |
| **画像リサイズ・圧縮** | ✅ | `--image-max-px <N>` で長辺を N px にリサイズし JPEG 再エンコード。`--image-quality <Q>` で品質調整。JSON 肥大化を防止。 |
| **OMML数式のLaTeX変換** | ✅ | `m:oMath` を検知し `$...$` 形式で埋め込み。分数（`\frac{}`）・上付き（`^{}`）・下付き（`_{}`）・平方根（`\sqrt{}`）を LaTeX に変換。 |
| **見出し検出の柔軟化** | ✅ | `heading_styles` キーに `"prefix:<str>"` / `"regex:<pattern>"` 記法で前方一致・正規表現マッチを指定可能。 |
| **進捗表示（CLI）** | ✅ | `indicatif` によるアニメーションプログレスバー。経過時間・処理中ファイル名をリアルタイム表示。完了後に成功/失敗の一覧を出力。 |
| **AI連携フォーマッティング** | 🚧 | `--features ai` で有効化。API呼び出しは未実装。 |
| **XLSXパース** | 🚧 | 構造のみ実装済み、本実装は今後の対応。 |

## 🗺 ロードマップ（残タスク）

| # | 機能 | 状態 | 概要 |
| :- | :--- | :---: | :--- |
| 3 | **XLSXパース本実装** | 🚧 | シート名を `heading`、内容をMarkdownテーブルとして `body_text` に格納。Shared Strings対応 |
| 9 | **PPTXパース** | 🔲 | スライド単位で `Section` 化。テキストボックスを座標順に結合、スライドノートを補足コンテキストとして抽出 |
| 10 | **神エクセル対応** | 🔲 | セル結合解決、書式ベースの見出し判定、浮遊テキストボックス抽出 |

## 🛠 技術スタック
| カテゴリ | ライブラリ | 選定理由 |
| :--- | :--- | :--- |
| **Core** | `Rust` | メモリ安全、高速、クロスコンパイルの容易性 |
| **Parsing** | `zip`, `quick-xml` | Officeの実体(ZIP+XML)を高速・省メモリで処理 |
| **Parallel** | `rayon` | 複数ファイルの並列処理によるスループット向上 |
| **Networking** | `ureq` | 依存関係が極めて少なく、API連携に十分な機能を保持 |
| **Serialization**| `serde`, `serde_json` | 厳密な型定義に基づいた安全なJSON生成 |
| **CLI** | `clap` | 型安全なCLI引数パース |
| **Progress** | `indicatif` | アニメーションプログレスバーによるリアルタイム進捗表示 |
| **Regex** | `regex` | 見出しスタイル名の正規表現マッチング |

## 🔄 処理フロー

1. **Scan**: ディレクトリ内の `.docx`, `.xlsx` をリストアップ。
2. **Parse**: XMLを走査し、見出しレベルに基づいてセクションを分割。変更履歴を最新状態で確定。
3. **Internal Structure**: 段落テキストを結合し、画像データを抽出してメモリ上に保持。
4. **AI Transformation**: APIを経由し、AIによるコンテキスト整形を実行（`--features ai` 時）。
5. **Output**: 元のファイルパスに基づき、拡張子を `.json` に置換して保存。

## 📦 ビルド & 実行

```bash
# 開発ビルド
cargo build

# AI連携機能あり（ureq が有効化される）
cargo build --features ai

# 実行例
cargo run -- --input ./docs --output ./out

# 単一ファイル
cargo run -- --input ./doc.docx --output ./out

# 設定ファイルを明示指定
cargo run -- --input ./docs --config ./my-config.json

# AI変換あり（ANTHROPIC_API_KEY 環境変数が必要）
cargo run --features ai -- --input ./docs --ai

# チャンク分割（最上位セクション単位で個別 JSON を出力）
cargo run -- --input ./docs --output ./out --split 1

# チャンク分割（2階層目単位で細かく分割）
cargo run -- --input ./docs --output ./out --split 2

# 画像リサイズ（長辺 1024px 以下に縮小、JPEG 品質 80%）
cargo run -- --input ./docs --output ./out --image-max-px 1024

# 画像リサイズ + 品質指定
cargo run -- --input ./docs --output ./out --image-max-px 512 --image-quality 70
```

## ⚙️ 設定ファイル（`docx2json.json`）

入力ディレクトリに `docx2json.json` を置くことで見出し検出をカスタマイズできます。
設定ファイルが存在しない場合はデフォルト設定が使用されます。

```json
{
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
  "image_max_px": 1024,
  "image_quality": 80
}
```

### 設定項目

| キー | デフォルト | 説明 |
| :--- | :--- | :--- |
| `heading_styles` | 標準スタイル名セット | `スタイル名: レベル` のマッピング。Heading1〜3・見出し1〜3を既定で認識。キー記法は下表参照。 |
| `ppr_underline_as_heading` | `true` | 段落デフォルト書式（`w:pPr > w:rPr`）の下線を見出しとして扱う。 |
| `run_underline_as_heading` | `false` | ランレベル（`w:r > w:rPr`）の下線を見出しとして扱う。Wordの「見出し」スタイルを使わず直接書式で見出しを表現した文書向け。 |
| `image_max_px` | `0`（無効） | 画像の最大辺長（px）。超過する画像をリサイズし JPEG 再エンコード。`--image-max-px` CLI 引数が優先。 |
| `image_quality` | `80` | JPEG 再エンコード品質（1〜100）。`image_max_px > 0` のときのみ有効。`--image-quality` CLI 引数が優先。 |

### `heading_styles` キー記法（#12）

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
AIが文脈を即座に理解できるよう、以下の再帰的なJSON構造を出力します。

```json
{
  "title": "ドキュメントタイトル",
  "sections": [
    {
      "context_path": ["第1章 導入"],
      "heading": "第1章 導入",
      "body_text": "セクション内の連続する段落を結合したテキスト。\n\n| 項目 | 値 |\n|------|----|\n| A | 1 |",
      "assets": [
        {
          "type": "image",
          "title": "図1 構成図",
          "data": "iVBORw0KGgoAAAANSUhEUgAA..."
        }
      ],
      "children": [
        {
          "context_path": ["第1章 導入", "1.1 背景"],
          "heading": "1.1 背景",
          "body_text": "サブセクションのテキスト。",
          "assets": [],
          "children": []
        }
      ]
    }
  ]
}
```

## 📁 ファイル構成

```
src/
├── main.rs        # CLIエントリー（引数パース、ディレクトリ走査、並列処理）
├── models.rs      # データ構造（Document / Section / Asset）
├── config.rs      # 設定ファイルの読み込みと管理
├── ai.rs          # AI変換（--features ai で有効化）
├── output.rs      # JSON書き出し
└── parser/
    ├── mod.rs     # ファイル種別ルーティング
    ├── docx.rs    # DOCXパーサー（実装済み）
    └── xlsx.rs    # XLSXパーサー（スタブ → 実装予定）
```
