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
| **一括バッチ処理** | ✅ | `rayon` による並列処理で複数ファイルを高速変換。 |
| **設定ファイルによる見出し制御** | ✅ | `docx2json.json` で見出し検出方法をカスタマイズ可能。 |
| **AI連携フォーマッティング** | 🚧 | `--features ai` で有効化。API呼び出しは未実装。 |
| **XLSXパース** | 🚧 | 構造のみ実装済み、本実装は今後の対応。 |

## 🗺 ロードマップ（Issue #2 より）

### 🔴 優先度：高（AI解析精度・信頼性に直結）

| # | 機能 | 概要 |
| :- | :--- | :--- |
| 1 | **テーブルのMarkdown変換** | `w:tbl` → `w:tr` → `w:tc` を走査し、Markdown表形式（`\| cell \|`）で `body_text` に統合 |
| 2 | **エラーハンドリング強化** | `anyhow` 導入。ファイル・行レベルのコンテキスト付きエラー。破損ファイルがあっても他ファイルの処理を継続 |
| 3 | **XLSXパース本実装** | シート名を `heading`、内容をMarkdownテーブルとして `body_text` に格納。Shared Strings対応 |

### 🟡 優先度：中（精度・UX向上）

| # | 機能 | 概要 |
| :- | :--- | :--- |
| 4 | **画像メタデータ抽出** | `wp:docPr` の `name`/`descr` 属性を `Asset.title` に格納。AIが画像の意味を把握可能に |
| 5 | **箇条書きのMarkdown化** | `w:numPr` を検知し `- ` 形式に変換。リスト構造を保持 |
| 6 | **RAG最適化（context_path）** | 各セクションに親見出しのパスリストを付与。分割後の単体セクションでも文脈を保持 |
| 7 | **セクション単位のチャンク分割** | `--split` オプションで指定レベルごとに個別JSONを出力 |
| 8 | **画像リサイズ・圧縮** | `image` クレートで高解像度画像をリサイズ（例: 1024px, JPEG 80%）してからBase64化。JSON肥大化を防止 |

### 🟢 優先度：低（拡張・特殊対応）

| # | 機能 | 概要 |
| :- | :--- | :--- |
| 9 | **PPTXパース** | スライド単位で `Section` 化。テキストボックスを座標順に結合、スライドノートを補足コンテキストとして抽出 |
| 10 | **神エクセル対応** | セル結合解決、書式ベースの見出し判定、浮遊テキストボックス抽出 |
| 11 | **日本語テキスト正規化** | 全角空白のインデント解析、丸数字の正規化、OMML数式のLaTeX変換 |
| 12 | **見出し検出の柔軟化** | `config.json` のマッチングを前方一致・正規表現に対応 |
| 13 | **進捗表示（CLI）** | `indicatif` によるプログレスバー。大量ファイル処理時のUX向上 |

## 🛠 技術スタック
| カテゴリ | ライブラリ | 選定理由 |
| :--- | :--- | :--- |
| **Core** | `Rust` | メモリ安全、高速、クロスコンパイルの容易性 |
| **Parsing** | `zip`, `quick-xml` | Officeの実体(ZIP+XML)を高速・省メモリで処理 |
| **Parallel** | `rayon` | 複数ファイルの並列処理によるスループット向上 |
| **Networking** | `ureq` | 依存関係が極めて少なく、API連携に十分な機能を保持 |
| **Serialization**| `serde`, `serde_json` | 厳密な型定義に基づいた安全なJSON生成 |
| **CLI** | `clap` | 型安全なCLI引数パース |

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
    "見出し3": 3
  },
  "ppr_underline_as_heading": true,
  "run_underline_as_heading": false
}
```

### 設定項目

| キー | デフォルト | 説明 |
| :--- | :--- | :--- |
| `heading_styles` | 標準スタイル名セット | `スタイル名: レベル` のマッピング。Heading1〜3・見出し1〜3を既定で認識。 |
| `ppr_underline_as_heading` | `true` | 段落デフォルト書式（`w:pPr > w:rPr`）の下線を見出しとして扱う。 |
| `run_underline_as_heading` | `false` | ランレベル（`w:r > w:rPr`）の下線を見出しとして扱う。Wordの「見出し」スタイルを使わず直接書式で見出しを表現した文書向け。 |

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
