# ユニットテスト一覧（src/parser/docx.rs）

## determine_role() 系（10件）

| テスト名 | 内容 | 確認するアサーション |
|----------|------|----------------------|
| `test_determine_role_warning_variants` | Warning ロールの正常系 | Warning / WarningBox / CautionNote / AlertStyle / DangerZone → `Warning` |
| `test_determine_role_note_variants` | Note ロールの正常系 | Note / NoteBox / note-text → `Note` |
| `test_determine_role_no_false_positive_footnote` | **誤検知防止**: "note" を含む非Note スタイル | Footnote / FootnoteText / Annotation → `None` |
| `test_determine_role_no_false_positive_code` | **誤検知防止**: "code" を含む非Code スタイル | Barcode → `None` |
| `test_determine_role_code_block_variants` | CodeBlock ロールの正常系 | CodeBlock / PreFormat / Verbatim / SourceText → `CodeBlock` |
| `test_determine_role_tip_variants` | Tip ロールの正常系 | Tip / TipBox / HintStyle → `Tip` |
| `test_determine_role_quote_variants` | Quote ロールの正常系 | Quote / BlockQuote / Quotation → `Quote` |
| `test_determine_role_japanese_keywords` | 日本語キーワードのマッチング | 警告スタイル→`Warning` / 注意事項→`Note` / ヒント→`Tip` / 引用スタイル→`Quote` |
| `test_determine_role_custom_mapping` | config カスタムマッピングの優先評価 | `semantic_role_styles` に登録したスタイルが組み込みルールより優先されること |
| `test_determine_role_unknown_style` | 未知スタイルは `None` を返すこと | Normal / BodyText / ListParagraph → `None` |

## style_words() 系（2件）

| テスト名 | 内容 | 確認するアサーション |
|----------|------|----------------------|
| `test_style_words_camel_case` | CamelCase の単語分割 | WarningBox→["warning","box"] / Footnote→["footnote"] / FootnoteText→["footnote","text"] |
| `test_style_words_separators` | ハイフン・アンダースコア・スペースでの分割 | note-text→["note","text"] / code_block→["code","block"] / "warning style"→["warning","style"] |

## その他ロジック（3件）

| テスト名 | 内容 | 確認するアサーション |
|----------|------|----------------------|
| `test_is_ordered_numfmt` | 番号付きリスト判定 | decimal/lowerLetter/upperRoman → `true` / bullet/空文字 → `false` |
| `test_bookmark_filter_logic` | `bookmarkStart` フィルタリング条件 | `_Toc`・`_GoBack` で始まる名前がスキップ対象と判定されること |
| `test_outline_level_conversion` | outlineLvl の 0-based → 1-based 変換式 | raw=0 → 1、raw=2 → 3 |
