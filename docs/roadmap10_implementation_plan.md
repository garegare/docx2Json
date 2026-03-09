# ロードマップ #10「神エクセル対応」実装計画

作成日: 2026-03-09
担当: AI実装
対象ブランチ: feature/roadmap-10-god-excel

---

## 概要

「神エクセル」（セル結合・書式で見出しを表現・浮遊テキストボックスを多用する Excel ファイル）を正確に解析する。

### 実装する3機能

| # | 機能 | 効果 |
|---|------|------|
| A | **セル結合解決** | `<mergeCell ref="A1:C3">` を展開し、結合元の値を結合先セルにコピー。列数カウントが正確になり Markdown テーブルが崩れなくなる。 |
| B | **書式ベースの見出し判定** | `xl/styles.xml` を解析し、太字・背景色・大フォントサイズのセルを「見出し行」と判定。見出し行が現れるたびに新 Section を生成し、階層構造を再現する。 |
| C | **浮遊テキストボックス抽出** | `xl/drawings/drawing*.xml` の `<xdr:sp>` テキストを Section の body_text に追記する。 |

---

## 変更対象ファイル

- `src/config.rs` — `XlsxHeadingConfig` 構造体を追加
- `src/parser/xlsx.rs` — メイン実装（全面拡張）
- `README.md` — 新機能の説明を追記

---

## 実装ステップ詳細

### Step 1: `config.rs` に `XlsxHeadingConfig` を追加

```rust
/// XLSX 書式ベース見出し判定の設定（#10）
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XlsxHeadingConfig {
    /// 書式ベース見出し判定を有効にするか（デフォルト: false）
    /// false のままだと従来の「先頭行 = ヘッダー」モードを維持（後方互換性）
    #[serde(default)]
    pub enabled: bool,

    /// 太字セルを見出し条件にするか（デフォルト: true）
    #[serde(default = "default_true")]
    pub detect_bold: bool,

    /// 非白・非透明の背景色セルを見出し条件にするか（デフォルト: true）
    #[serde(default = "default_true")]
    pub detect_fill: bool,

    /// 見出し判定の最小フォントサイズ（pt）。0.0 = 無効（デフォルト: 0.0）
    #[serde(default)]
    pub heading_font_size_threshold: f32,

    /// 行内で「見出し書式」セルが占める割合の閾値（0.0〜1.0、デフォルト: 0.5）
    /// これ以上の割合で見出し書式なら行全体を見出し行と判定する
    #[serde(default = "default_heading_cell_ratio")]
    pub heading_cell_ratio: f32,
}

fn default_heading_cell_ratio() -> f32 { 0.5 }
```

`Config` 構造体に追加:
```rust
/// XLSX 書式ベース見出し判定（#10 神エクセル対応）
/// null または省略時は従来モード（先頭行ヘッダー）
#[serde(default)]
pub xlsx_heading: Option<XlsxHeadingConfig>,
```

---

### Step 2: `xlsx.rs` — 内部データ構造を追加

```rust
/// スタイルインデックス付きセル情報
struct CellInfo {
    value: String,
    style_idx: Option<usize>,  // xl/styles.xml の cellXfs インデックス
}

/// マージセル範囲（0-indexed）
struct MergeRange {
    min_row: usize,
    min_col: usize,
    max_row: usize,
    max_col: usize,
}

/// parse_worksheet の戻り値（拡張版）
struct WorksheetData {
    cells: HashMap<(usize, usize), CellInfo>,
    max_row: usize,
    max_col: usize,
    merges: Vec<MergeRange>,
}

/// 解決済みセルスタイル情報
#[derive(Default, Clone)]
struct CellStyleInfo {
    bold: bool,
    font_size: f32,   // pt (0.0 = 不明)
    has_fill: bool,   // 非白・非透明の背景色あり
}

/// xl/styles.xml から読み取ったスタイルテーブル
struct XlsxStyles {
    /// cellXfs[i] → 解決済みスタイル情報
    cell_styles: Vec<CellStyleInfo>,
}
```

---

### Step 3: `parse_worksheet` を拡張（`WorksheetData` を返す）

**現在の戻り値**: `Result<Vec<Vec<String>>>`
**変更後の戻り値**: `Result<WorksheetData>`

追加する解析ロジック:

1. `<c>` 要素の `s` 属性 → `style_idx` として `CellInfo` に格納
2. `<mergeCells>` セクションの `<mergeCell ref="A1:C3">` を収集 → `merges` に格納

```xml
<!-- sheet*.xml の構造例 -->
<mergeCells count="3">
  <mergeCell ref="A1:C1"/>
  <mergeCell ref="A2:A5"/>
  <mergeCell ref="D2:F5"/>
</mergeCells>
```

`parse_merge_range("A1:C1")` ヘルパー:
- `":"` で分割し、各セルアドレスを行・列インデックスに変換
- `parse_cell_address("A1")` → `(row=0, col=0)`

---

### Step 4: `apply_merge_cells` — セル結合展開

```
結合元セル (min_row, min_col) の値・スタイルを
結合範囲全体 (min_row..=max_row, min_col..=max_col) にコピーする
```

処理順:
1. `WorksheetData.cells` を可変借用
2. 各 `MergeRange` の結合元の値・スタイルを取得
3. 結合元以外の全セルに同じ値・スタイルを書き込み（既存値は上書き）

---

### Step 5: `parse_styles` — `xl/styles.xml` 解析

XLSX スタイル解決の参照チェーン:

```
cellXfs[i]
  └─ fontId → fonts[fontId]
  │    └─ <b/>       → bold = true
  │    └─ <sz val="14"/> → font_size = 14.0
  └─ fillId → fills[fillId]
       └─ <patternFill patternType="solid">
            └─ <fgColor rgb="FFFF0000"/> → has_fill = true
```

注意事項:
- `fills[0]` は "none"、`fills[1]` は "gray125" が標準デフォルトなので `fillId >= 2` のみ有効
- `fgColor theme="..."` の場合も `has_fill = true` とする（色値の解決は不要）
- `indexed` カラーは `has_fill = true` とする

---

### Step 6: `classify_row` — 行分類

```rust
enum RowKind { Heading, Data }

fn classify_row(
    row: &[Option<&CellInfo>],  // 列数分のスライス（空セルは None）
    styles: &XlsxStyles,
    cfg: &XlsxHeadingConfig,
) -> RowKind
```

ロジック:
1. 値が空でないセルを `non_empty` として収集（空行は Data）
2. `non_empty` のうち「見出し書式」を持つセルの個数をカウント
3. 割合が `cfg.heading_cell_ratio` 以上 → `Heading`、未満 → `Data`

「見出し書式」の条件（いずれかが true）:
- `cfg.detect_bold && style.bold`
- `cfg.detect_fill && style.has_fill`
- `cfg.heading_font_size_threshold > 0.0 && style.font_size >= cfg.heading_font_size_threshold`

---

### Step 7: `sheet_to_section_with_headings` — 書式ベース Section 生成

**入力**: 密な 2D `Vec<Vec<CellInfo>>`、`XlsxStyles`、`XlsxHeadingConfig`、`max_rows`
**出力**: `Section`（children を持つ）

アルゴリズム:

```
シート全体を走査:
  見出し行 → 現在のグループを確定、新グループを開始
  データ行 → 現在のグループに追加

グループ = { heading_cells: Vec<String>, data_rows: Vec<Vec<String>> }
           ↓
Section { heading: 見出し行の値を結合, body_text: data_rows の Markdown 表 }
```

見出し行が一度も出現しない場合は従来の `sheet_to_section_flat` にフォールバック。
`max_rows` によるチャンク分割は各 Section の data_rows に対して適用。

---

### Step 8: `parse_sheet_rels` + `parse_drawing` — 浮遊テキストボックス

#### `parse_sheet_rels`

```
xl/worksheets/_rels/sheet1.xml.rels
→ Type が "...drawing" のリレーションシップを抽出
→ Target: "../drawings/drawing1.xml" → "xl/drawings/drawing1.xml" に解決
```

#### `parse_drawing`

```xml
<!-- drawing1.xml の構造 -->
<xdr:wsDr>
  <xdr:twoCellAnchor>
    <xdr:sp>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:t>テキストボックスの内容</a:t></a:r>
        </a:p>
      </xdr:txBody>
    </xdr:sp>
  </xdr:twoCellAnchor>
</xdr:wsDr>
```

`<xdr:sp>` 内の `<a:t>` テキストを収集し、スペースで結合して返す。

---

### Step 9: 新 `sheet_to_section` — オーケストレーター

現在の `sheet_to_section` を `sheet_to_section_flat` に改名し、
新しい `sheet_to_section` が全処理を統括する:

```rust
fn sheet_to_section(
    name: &str,
    mut data: WorksheetData,
    styles: &XlsxStyles,
    drawing_texts: Vec<String>,
    config: &Config,
) -> Section {
    // 1. セル結合展開
    apply_merge_cells(&mut data.cells, &data.merges);

    // 2. 密グリッドに変換
    let grid = to_dense_grid(&data);

    // 3. Section 生成（書式ベース or フラット）
    let mut section = if let Some(ref hcfg) = config.xlsx_heading {
        if hcfg.enabled {
            sheet_to_section_with_headings(name, grid, styles, hcfg, config.xlsx_max_rows)
        } else {
            sheet_to_section_flat(name, to_string_grid(&grid), config.xlsx_max_rows)
        }
    } else {
        sheet_to_section_flat(name, to_string_grid(&grid), config.xlsx_max_rows)
    };

    // 4. 浮遊テキストボックスを body_text に追記
    if !drawing_texts.is_empty() {
        let drawings_md = drawing_texts.join("\n\n");
        if section.body_text.is_empty() {
            section.body_text = drawings_md;
        } else {
            section.body_text.push_str("\n\n---\n\n");
            section.body_text.push_str(&drawings_md);
        }
    }

    section
}
```

---

### Step 10: `parse` エントリポイントを更新

```rust
// 追加: styles を一度だけ解析
let styles = parse_styles(&mut archive).unwrap_or_default();

// 各シートの処理を更新
match parse_worksheet(&mut archive, &sheet_path, &shared_strings) {
    Ok(data) => {
        // 浮遊テキストボックス
        let drawing_texts = parse_sheet_rels(&mut archive, sheet_idx)
            .and_then(|p| parse_drawing(&mut archive, &p))
            .unwrap_or_default();

        let section = sheet_to_section(name, data, &styles, drawing_texts, config);
        sections.push(section);
    }
    ...
}
```

---

### Step 11: `cargo build` & 動作確認

```bash
# ビルド（警告ゼロを確認）
cargo build 2>&1

# タイムスタンプ付きテスト実行
TS=$(date +%Y%m%d_%H%M%S)
CMD="cargo run -- --input samples/tb_r2fu_99_digi_a.xlsx --output out/$TS"
mkdir -p out/$TS
{
  echo "Date: $(date)"
  echo "Git: $(git rev-parse --short HEAD) ($(git branch --show-current))"
  echo "Command: $CMD"
} > out/$TS/run_info.txt
eval $CMD

# 書式ベース見出しモードでのテスト（docx2json.json に設定追記後）
TS2=$(date +%Y%m%d_%H%M%S)
CMD2="cargo run -- --input samples/tb_r2fu_99_digi_a.xlsx --output out/$TS2"
mkdir -p out/$TS2
{
  echo "Date: $(date)"
  echo "Git: $(git rev-parse --short HEAD) ($(git branch --show-current))"
  echo "Command: $CMD2"
} > out/$TS2/run_info.txt
eval $CMD2
```

---

## 設定ファイル拡張例

```json
{
  "xlsx_max_rows": 100,
  "xlsx_heading": {
    "enabled": true,
    "detect_bold": true,
    "detect_fill": true,
    "heading_font_size_threshold": 0.0,
    "heading_cell_ratio": 0.5
  }
}
```

---

## 後方互換性

- `xlsx_heading` が設定ファイルに存在しない（または `enabled: false`）場合は従来モード
  - 先頭行 = ヘッダー行として Markdown テーブルを生成
  - セル結合展開は**常に**実行（壊れたテーブルを修正する副次効果があるため）
  - 浮遊テキストボックスは**常に**body_text に追記

---

## ファイル構成（変更後）

```
src/
  config.rs           ← XlsxHeadingConfig 追加
  parser/
    xlsx.rs           ← 全面拡張（約600行→800行程度）
```
