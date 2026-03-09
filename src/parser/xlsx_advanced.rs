/// 神エクセル対応 XLSX パーサー（ロードマップ #10）
///
/// `config.xlsx_heading.enabled == true` のときのみ呼ばれる。
/// 既存の `xlsx.rs` は一切変更しない。
///
/// ## 実装する3機能
/// - **A: セル結合解決** — `<mergeCell>` を展開し結合元の値を全セルにコピー（Phase 1）
/// - **B: 書式ベースの見出し判定** — `xl/styles.xml` の太字・背景色行を見出しに昇格（Phase 2）
/// - **C: 浮遊テキストボックス抽出** — `xl/drawings/drawing*.xml` のテキスト抽出（Phase 3）
use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::Path;

use anyhow::{Context, Result};
use quick_xml::events::Event;
use quick_xml::Reader;
use zip::ZipArchive;

use crate::config::{Config, XlsxHeadingConfig};
use crate::models::{Document, Section};

// ============================================================
// 内部データ構造
// ============================================================

/// スタイルインデックス付きセル情報
struct CellInfo {
    value: String,
    /// xl/styles.xml の cellXfs インデックス（`<c s="...">` の s 属性）
    style_idx: Option<usize>,
}

/// マージセル範囲（0-indexed、両端含む）
struct MergeRange {
    min_row: usize,
    min_col: usize,
    max_row: usize,
    max_col: usize,
}

/// `parse_worksheet` の拡張戻り値
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
    font_size: f32, // pt (0.0 = 不明)
    has_fill: bool, // 非白・非透明の背景色あり
}

/// `xl/styles.xml` から読み取ったスタイルテーブル
#[derive(Default)]
struct XlsxStyles {
    /// cellXfs[i] → 解決済みスタイル情報
    cell_styles: Vec<CellStyleInfo>,
}

/// 行の分類結果
enum RowKind {
    Heading, // 書式ベース見出し行
    Data,    // データ行
}

// ============================================================
// エントリポイント
// ============================================================

/// 神エクセル対応 XLSX パーサーのエントリポイント
pub fn parse(path: &Path, config: &Config) -> Result<Document> {
    let title = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("Untitled")
        .to_string();

    let file = File::open(path)
        .with_context(|| format!("ファイルを開けません: {}", path.display()))?;
    let mut archive = ZipArchive::new(file).context("ZIPアーカイブとして開けません")?;

    // 1. ワークブックのリレーションシップ（rId → シートパス）
    let rels = parse_workbook_rels(&mut archive).unwrap_or_default();

    // 2. シート名と rId のリスト
    let sheets = parse_workbook(&mut archive)?;

    // 3. 共有文字列テーブル
    let shared_strings = parse_shared_strings(&mut archive).unwrap_or_default();

    // 4. スタイルテーブル（xl/styles.xml → cellXfs → fonts/fills を解決）
    let styles = parse_styles(&mut archive).unwrap_or_default();

    // 5. 各シートを Section に変換
    let mut sections = Vec::new();
    for (sheet_idx, (name, rid)) in sheets.iter().enumerate() {
        let target = match rels.get(rid.as_str()) {
            Some(t) => t.clone(),
            None => {
                eprintln!(
                    "Warning: シート '{}' (rId={}) のパスが見つかりません",
                    name, rid
                );
                continue;
            }
        };

        let sheet_path = resolve_path(&target);

        match parse_worksheet(&mut archive, &sheet_path, &shared_strings) {
            Ok(data) => {
                // 浮遊テキストボックス（Phase 3 で本実装）
                let drawing_texts =
                    parse_sheet_drawings(&mut archive, &sheet_path, sheet_idx)
                        .unwrap_or_default();

                let section = build_section(name, data, &styles, drawing_texts, config);
                sections.push(section);
            }
            Err(e) => {
                eprintln!("Warning: シート '{}' の解析に失敗しました: {e}", name);
            }
        }
    }

    Ok(Document { title, sections })
}

// ============================================================
// Section 構築
// ============================================================

/// WorksheetData から Section を構築するオーケストレーター
fn build_section(
    name: &str,
    mut data: WorksheetData,
    styles: &XlsxStyles,
    drawing_texts: Vec<String>,
    config: &Config,
) -> Section {
    // A: セル結合展開（常に実行）
    apply_merge_cells(&mut data);

    // B: Section 生成（書式ベース or 従来フラット）
    let mut section = if config.xlsx_heading.as_ref().map_or(false, |h| h.enabled) {
        let hcfg = config.xlsx_heading.as_ref().unwrap();
        build_section_with_headings(name, &data, styles, hcfg, config.xlsx_max_rows)
    } else {
        let grid = to_dense_grid(&data);
        build_section_flat(name, grid, config.xlsx_max_rows)
    };

    // C: 浮遊テキストボックスを body_text に追記（常に実行）
    if !drawing_texts.is_empty() {
        let drawings_text = drawing_texts.join("\n\n");
        if section.body_text.is_empty() {
            section.body_text = drawings_text;
        } else {
            section.body_text.push_str("\n\n---\n\n");
            section.body_text.push_str(&drawings_text);
        }
    }

    section
}

/// 従来モード: 先頭行ヘッダー・Markdown テーブルのフラット Section
///
/// xlsx.rs の `sheet_to_section` と同等の動作をする。
fn build_section_flat(name: &str, grid: Vec<Vec<String>>, max_rows: usize) -> Section {
    if grid.is_empty() {
        return Section {
            context_path: Vec::new(),
            heading: name.to_string(),
            body_text: String::new(),
            assets: Vec::new(),
            children: Vec::new(),
        };
    }

    let data_row_count = grid.len().saturating_sub(1);

    if max_rows == 0 || data_row_count <= max_rows {
        return Section {
            context_path: Vec::new(),
            heading: name.to_string(),
            body_text: grid_to_markdown(&grid),
            assets: Vec::new(),
            children: Vec::new(),
        };
    }

    // 行数超過: ヘッダー保持で子 Section に分割
    let header = grid[0].clone();
    let data_rows = &grid[1..];
    let chunk_count = (data_row_count + max_rows - 1) / max_rows;

    let children: Vec<Section> = data_rows
        .chunks(max_rows)
        .enumerate()
        .map(|(i, chunk)| {
            let start = i * max_rows + 1;
            let end = start + chunk.len() - 1;
            let mut child_rows = Vec::with_capacity(chunk.len() + 1);
            child_rows.push(header.clone());
            child_rows.extend_from_slice(chunk);
            Section {
                context_path: Vec::new(),
                heading: format!("{} [行 {}–{}]", name, start, end),
                body_text: grid_to_markdown(&child_rows),
                assets: Vec::new(),
                children: Vec::new(),
            }
        })
        .collect();

    Section {
        context_path: Vec::new(),
        heading: name.to_string(),
        body_text: format!(
            "（全 {} 行 / {} 行ずつ {} チャンクに分割）",
            data_row_count, max_rows, chunk_count
        ),
        assets: Vec::new(),
        children,
    }
}

/// 書式ベース見出し判定モード（Phase 2 本実装）
///
/// 太字・背景色などの書式を持つ行を見出しと判定し、
/// 見出しが現れるたびに新しい子 Section を作成する。
///
/// 見出し行が一度も出現しない場合は従来フラットモードにフォールバックする。
fn build_section_with_headings(
    name: &str,
    data: &WorksheetData,
    styles: &XlsxStyles,
    hcfg: &XlsxHeadingConfig,
    max_rows: usize,
) -> Section {
    // ヘッダー前のデータ行（最初の見出し行より前にある行）
    let mut pre_heading_rows: Vec<Vec<String>> = Vec::new();
    // (見出しテキスト, データ行リスト) のグループリスト
    let mut groups: Vec<(String, Vec<Vec<String>>)> = Vec::new();
    let mut current_heading: Option<String> = None;
    let mut current_data: Vec<Vec<String>> = Vec::new();
    let mut has_any_heading = false;

    for r in 0..data.max_row {
        // 行の値を収集
        let row_values: Vec<String> = (0..data.max_col)
            .map(|c| {
                data.cells
                    .get(&(r, c))
                    .map_or(String::new(), |cell| cell.value.clone())
            })
            .collect();

        // 完全空行はスキップ
        if row_values.iter().all(|v| v.is_empty()) {
            continue;
        }

        match classify_row(r, data, styles, hcfg) {
            RowKind::Heading => {
                // 直前のグループを確定
                if let Some(h) = current_heading.take() {
                    groups.push((h, std::mem::take(&mut current_data)));
                } else if !pre_heading_rows.is_empty() {
                    // 見出し前のデータ行は空見出しグループとして扱う
                    groups.push((String::new(), std::mem::take(&mut pre_heading_rows)));
                }
                // 見出しテキスト: 空でないセルをスペースで結合
                let heading_text = row_values
                    .iter()
                    .filter(|v| !v.is_empty())
                    .cloned()
                    .collect::<Vec<_>>()
                    .join(" ");
                current_heading = Some(heading_text);
                has_any_heading = true;
            }
            RowKind::Data => {
                if current_heading.is_none() {
                    pre_heading_rows.push(row_values);
                } else {
                    current_data.push(row_values);
                }
            }
        }
    }

    // 最終グループを確定
    if let Some(h) = current_heading {
        groups.push((h, current_data));
    }

    // 見出しが一切なければ従来フラットモードにフォールバック
    if !has_any_heading {
        let grid = to_dense_grid(data);
        return build_section_flat(name, grid, max_rows);
    }

    // 親 Section の body_text = 見出し前のデータ（通常は空）
    let parent_body = if pre_heading_rows.is_empty() {
        String::new()
    } else {
        grid_to_markdown(&pre_heading_rows)
    };

    // 各グループを子 Section に変換
    let children: Vec<Section> = groups
        .into_iter()
        .filter(|(h, rows)| !h.is_empty() || !rows.is_empty())
        .map(|(heading, data_rows)| {
            // max_rows による子 Section のチャンク分割
            if max_rows == 0 || data_rows.len() <= max_rows {
                Section {
                    context_path: Vec::new(),
                    heading,
                    body_text: grid_to_markdown(&data_rows),
                    assets: Vec::new(),
                    children: Vec::new(),
                }
            } else {
                // 行数超過: チャンク分割して孫 Section を生成
                let chunk_count = (data_rows.len() + max_rows - 1) / max_rows;
                let chunk_children: Vec<Section> = data_rows
                    .chunks(max_rows)
                    .enumerate()
                    .map(|(i, chunk)| {
                        let start = i * max_rows + 1;
                        let end = start + chunk.len() - 1;
                        Section {
                            context_path: Vec::new(),
                            heading: format!("{} [行 {}–{}]", heading, start, end),
                            body_text: grid_to_markdown(chunk),
                            assets: Vec::new(),
                            children: Vec::new(),
                        }
                    })
                    .collect();
                Section {
                    context_path: Vec::new(),
                    heading,
                    body_text: format!(
                        "（全 {} 行 / {} 行ずつ {} チャンクに分割）",
                        data_rows.len(),
                        max_rows,
                        chunk_count
                    ),
                    assets: Vec::new(),
                    children: chunk_children,
                }
            }
        })
        .collect();

    Section {
        context_path: Vec::new(),
        heading: name.to_string(),
        body_text: parent_body,
        assets: Vec::new(),
        children,
    }
}

/// 行を Heading / Data に分類する
///
/// 行内の空でないセルのうち、「見出し書式」を持つセルの割合が
/// `hcfg.heading_cell_ratio` 以上なら `Heading`、未満なら `Data`。
///
/// 見出し書式の条件（いずれかが true）:
/// - `detect_bold && style.bold`
/// - `detect_fill && style.has_fill`
/// - `heading_font_size_threshold > 0 && style.font_size >= threshold`
fn classify_row(
    row_idx: usize,
    data: &WorksheetData,
    styles: &XlsxStyles,
    hcfg: &XlsxHeadingConfig,
) -> RowKind {
    let non_empty_cells: Vec<&CellInfo> = (0..data.max_col)
        .filter_map(|c| data.cells.get(&(row_idx, c)))
        .filter(|cell| !cell.value.is_empty())
        .collect();

    if non_empty_cells.is_empty() {
        return RowKind::Data;
    }

    let heading_count = non_empty_cells
        .iter()
        .filter(|cell| {
            let style = cell
                .style_idx
                .and_then(|i| styles.cell_styles.get(i))
                .cloned()
                .unwrap_or_default();
            (hcfg.detect_bold && style.bold)
                || (hcfg.detect_fill && style.has_fill)
                || (hcfg.heading_font_size_threshold > 0.0
                    && style.font_size >= hcfg.heading_font_size_threshold)
        })
        .count();

    let ratio = heading_count as f32 / non_empty_cells.len() as f32;
    if ratio >= hcfg.heading_cell_ratio {
        RowKind::Heading
    } else {
        RowKind::Data
    }
}

// ============================================================
// セル結合展開（A: Phase 1 実装済み）
// ============================================================

/// セル結合を展開する: 結合元セルの値・スタイルを結合範囲全体にコピーする
fn apply_merge_cells(data: &mut WorksheetData) {
    for m in &data.merges {
        let (origin_val, origin_style) = data
            .cells
            .get(&(m.min_row, m.min_col))
            .map(|c| (c.value.clone(), c.style_idx))
            .unwrap_or_default();

        for r in m.min_row..=m.max_row {
            for c in m.min_col..=m.max_col {
                if r == m.min_row && c == m.min_col {
                    continue;
                }
                data.cells.entry((r, c)).or_insert_with(|| CellInfo {
                    value: origin_val.clone(),
                    style_idx: origin_style,
                });
            }
        }
    }
}

/// `WorksheetData` の sparse cells を密な 2D Vec<Vec<String>> に変換する
fn to_dense_grid(data: &WorksheetData) -> Vec<Vec<String>> {
    if data.max_row == 0 || data.max_col == 0 {
        return Vec::new();
    }
    let mut grid = vec![vec![String::new(); data.max_col]; data.max_row];
    for ((r, c), cell) in &data.cells {
        if *r < data.max_row && *c < data.max_col {
            grid[*r][*c] = cell.value.clone();
        }
    }
    grid
}

// ============================================================
// XML パーサー群
// ============================================================

/// `xl/_rels/workbook.xml.rels` を解析して `rId → Target` の Map を返す
fn parse_workbook_rels(archive: &mut ZipArchive<File>) -> Result<HashMap<String, String>> {
    let content = read_zip_entry(archive, "xl/_rels/workbook.xml.rels")?;
    let mut map = HashMap::new();
    let mut reader = Reader::from_str(&content);
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event()? {
            Event::Empty(e) | Event::Start(e)
                if e.local_name().as_ref() == b"Relationship" =>
            {
                let mut id = String::new();
                let mut target = String::new();
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"Id" => id = decode_bytes(&attr.value),
                        b"Target" => target = decode_bytes(&attr.value),
                        _ => {}
                    }
                }
                if !id.is_empty() && !target.is_empty() {
                    map.insert(id, target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(map)
}

/// `xl/workbook.xml` を解析してシートのリスト `(name, rId)` を順序付きで返す
fn parse_workbook(archive: &mut ZipArchive<File>) -> Result<Vec<(String, String)>> {
    let content = read_zip_entry(archive, "xl/workbook.xml")?;
    let mut sheets = Vec::new();
    let mut reader = Reader::from_str(&content);
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event()? {
            Event::Empty(e) | Event::Start(e) if e.local_name().as_ref() == b"sheet" => {
                let mut name = String::new();
                let mut rid = String::new();
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"name" => name = decode_bytes(&attr.value),
                        b"id" => rid = decode_bytes(&attr.value),
                        _ => {}
                    }
                }
                if !name.is_empty() && !rid.is_empty() {
                    sheets.push((name, rid));
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(sheets)
}

/// `xl/sharedStrings.xml` を解析して共有文字列テーブルを返す
fn parse_shared_strings(archive: &mut ZipArchive<File>) -> Result<Vec<String>> {
    let content = read_zip_entry(archive, "xl/sharedStrings.xml")?;
    let mut strings = Vec::new();
    let mut reader = Reader::from_str(&content);
    reader.config_mut().trim_text(false);
    let mut in_si = false;
    let mut current = String::new();
    loop {
        match reader.read_event()? {
            Event::Start(e) if e.local_name().as_ref() == b"si" => {
                in_si = true;
                current.clear();
            }
            Event::End(e) if e.local_name().as_ref() == b"si" => {
                strings.push(current.trim().to_string());
                in_si = false;
            }
            Event::Text(e) if in_si => {
                current.push_str(&e.unescape().unwrap_or_default());
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(strings)
}

/// `xl/worksheets/sheet*.xml` を解析して `WorksheetData` を返す
///
/// 既存の `xlsx.rs::parse_worksheet` との差分:
/// - `<c>` の `s` 属性を `style_idx` として収集
/// - `<mergeCells>` の `<mergeCell ref="...">` を収集
fn parse_worksheet(
    archive: &mut ZipArchive<File>,
    path: &str,
    shared_strings: &[String],
) -> Result<WorksheetData> {
    let content = read_zip_entry(archive, path)?;
    let mut reader = Reader::from_str(&content);
    reader.config_mut().trim_text(true);

    let mut cells: HashMap<(usize, usize), CellInfo> = HashMap::new();
    let mut max_row = 0usize;
    let mut max_col = 0usize;
    let mut merges: Vec<MergeRange> = Vec::new();

    let mut cur_row = 0usize;
    let mut cur_col = 0usize;
    let mut cell_type = String::new();
    let mut cell_style: Option<usize> = None;
    let mut in_v = false;
    let mut in_t = false;
    let mut cell_buf = String::new();

    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) => {
                match e.local_name().as_ref() {
                    b"row" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"r" {
                                let row_num: usize =
                                    decode_bytes(&attr.value).parse().unwrap_or(1);
                                cur_row = row_num.saturating_sub(1);
                                max_row = max_row.max(cur_row + 1);
                            }
                        }
                    }
                    b"c" => {
                        cell_type.clear();
                        cell_style = None;
                        cell_buf.clear();
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"r" => {
                                    let cell_ref = decode_bytes(&attr.value);
                                    cur_col = col_index(&cell_ref);
                                    max_col = max_col.max(cur_col + 1);
                                }
                                b"t" => cell_type = decode_bytes(&attr.value),
                                b"s" => {
                                    cell_style = decode_bytes(&attr.value).parse().ok();
                                }
                                _ => {}
                            }
                        }
                    }
                    b"v" => in_v = true,
                    b"t" => in_t = true,
                    b"mergeCell" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"ref" {
                                let ref_str = decode_bytes(&attr.value);
                                if let Some(mr) = parse_merge_range(&ref_str) {
                                    merges.push(mr);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => match e.local_name().as_ref() {
                b"v" => in_v = false,
                b"t" => in_t = false,
                b"c" => {
                    let val = resolve_cell_value(&cell_type, &cell_buf, shared_strings);
                    if !val.is_empty() {
                        cells.insert(
                            (cur_row, cur_col),
                            CellInfo {
                                value: val,
                                style_idx: cell_style,
                            },
                        );
                    }
                }
                _ => {}
            },
            Event::Text(e) if in_v || in_t => {
                cell_buf.push_str(&e.unescape().unwrap_or_default());
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(WorksheetData {
        cells,
        max_row,
        max_col,
        merges,
    })
}

/// `xl/styles.xml` を解析してスタイルテーブルを返す（Phase 2 本実装）
///
/// 参照チェーン: `cellXfs[i].fontId` → `fonts[fontId]` (bold, font_size)
///              `cellXfs[i].fillId` → `fills[fillId]` (has_fill)
fn parse_styles(archive: &mut ZipArchive<File>) -> Result<XlsxStyles> {
    let content = match read_zip_entry(archive, "xl/styles.xml") {
        Ok(c) => c,
        Err(_) => return Ok(XlsxStyles::default()),
    };

    let mut reader = Reader::from_str(&content);
    reader.config_mut().trim_text(true);

    // 中間データ
    struct FontRec {
        bold: bool,
        font_size: f32,
    }
    struct FillRec {
        has_fill: bool,
    }

    let mut fonts: Vec<FontRec> = Vec::new();
    let mut fills: Vec<FillRec> = Vec::new();
    let mut xf_records: Vec<(usize, usize)> = Vec::new(); // (font_id, fill_id)

    // 状態フラグ
    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_cell_xfs = false;
    let mut in_font = false;
    let mut in_fill = false;
    let mut in_xf = false; // <xf> が Start で開始された（End で push）

    // 現在処理中のフォント情報
    let mut cur_bold = false;
    let mut cur_font_size = 0.0f32;

    // 現在処理中のフィル情報
    let mut cur_pattern_type = String::new();
    let mut cur_has_fgcolor = false;

    // 現在処理中の cellXf 情報（Start 型 <xf> 用）
    let mut cur_xf_font_id = 0usize;
    let mut cur_xf_fill_id = 0usize;

    loop {
        match reader.read_event()? {
            // ---- Start イベント ----
            Event::Start(e) => match e.local_name().as_ref() {
                b"fonts" => in_fonts = true,
                b"fills" => in_fills = true,
                b"cellXfs" => in_cell_xfs = true,
                b"font" if in_fonts => {
                    in_font = true;
                    cur_bold = false;
                    cur_font_size = 0.0;
                }
                b"fill" if in_fills => {
                    in_fill = true;
                    cur_pattern_type.clear();
                    cur_has_fgcolor = false;
                }
                b"patternFill" if in_fill => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"patternType" {
                            cur_pattern_type = decode_bytes(&attr.value);
                        }
                    }
                }
                b"fgColor" if in_fill => {
                    // fgColor に rgb / theme / indexed のいずれかがあれば色あり
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"rgb" | b"theme" | b"indexed" => cur_has_fgcolor = true,
                            _ => {}
                        }
                    }
                }
                b"xf" if in_cell_xfs => {
                    cur_xf_font_id = 0;
                    cur_xf_fill_id = 0;
                    in_xf = true;
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"fontId" => {
                                cur_xf_font_id = decode_bytes(&attr.value).parse().unwrap_or(0)
                            }
                            b"fillId" => {
                                cur_xf_fill_id = decode_bytes(&attr.value).parse().unwrap_or(0)
                            }
                            _ => {}
                        }
                    }
                }
                // <b> タグが Start で書かれる場合（<b></b> 形式）
                b"b" if in_font => cur_bold = true,
                _ => {}
            },

            // ---- Empty イベント（自己終了タグ）----
            Event::Empty(e) => match e.local_name().as_ref() {
                b"patternFill" if in_fill => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"patternType" {
                            cur_pattern_type = decode_bytes(&attr.value);
                        }
                    }
                }
                b"fgColor" if in_fill => {
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"rgb" | b"theme" | b"indexed" => cur_has_fgcolor = true,
                            _ => {}
                        }
                    }
                }
                // <b/> : 太字（最も一般的な形式）
                b"b" if in_font => {
                    // <b val="0"/> は非太字
                    let mut is_bold = true;
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"val"
                            && decode_bytes(&attr.value) == "0"
                        {
                            is_bold = false;
                        }
                    }
                    cur_bold = is_bold;
                }
                // <sz val="N"/> : フォントサイズ
                b"sz" if in_font => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"val" {
                            cur_font_size = decode_bytes(&attr.value).parse().unwrap_or(0.0);
                        }
                    }
                }
                // <xf .../> : cellXfs 内の自己終了型 cell format
                b"xf" if in_cell_xfs => {
                    let mut font_id = 0usize;
                    let mut fill_id = 0usize;
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"fontId" => font_id = decode_bytes(&attr.value).parse().unwrap_or(0),
                            b"fillId" => fill_id = decode_bytes(&attr.value).parse().unwrap_or(0),
                            _ => {}
                        }
                    }
                    xf_records.push((font_id, fill_id));
                }
                _ => {}
            },

            // ---- End イベント ----
            Event::End(e) => match e.local_name().as_ref() {
                b"fonts" => in_fonts = false,
                b"fills" => in_fills = false,
                b"cellXfs" => in_cell_xfs = false,
                b"font" if in_fonts => {
                    fonts.push(FontRec {
                        bold: cur_bold,
                        font_size: cur_font_size,
                    });
                    in_font = false;
                }
                b"fill" if in_fills => {
                    // patternType == "solid" かつ fgColor 属性あり → 有色背景
                    let has_fill = cur_pattern_type == "solid" && cur_has_fgcolor;
                    fills.push(FillRec { has_fill });
                    in_fill = false;
                }
                b"xf" if in_xf && in_cell_xfs => {
                    xf_records.push((cur_xf_font_id, cur_xf_fill_id));
                    in_xf = false;
                }
                _ => {}
            },

            Event::Eof => break,
            _ => {}
        }
    }

    // cellXfs の各エントリを解決済み CellStyleInfo に変換
    let cell_styles = xf_records
        .iter()
        .map(|&(font_id, fill_id)| CellStyleInfo {
            bold: fonts.get(font_id).map_or(false, |f| f.bold),
            font_size: fonts.get(font_id).map_or(0.0, |f| f.font_size),
            has_fill: fills.get(fill_id).map_or(false, |f| f.has_fill),
        })
        .collect();

    Ok(XlsxStyles { cell_styles })
}

/// シートに関連付けられた Drawing ファイルからテキストを抽出する（Phase 3 で本実装）
///
/// Phase 2 では空ベクターを返すスタブ。
fn parse_sheet_drawings(
    _archive: &mut ZipArchive<File>,
    _sheet_path: &str,
    _sheet_idx: usize,
) -> Result<Vec<String>> {
    // TODO Phase 3: xl/worksheets/_rels/sheet*.xml.rels → xl/drawings/drawing*.xml を解析
    Ok(Vec::new())
}

// ============================================================
// Markdown 生成
// ============================================================

/// セルグリッドを Markdown 表に変換する（xlsx.rs と同一ロジック）
fn grid_to_markdown(rows: &[Vec<String>]) -> String {
    if rows.is_empty() {
        return String::new();
    }
    let col_count = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    if col_count == 0 {
        return String::new();
    }
    let mut out = String::new();
    for (i, row) in rows.iter().enumerate() {
        let cells: Vec<String> = (0..col_count)
            .map(|c| row.get(c).map(|s| escape_cell(s)).unwrap_or_default())
            .collect();
        out.push_str(&format!("| {} |\n", cells.join(" | ")));
        if i == 0 {
            let sep = vec!["---"; col_count];
            out.push_str(&format!("|{}|\n", sep.join("|")));
        }
    }
    out
}

fn escape_cell(s: &str) -> String {
    s.replace('|', "\\|")
        .replace('\n', " ")
        .replace('\r', "")
}

// ============================================================
// ユーティリティ
// ============================================================

/// セル参照の列部分（英字）をゼロ始まりの列インデックスに変換する
fn col_index(cell_ref: &str) -> usize {
    cell_ref
        .chars()
        .take_while(|c| c.is_ascii_alphabetic())
        .fold(0usize, |acc, c| {
            acc * 26 + (c.to_ascii_uppercase() as usize - b'A' as usize + 1)
        })
        .saturating_sub(1)
}

/// セルアドレス（"A1"）を (row, col) の 0-indexed タプルに変換する
fn parse_cell_address(addr: &str) -> Option<(usize, usize)> {
    let col_str: String = addr.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let row_str: String = addr
        .chars()
        .skip_while(|c| c.is_ascii_alphabetic())
        .collect();
    if col_str.is_empty() || row_str.is_empty() {
        return None;
    }
    let col = col_str
        .chars()
        .fold(0usize, |acc, c| {
            acc * 26 + (c.to_ascii_uppercase() as usize - b'A' as usize + 1)
        })
        .saturating_sub(1);
    let row: usize = row_str.parse::<usize>().ok()?.saturating_sub(1);
    Some((row, col))
}

/// "A1:C3" 形式のセル範囲文字列を `MergeRange` に変換する
fn parse_merge_range(ref_str: &str) -> Option<MergeRange> {
    let mut parts = ref_str.splitn(2, ':');
    let start = parse_cell_address(parts.next()?)?;
    let end = parse_cell_address(parts.next()?)?;
    Some(MergeRange {
        min_row: start.0,
        min_col: start.1,
        max_row: end.0,
        max_col: end.1,
    })
}

/// `xl/_rels/workbook.xml.rels` の Target パスを ZIP 内絶対パスに解決する
fn resolve_path(target: &str) -> String {
    if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        format!("xl/{}", target)
    }
}

/// バイト列を UTF-8 文字列にデコードする
fn decode_bytes(val: &[u8]) -> String {
    String::from_utf8_lossy(val).into_owned()
}

/// ZIPアーカイブから指定パスのエントリを文字列として読み出す
fn read_zip_entry(archive: &mut ZipArchive<File>, name: &str) -> Result<String> {
    let mut entry = archive
        .by_name(name)
        .with_context(|| format!("ZIPエントリが見つかりません: {name}"))?;
    let mut buf = String::new();
    entry
        .read_to_string(&mut buf)
        .with_context(|| format!("ZIPエントリの読み込みに失敗: {name}"))?;
    Ok(buf)
}

/// セルの型と生の値から表示文字列を返す
fn resolve_cell_value(cell_type: &str, raw: &str, shared_strings: &[String]) -> String {
    match cell_type {
        "s" => raw
            .trim()
            .parse::<usize>()
            .ok()
            .and_then(|i| shared_strings.get(i))
            .cloned()
            .unwrap_or_default(),
        "b" => {
            if raw == "1" {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        "e" => raw.to_string(),
        _ => raw.trim().to_string(),
    }
}
