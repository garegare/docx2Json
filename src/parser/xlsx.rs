use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::Path;

use anyhow::{Context, Result};
use quick_xml::events::Event;
use quick_xml::Reader;
use zip::ZipArchive;

use crate::config::Config;
use crate::models::{Document, Element, ElementMetadata, Section};

/// セル結合範囲: (min_row, min_col, max_row, max_col)（0-based、両端含む）
type MergeRange = (usize, usize, usize, usize);

/// XLSXファイルを解析してDocumentを返す
///
/// 各シートを1つのSectionに変換する。シートのデータ行数が `config.xlsx.max_rows` を
/// 超える場合、ヘッダー行を保持したまま子Sectionに分割する。
pub fn parse(path: &Path, config: &Config) -> Result<Document> {
    let title = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("Untitled")
        .to_string();

    let file = File::open(path)
        .with_context(|| format!("ファイルを開けません: {}", path.display()))?;
    let mut archive = ZipArchive::new(file).context("ZIPアーカイブとして開けません")?;

    // 1. xl/_rels/workbook.xml.rels → rId → ファイルパス
    let rels = parse_workbook_rels(&mut archive).unwrap_or_default();

    // 2. xl/workbook.xml → シート名と rId のリスト（順序保持）
    let sheets = parse_workbook(&mut archive)?;

    // 3. xl/sharedStrings.xml → 共有文字列テーブル（存在しない場合は空）
    let shared_strings = parse_shared_strings(&mut archive).unwrap_or_default();

    // 4. 各シートを Section に変換
    let mut sections = Vec::new();
    for (name, rid) in &sheets {
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

        // "worksheets/sheet1.xml" → "xl/worksheets/sheet1.xml"
        let sheet_path = resolve_path(&target);

        match parse_worksheet(&mut archive, &sheet_path, &shared_strings) {
            Ok((grid, merges)) => {
                let section = sheet_to_section(name, grid, merges, config.xlsx.max_rows);
                sections.push(section);
            }
            Err(e) => {
                eprintln!("Warning: シート '{}' の解析に失敗しました: {e}", name);
            }
        }
    }

    Ok(Document { title, sections })
}

/// シート名とグリッドデータから Section を生成する
///
/// - `max_rows == 0` またはデータ行数 ≤ max_rows の場合:
///   フラットな Markdown 表を body_text に格納する。
/// - データ行数 > max_rows の場合:
///   親 Section の body_text に概要（行数・チャンク数）を格納し、
///   ヘッダー行を引き継いだ子 Section に max_rows 行ずつ分割する。
fn sheet_to_section(name: &str, rows: Vec<Vec<String>>, merges: Vec<(usize, usize, usize, usize)>, max_rows: usize) -> Section {
    if rows.is_empty() {
        return Section {
            context_path: Vec::new(),
            heading: name.to_string(),
            body_text: String::new(),
            assets: Vec::new(),
            children: Vec::new(),
            ..Default::default()
        };
    }

    let data_row_count = rows.len().saturating_sub(1); // ヘッダー行を除くデータ行数

    if max_rows == 0 || data_row_count <= max_rows {
        // 制限なし、または行数が範囲内: フラットに出力
        return Section {
            context_path: Vec::new(),
            heading: name.to_string(),
            body_text: grid_to_markdown(&rows),
            elements: grid_to_elements(&rows, &merges),
            assets: Vec::new(),
            children: Vec::new(),
            ..Default::default()
        };
    }

    // データ行数超過: ヘッダー行を保持しながら子 Section に分割
    let header = rows[0].clone();
    let data_rows = &rows[1..];
    let chunk_count = data_row_count.div_ceil(max_rows);

    let children: Vec<Section> = data_rows
        .chunks(max_rows)
        .enumerate()
        .map(|(i, chunk)| {
            let start = i * max_rows + 1; // 1-indexed データ行番号
            let end = start + chunk.len() - 1;

            // 子 Section の body_text: ヘッダー行 + このチャンクのデータ行
            let mut child_rows = Vec::with_capacity(chunk.len() + 1);
            child_rows.push(header.clone());
            child_rows.extend_from_slice(chunk);

            // チャンクに対応するマージ情報を child_rows のローカル座標に変換する。
            // - ヘッダー行 (orig_row == 0) のマージはそのまま引き継ぐ。
            // - データ行のマージはチャンク内 (orig_row >= chunk_orig_start) のものだけ抽出し、
            //   child_rows 上の行番号（1-based）に変換する。
            // チャンクをまたぐ rowspan は chunk 末尾でクランプする。
            let chunk_orig_start = 1 + i * max_rows; // シート上のデータ開始行（0-based）
            let chunk_orig_end = chunk_orig_start + chunk.len() - 1; // 同・終端行（含む）
            let child_merges: Vec<(usize, usize, usize, usize)> = merges
                .iter()
                .filter_map(|&(min_r, min_c, max_r, max_c)| {
                    if min_r == 0 && max_r == 0 {
                        // ヘッダー行のみのマージ: child_rows[0] にそのまま適用
                        return Some((0, min_c, 0, max_c));
                    }
                    if min_r >= chunk_orig_start && min_r <= chunk_orig_end {
                        // データ行のマージ: child_rows 上の行番号に変換
                        // child_rows[0] はヘッダーなので +1 オフセット
                        let new_min_r = min_r - chunk_orig_start + 1;
                        let new_max_r = max_r.min(chunk_orig_end) - chunk_orig_start + 1;
                        return Some((new_min_r, min_c, new_max_r, max_c));
                    }
                    None
                })
                .collect();

            Section {
                context_path: Vec::new(), // fill_context_path() で後から設定
                heading: format!("{} [行 {}–{}]", name, start, end),
                body_text: grid_to_markdown(&child_rows),
                elements: grid_to_elements(&child_rows, &child_merges),
                assets: Vec::new(),
                children: Vec::new(),
                ..Default::default()
            }
        })
        .collect();

    let summary = format!(
        "（全 {} 行 / {} 行ずつ {} チャンクに分割）",
        data_row_count, max_rows, chunk_count
    );
    Section {
        context_path: Vec::new(),
        heading: name.to_string(),
        body_text: summary.clone(),
        elements: vec![crate::models::Element::Paragraph {
            text: summary,
            metadata: crate::models::ElementMetadata::default(),
        }],
        assets: Vec::new(),
        children,
        ..Default::default()
    }
}

/// グリッドを構造化 Element のリストに変換する
///
/// 空行を区切りとして「ブロック」を検出し、ブロックごとに Element を生成する:
/// - 1行かつ非空セルが1つのみ → `Paragraph`（タイトル行・ラベル等）
/// - それ以外 → `Table`（複数行 or 複数セル）
fn grid_to_elements(rows: &[Vec<String>], sheet_merges: &[(usize, usize, usize, usize)]) -> Vec<Element> {
    split_into_blocks(rows)
        .into_iter()
        .map(|(block, block_start)| block_to_element(block, block_start, sheet_merges))
        .collect()
}

/// グリッドを空行区切りでブロックに分割する
///
/// 全セルが空文字の行を区切りとして扱い、連続する非空行をひとまとめにする。
/// 戻り値: (ブロック, シート内の開始行インデックス)
fn split_into_blocks(rows: &[Vec<String>]) -> Vec<(Vec<Vec<String>>, usize)> {
    let mut blocks = Vec::new();
    let mut current: Vec<Vec<String>> = Vec::new();
    let mut block_start = 0usize;

    for (row_idx, row) in rows.iter().enumerate() {
        if row.iter().all(|s| s.is_empty()) {
            if !current.is_empty() {
                blocks.push((std::mem::take(&mut current), block_start));
            }
            block_start = row_idx + 1;
        } else {
            if current.is_empty() {
                block_start = row_idx;
            }
            current.push(row.clone());
        }
    }
    if !current.is_empty() {
        blocks.push((current, block_start));
    }
    blocks
}

/// ブロック（非空行の連続）から Element を生成する
///
/// 1行で以下のいずれかに該当する場合は `Paragraph` として扱う:
/// - 非空セルが1つのみ（通常の単一セルタイトル）
/// - 非空セルが複数あるが全て同一値（横結合タイトルの展開結果）
///
/// それ以外は `Table` として扱う。
/// `block_start` はシートグリッド内でのこのブロックの開始行インデックス（0-based）。
/// `sheet_merges` は (min_row, min_col, max_row, max_col) 形式のシート全体のマージ情報。
fn block_to_element(
    block: Vec<Vec<String>>,
    block_start: usize,
    sheet_merges: &[(usize, usize, usize, usize)],
) -> Element {
    if block.len() == 1 {
        let non_empty: Vec<&String> = block[0].iter().filter(|s| !s.is_empty()).collect();
        let is_title = match non_empty.len() {
            0 => false,
            1 => true,
            // 全セルが同一値 → 横結合タイトルが展開されたもの
            _ => non_empty.windows(2).all(|w| w[0] == w[1]),
        };
        if is_title {
            return Element::Paragraph {
                text: non_empty[0].clone(),
                metadata: ElementMetadata::default(),
            };
        }
    }

    let block_rows = block.len();
    let block_end = block_start + block_rows - 1; // ブロック最終行（シート絶対座標、含む）
    // シート絶対座標のマージをブロックローカル (row, col, rowspan, colspan) に変換
    let merges: Vec<(usize, usize, usize, usize)> = sheet_merges
        .iter()
        .filter_map(|&(min_r, min_c, max_r, max_c)| {
            // このブロック内に開始セルがあるものだけ対象
            if min_r < block_start || min_r >= block_start + block_rows {
                return None;
            }
            // ブロック外にはみ出す rowspan はブロック末尾でクランプする
            let clamped_max_r = max_r.min(block_end);
            let rowspan = clamped_max_r - min_r + 1;
            let colspan = max_c - min_c + 1;
            if rowspan == 1 && colspan == 1 {
                return None; // 1×1 は結合なし
            }
            Some((min_r - block_start, min_c, rowspan, colspan))
        })
        .collect();

    Element::Table {
        rows: block,
        merges,
        metadata: ElementMetadata::default(),
    }
}

// ---- XML パーサー群 ----

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
                        // r:id の local_name は "id"
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
///
/// リッチテキスト（`<si><r><t>...</t></r></si>`）にも対応し、
/// `<t>` 要素のテキストを順に結合して1エントリとする。
fn parse_shared_strings(archive: &mut ZipArchive<File>) -> Result<Vec<String>> {
    let content = read_zip_entry(archive, "xl/sharedStrings.xml")?;
    let mut strings = Vec::new();

    let mut reader = Reader::from_str(&content);
    // 空白保持: セル値の先頭・末尾スペースを守るため trim_text を false に
    reader.config_mut().trim_text(false);

    let mut in_si = false;
    let mut in_rph = false; // <rPh>（ルビ・読み仮名）内は読み飛ばす
    let mut current = String::new();

    loop {
        match reader.read_event()? {
            Event::Start(e) if e.local_name().as_ref() == b"si" => {
                in_si = true;
                in_rph = false;
                current.clear();
            }
            Event::End(e) if e.local_name().as_ref() == b"si" => {
                strings.push(current.trim().to_string());
                in_si = false;
            }
            Event::Start(e) if in_si && e.local_name().as_ref() == b"rPh" => {
                in_rph = true;
            }
            Event::End(e) if in_si && e.local_name().as_ref() == b"rPh" => {
                in_rph = false;
            }
            Event::Text(e) if in_si && !in_rph => {
                current.push_str(&e.unescape().unwrap_or_default());
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(strings)
}

/// `xl/worksheets/sheet*.xml` を解析してセルグリッドを返す
///
/// 戻り値: 行×列の密な 2D Vec（空セルは空文字列）
///
/// `<mergeCells>` を解析し、結合元セルの値を結合範囲全体に展開する。
/// これにより、ドキュメント風 Excel（神エクセル）で多用されるタイトルや
/// ラベル用の結合セルが正しく取り込まれる。
fn parse_worksheet(
    archive: &mut ZipArchive<File>,
    path: &str,
    shared_strings: &[String],
) -> Result<(Vec<Vec<String>>, Vec<MergeRange>)> {
    let content = read_zip_entry(archive, path)?;
    let mut reader = Reader::from_str(&content);
    reader.config_mut().trim_text(true);

    // スパースグリッド: (row_idx, col_idx) → セル値
    let mut sparse: HashMap<(usize, usize), String> = HashMap::new();
    let mut max_row = 0usize;
    let mut max_col = 0usize;

    // 結合セル範囲リスト: (min_row, min_col, max_row, max_col)（0-indexed、両端含む）
    let mut merges: Vec<(usize, usize, usize, usize)> = Vec::new();

    // 現在処理中のセル情報
    let mut cur_row = 0usize;
    let mut cur_col = 0usize;
    let mut cell_type = String::new();
    let mut in_v = false; // <v> 要素内（数値・共有文字列インデックス・式結果）
    let mut in_t = false; // <t> 要素内（inlineStr）
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
                        cell_buf.clear();
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"r" => {
                                    let cell_ref = decode_bytes(&attr.value);
                                    cur_col = col_index(&cell_ref);
                                    max_col = max_col.max(cur_col + 1);
                                }
                                b"t" => cell_type = decode_bytes(&attr.value),
                                _ => {}
                            }
                        }
                    }
                    b"v" => in_v = true,
                    b"t" => in_t = true,
                    // 結合セル範囲を収集: "A1:C3" 形式の ref 属性を解析
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
                    // セル値を確定してスパースグリッドに格納
                    let val = resolve_cell_value(&cell_type, &cell_buf, shared_strings);
                    if !val.is_empty() {
                        sparse.insert((cur_row, cur_col), val);
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

    // スパースグリッド → 密な 2D Vec
    if max_row == 0 || max_col == 0 {
        return Ok((Vec::new(), Vec::new()));
    }
    let mut grid = vec![vec![String::new(); max_col]; max_row];
    for ((r, c), val) in sparse {
        grid[r][c] = val;
    }

    expand_merges(&mut grid, &merges);
    Ok((grid, merges))
}

/// セルの型と生の値から表示文字列を返す
///
/// | `t` 属性 | 意味 | 処理 |
/// |----------|------|------|
/// | `"s"`    | 共有文字列インデックス | shared_strings テーブルを引く |
/// | `"b"`    | 真偽値 | "1" → "TRUE", それ以外 → "FALSE" |
/// | `"e"`    | エラー (#DIV/0! 等) | 生値そのまま |
/// | 省略     | 数値・文字列式結果 | trim して返す |
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

// ---- Markdown 生成 ----

/// セルグリッドを Markdown 表に変換する
///
/// - 先頭行をヘッダーとして扱い、2行目にセパレータ（`|---|`）を挿入する
/// - 列数はグリッド内の最大列数に揃える
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
            // ヘッダー直後にセパレータ行を挿入
            let sep = vec!["---"; col_count];
            out.push_str(&format!("|{}|\n", sep.join("|")));
        }
    }
    out
}

/// Markdown テーブルセル内の特殊文字をエスケープする
fn escape_cell(s: &str) -> String {
    s.replace('|', "\\|")
        .replace('\n', " ")
        .replace('\r', "")
}

// ---- ユーティリティ ----

/// 結合セル範囲リストに従い、グリッドの結合元（左上）の値を範囲全体にコピーする
///
/// - 結合元が空文字の場合は展開しない
/// - グリッド範囲外の merge 指定は無視する（破損ファイルへの防御）
fn expand_merges(grid: &mut [Vec<String>], merges: &[(usize, usize, usize, usize)]) {
    for (min_row, min_col, max_row, max_col) in merges {
        let origin = grid
            .get(*min_row)
            .and_then(|r| r.get(*min_col))
            .cloned()
            .unwrap_or_default();
        if origin.is_empty() {
            continue;
        }
        for r in *min_row..=*max_row {
            for c in *min_col..=*max_col {
                if (r, c) == (*min_row, *min_col) {
                    continue; // 結合元はスキップ
                }
                if let Some(cell) = grid.get_mut(r).and_then(|row| row.get_mut(c)) {
                    cell.clone_from(&origin);
                }
            }
        }
    }
}

/// セル参照（"A1"、"AB12" 等）からゼロ始まりの列インデックスを返す
///
/// A=0, B=1, …, Z=25, AA=26, AB=27, …
fn col_index(cell_ref: &str) -> usize {
    // parse_cell_address でアドレス全体を解析し、列インデックスのみを返す
    parse_cell_address(cell_ref)
        .map(|(_row, col)| col)
        .unwrap_or(0)
}

/// "A1:C3" 形式のセル範囲文字列を `(min_row, min_col, max_row, max_col)`（0-indexed）に変換する
fn parse_merge_range(ref_str: &str) -> Option<(usize, usize, usize, usize)> {
    let mut parts = ref_str.splitn(2, ':');
    let (sr, sc) = parse_cell_address(parts.next()?)?;
    let (er, ec) = parse_cell_address(parts.next()?)?;
    Some((sr, sc, er, ec))
}

/// "A1"、"AB12" 等のセルアドレス文字列を `(row, col)`（0-indexed）に変換する
fn parse_cell_address(addr: &str) -> Option<(usize, usize)> {
    let col_str: String = addr.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let row_str: String = addr.chars().skip_while(|c| c.is_ascii_alphabetic()).collect();
    let col = col_str
        .chars()
        .fold(0usize, |acc, c| {
            acc * 26 + (c.to_ascii_uppercase() as usize - b'A' as usize + 1)
        })
        .checked_sub(1)?;
    let row: usize = row_str.parse::<usize>().ok()?.checked_sub(1)?;
    Some((row, col))
}

/// `xl/_rels/workbook.xml.rels` の Target パスを ZIP 内絶対パスに解決する
///
/// - 絶対パス（`/xl/worksheets/sheet1.xml`）→ 先頭スラッシュを除去
/// - 相対パス（`worksheets/sheet1.xml`）→ `xl/` を前置
fn resolve_path(target: &str) -> String {
    if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        format!("xl/{}", target)
    }
}

/// `Cow<[u8]>` を UTF-8 文字列にデコードする
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_index() {
        assert_eq!(col_index("A1"), 0);
        assert_eq!(col_index("Z1"), 25);
        assert_eq!(col_index("AA1"), 26);
        assert_eq!(col_index("AB1"), 27);
    }

    #[test]
    fn test_parse_cell_address() {
        assert_eq!(parse_cell_address("A1"), Some((0, 0)));
        assert_eq!(parse_cell_address("C3"), Some((2, 2)));
        assert_eq!(parse_cell_address("AA10"), Some((9, 26)));
        assert_eq!(parse_cell_address(""), None);
    }

    #[test]
    fn test_parse_merge_range() {
        assert_eq!(parse_merge_range("A1:C3"), Some((0, 0, 2, 2)));
        assert_eq!(parse_merge_range("B2:D4"), Some((1, 1, 3, 3)));
        // 単一セル（A1:A1）も受け付ける
        assert_eq!(parse_merge_range("A1:A1"), Some((0, 0, 0, 0)));
        // コロンなし → None
        assert_eq!(parse_merge_range("A1"), None);
    }

    #[test]
    fn test_merge_expansion_horizontal() {
        // A1:C1 の横結合: "タイトル" が B1, C1 にコピーされる
        let mut grid = vec![
            vec!["タイトル".to_string(), String::new(), String::new()],
            vec!["A".to_string(), "B".to_string(), "C".to_string()],
        ];
        expand_merges(&mut grid, &[(0, 0, 0, 2)]);
        assert_eq!(grid[0], vec!["タイトル", "タイトル", "タイトル"]);
        assert_eq!(grid[1], vec!["A", "B", "C"]); // 他行は変化なし
    }

    #[test]
    fn test_merge_expansion_vertical() {
        // A1:A3 の縦結合: "ラベル" が A2, A3 にコピーされる
        let mut grid = vec![
            vec!["ラベル".to_string(), "X".to_string()],
            vec![String::new(), "Y".to_string()],
            vec![String::new(), "Z".to_string()],
        ];
        expand_merges(&mut grid, &[(0, 0, 2, 0)]);
        assert_eq!(grid[0][0], "ラベル");
        assert_eq!(grid[1][0], "ラベル");
        assert_eq!(grid[2][0], "ラベル");
    }

    #[test]
    fn test_merge_expansion_empty_origin_skipped() {
        // 結合元が空の場合は展開しない
        let mut grid = vec![vec![String::new(), String::new(), String::new()]];
        expand_merges(&mut grid, &[(0, 0, 0, 2)]);
        assert_eq!(grid[0], vec!["", "", ""]);
    }

    // ---- grid_to_elements / split_into_blocks / block_to_element のテスト ----

    #[test]
    fn test_split_into_blocks_single_block() {
        let rows = vec![
            vec!["A".to_string(), "B".to_string()],
            vec!["1".to_string(), "2".to_string()],
        ];
        let blocks = split_into_blocks(&rows);
        assert_eq!(blocks.len(), 1);
        assert_eq!(blocks[0].0.len(), 2);
        assert_eq!(blocks[0].1, 0); // start_row
    }

    #[test]
    fn test_split_into_blocks_two_blocks() {
        let rows = vec![
            vec!["タイトル".to_string(), "".to_string()],
            vec!["".to_string(), "".to_string()], // 空行
            vec!["A".to_string(), "B".to_string()],
            vec!["1".to_string(), "2".to_string()],
        ];
        let blocks = split_into_blocks(&rows);
        assert_eq!(blocks.len(), 2);
        assert_eq!(blocks[0].0.len(), 1);
        assert_eq!(blocks[0].1, 0); // start_row
        assert_eq!(blocks[1].0.len(), 2);
        assert_eq!(blocks[1].1, 2); // start_row（空行の次）
    }

    #[test]
    fn test_split_into_blocks_empty() {
        let rows: Vec<Vec<String>> = vec![];
        assert_eq!(split_into_blocks(&rows).len(), 0);
    }

    #[test]
    fn test_block_to_element_paragraph() {
        // 1行・1セル → Paragraph
        let block = vec![vec!["申請書".to_string(), "".to_string(), "".to_string()]];
        match block_to_element(block, 0, &[]) {
            Element::Paragraph { text, .. } => assert_eq!(text, "申請書"),
            other => panic!("Expected Paragraph, got {:?}", other),
        }
    }

    #[test]
    fn test_block_to_element_table_multi_row() {
        // 複数行 → Table
        let block = vec![
            vec!["A".to_string(), "B".to_string()],
            vec!["1".to_string(), "2".to_string()],
        ];
        match block_to_element(block, 0, &[]) {
            Element::Table { rows, .. } => assert_eq!(rows.len(), 2),
            other => panic!("Expected Table, got {:?}", other),
        }
    }

    #[test]
    fn test_block_to_element_table_multi_cell() {
        // 1行・複数セルで値が異なる → Table
        let block = vec![vec!["部署".to_string(), "総務部".to_string()]];
        match block_to_element(block, 0, &[]) {
            Element::Table { rows, .. } => assert_eq!(rows.len(), 1),
            other => panic!("Expected Table, got {:?}", other),
        }
    }

    #[test]
    fn test_block_to_element_paragraph_from_merged_title() {
        // 横結合タイトルの展開結果（全セルが同一値）→ Paragraph
        let block = vec![vec![
            "API仕様書".to_string(),
            "API仕様書".to_string(),
            "API仕様書".to_string(),
        ]];
        match block_to_element(block, 0, &[]) {
            Element::Paragraph { text, .. } => assert_eq!(text, "API仕様書"),
            other => panic!("Expected Paragraph, got {:?}", other),
        }
    }

    #[test]
    fn test_block_to_element_table_when_values_differ() {
        // 1行・複数セルで値が異なる → Table（誤検知しない）
        let block = vec![vec![
            "パラメータ名".to_string(),
            "型".to_string(),
            "必須".to_string(),
        ]];
        match block_to_element(block, 0, &[]) {
            Element::Table { .. } => {}
            other => panic!("Expected Table, got {:?}", other),
        }
    }

    #[test]
    fn test_block_to_element_merges_block_local() {
        // シート行2-4のブロック（block_start=2）に (2,0,4,1) のマージ → ローカル (0,0,3,2)
        let block = vec![
            vec!["結合".to_string(), "結合".to_string(), "A".to_string()],
            vec!["結合".to_string(), "結合".to_string(), "B".to_string()],
            vec!["結合".to_string(), "結合".to_string(), "C".to_string()],
        ];
        let sheet_merges = vec![(2usize, 0usize, 4usize, 1usize)]; // min_row=2, min_col=0, max_row=4, max_col=1
        match block_to_element(block, 2, &sheet_merges) {
            Element::Table { merges, .. } => {
                assert_eq!(merges.len(), 1);
                assert_eq!(merges[0], (0, 0, 3, 2)); // row=0, col=0, rowspan=3, colspan=2
            }
            other => panic!("Expected Table, got {:?}", other),
        }
    }

    #[test]
    fn test_grid_to_elements_mixed() {
        // タイトル行 + 空行 + データテーブル → [Paragraph, Table]
        let rows = vec![
            vec!["申請書".to_string(), "".to_string()],
            vec!["".to_string(), "".to_string()],
            vec!["部署".to_string(), "総務部".to_string()],
            vec!["担当".to_string(), "田中".to_string()],
        ];
        let elements = grid_to_elements(&rows, &[]);
        assert_eq!(elements.len(), 2);
        assert!(matches!(elements[0], Element::Paragraph { .. }));
        assert!(matches!(elements[1], Element::Table { .. }));
    }

    #[test]
    fn test_merge_expansion_out_of_bounds_safe() {
        // 壊れた merge 指定でもパニックしない
        let mut grid = vec![vec!["値".to_string()]];
        expand_merges(&mut grid, &[(0, 0, 5, 5)]); // グリッド範囲外
        // グリッド内の (0,0) は変化なし、範囲外への書き込みは無視される
        assert_eq!(grid[0][0], "値");
    }
}
