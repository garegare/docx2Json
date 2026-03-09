use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::Path;

use anyhow::{Context, Result};
use quick_xml::events::Event;
use quick_xml::Reader;
use zip::ZipArchive;

use crate::config::Config;
use crate::models::{Document, Section};

/// XLSXファイルを解析してDocumentを返す
///
/// 各シートを1つのSectionに変換する。シートのデータ行数が `config.xlsx_max_rows` を
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
            Ok(grid) => {
                let section = sheet_to_section(name, grid, config.xlsx_max_rows);
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
fn sheet_to_section(name: &str, rows: Vec<Vec<String>>, max_rows: usize) -> Section {
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

            Section {
                context_path: Vec::new(), // fill_context_path() で後から設定
                heading: format!("{} [行 {}–{}]", name, start, end),
                body_text: grid_to_markdown(&child_rows),
                assets: Vec::new(),
                children: Vec::new(),
                ..Default::default()
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
        ..Default::default()
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

/// `xl/worksheets/sheet*.xml` を解析してセルグリッドを返す
///
/// 戻り値: 行×列の密な 2D Vec（空セルは空文字列）
fn parse_worksheet(
    archive: &mut ZipArchive<File>,
    path: &str,
    shared_strings: &[String],
) -> Result<Vec<Vec<String>>> {
    let content = read_zip_entry(archive, path)?;
    let mut reader = Reader::from_str(&content);
    reader.config_mut().trim_text(true);

    // スパースグリッド: (row_idx, col_idx) → セル値
    let mut sparse: HashMap<(usize, usize), String> = HashMap::new();
    let mut max_row = 0usize;
    let mut max_col = 0usize;

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
        return Ok(Vec::new());
    }
    let mut grid = vec![vec![String::new(); max_col]; max_row];
    for ((r, c), val) in sparse {
        grid[r][c] = val;
    }
    Ok(grid)
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

/// セル参照（"A1"、"AB12" 等）からゼロ始まりの列インデックスを返す
///
/// A=0, B=1, …, Z=25, AA=26, AB=27, …
fn col_index(cell_ref: &str) -> usize {
    cell_ref
        .chars()
        .take_while(|c| c.is_ascii_alphabetic())
        .fold(0usize, |acc, c| {
            acc * 26 + (c.to_ascii_uppercase() as usize - b'A' as usize + 1)
        })
        .saturating_sub(1)
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
