use std::fmt::Write as _;
use std::path::PathBuf;

use anyhow::{Context, Result};

use crate::models::{Document, Element, Section};

/// `to-asciidoc` サブコマンドの引数
#[derive(clap::Args)]
pub struct Args {
    /// 入力 document.json ファイルのパス
    #[arg(long)]
    pub input: PathBuf,

    /// 出力 .adoc ファイルのパス（省略時は入力 JSON と同じディレクトリに .adoc を生成）
    #[arg(long)]
    pub output: Option<PathBuf>,
}

pub fn run(args: Args) -> Result<()> {
    let json = std::fs::read_to_string(&args.input)
        .with_context(|| format!("ファイルを開けません: {}", args.input.display()))?;
    let doc: Document = serde_json::from_str(&json)
        .with_context(|| format!("JSON のパースに失敗: {}", args.input.display()))?;

    let adoc = convert(&doc);

    let out_path = match &args.output {
        Some(path) => path.clone(),
        None => args.input.with_extension("adoc"),
    };
    std::fs::write(&out_path, &adoc)
        .with_context(|| format!("書き込みに失敗: {}", out_path.display()))?;
    eprintln!("✓ {}", out_path.display());
    Ok(())
}

/// Document → AsciiDoc 文字列
fn convert(doc: &Document) -> String {
    let mut out = String::new();

    // ドキュメントタイトル
    writeln!(out, "= {}", doc.title).unwrap();
    writeln!(out, ":toc:").unwrap();
    writeln!(out, ":toc-title: 目次").unwrap();
    writeln!(out).unwrap();

    for section in &doc.sections {
        write_section(&mut out, section, 1);
    }
    out
}

/// セクションを再帰的に書き出す（depth = 1 が最上位）
fn write_section(out: &mut String, section: &Section, depth: usize) {
    // 見出し: depth 1 → "==", depth 2 → "===" ...
    let prefix = "=".repeat(depth + 1);
    writeln!(out, "{} {}", prefix, section.heading).unwrap();
    writeln!(out).unwrap();

    // assets マップ: asset_id → (title, data)
    let asset_map: std::collections::HashMap<&str, &crate::models::Asset> = section
        .assets
        .iter()
        .filter_map(|a| a.id.as_deref().map(|id| (id, a)))
        .collect();

    if !section.elements.is_empty() {
        for elem in &section.elements {
            write_element(out, elem, &asset_map);
        }
    } else if !section.body_text.is_empty() {
        // elements がない旧形式 JSON は body_text を Listing ブロックとして出力
        writeln!(out, "----").unwrap();
        writeln!(out, "{}", section.body_text).unwrap();
        writeln!(out, "----").unwrap();
        writeln!(out).unwrap();
    }

    for child in &section.children {
        write_section(out, child, depth + 1);
    }
}

fn write_element(
    out: &mut String,
    elem: &Element,
    assets: &std::collections::HashMap<&str, &crate::models::Asset>,
) {
    match elem {
        Element::Paragraph { text, metadata } => {
            if text.is_empty() {
                return;
            }
            // SemanticRole に応じたアドモニション
            let adoc = match &metadata.role {
                Some(crate::models::SemanticRole::Note) => {
                    format!("NOTE: {}\n", text)
                }
                Some(crate::models::SemanticRole::Warning) => {
                    format!("WARNING: {}\n", text)
                }
                Some(crate::models::SemanticRole::Tip) => {
                    format!("TIP: {}\n", text)
                }
                Some(crate::models::SemanticRole::CodeBlock) => {
                    format!("[source]\n----\n{}\n----\n", text)
                }
                Some(crate::models::SemanticRole::Quote) => {
                    format!("[quote]\n____\n{}\n____\n", text)
                }
                _ => format!("{}\n", text),
            };
            writeln!(out, "{}", adoc).unwrap();
        }

        Element::Table { rows, merges, .. } => {
            if rows.is_empty() {
                return;
            }
            let col_count = rows.iter().map(|r| r.len()).max().unwrap_or(0);
            if col_count == 0 {
                return;
            }

            // cols 指定（均等幅）
            let cols: Vec<&str> = vec!["1"; col_count];
            writeln!(out, r#"[cols="{}",options="header"]"#, cols.join(",")).unwrap();
            writeln!(out, "|===").unwrap();

            if merges.is_empty() {
                // 結合なし: コンパクトな1行形式
                for (i, row) in rows.iter().enumerate() {
                    if i == 1 {
                        writeln!(out).unwrap();
                    }
                    let cells: Vec<String> = (0..col_count)
                        .map(|c| escape_cell(row.get(c).map(|s| s.as_str()).unwrap_or("")))
                        .collect();
                    writeln!(out, "| {}", cells.join(" | ")).unwrap();
                }
            } else {
                // 結合あり: 1セル1行形式 + スパン記法
                use std::collections::{HashMap, HashSet};
                let span_map: HashMap<(usize, usize), (usize, usize)> = merges
                    .iter()
                    .map(|&(r, c, rs, cs)| ((r, c), (rs, cs)))
                    .collect();
                let mut covered: HashSet<(usize, usize)> = HashSet::new();
                for &(r, c, rs, cs) in merges {
                    for dr in 0..rs {
                        for dc in 0..cs {
                            if dr == 0 && dc == 0 {
                                continue;
                            }
                            covered.insert((r + dr, c + dc));
                        }
                    }
                }
                for (row_idx, row) in rows.iter().enumerate() {
                    if row_idx == 1 {
                        writeln!(out).unwrap();
                    }
                    for col_idx in 0..col_count {
                        if covered.contains(&(row_idx, col_idx)) {
                            continue;
                        }
                        let text =
                            escape_cell(row.get(col_idx).map(|s| s.as_str()).unwrap_or(""));
                        let prefix = span_map
                            .get(&(row_idx, col_idx))
                            .map(|&(rs, cs)| cell_span_prefix(cs, rs))
                            .unwrap_or_default();
                        writeln!(out, "{}| {}", prefix, text).unwrap();
                    }
                }
            }
            writeln!(out, "|===").unwrap();
            writeln!(out).unwrap();
        }

        Element::AssetRef { asset_id, metadata } => {
            if let Some(asset) = assets.get(asset_id.as_str()) {
                let alt = if !asset.title.is_empty() {
                    asset.title.clone()
                } else {
                    metadata
                        .caption
                        .clone()
                        .unwrap_or_else(|| asset_id.clone())
                };
                if !asset.data.is_empty() {
                    use base64::Engine;
                    let b64 = base64::engine::general_purpose::STANDARD.encode(&asset.data);
                    writeln!(out, "image::data:image/jpeg;base64,{}[{}]", b64, alt).unwrap();
                } else {
                    writeln!(out, "image::{}[{}]", asset_id, alt).unwrap();
                }
                writeln!(out).unwrap();
            }
        }
    }
}

/// AsciiDoc テーブルセルのスパン記法プレフィックスを生成する
///
/// - colspan 2, rowspan 1 → `"2+"`
/// - colspan 1, rowspan 3 → `".3+"`
/// - colspan 2, rowspan 3 → `"2.3+"`
/// - colspan 1, rowspan 1 → `""` (スパンなし)
fn cell_span_prefix(colspan: usize, rowspan: usize) -> String {
    match (colspan > 1, rowspan > 1) {
        (true, true) => format!("{}.{}+", colspan, rowspan),
        (true, false) => format!("{}+", colspan),
        (false, true) => format!(".{}+", rowspan),
        (false, false) => String::new(),
    }
}

/// AsciiDoc テーブルセル内の `|` をエスケープし、改行をスペースに変換
fn escape_cell(s: &str) -> String {
    s.replace('|', "\\|")
        .replace('\n', " ")
        .replace('\r', "")
}
