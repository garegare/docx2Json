use std::collections::{BTreeSet, HashMap, HashSet};
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
    let asset_map: HashMap<&str, &crate::models::Asset> = section
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
    assets: &HashMap<&str, &crate::models::Asset>,
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

            // Step 1: 神エクセル等の多列テーブルを論理列に圧縮
            let (rows, merges) = compress_columns(rows, merges);

            // Step 2: 隣接する排他的列を階層インデントで統合
            let (rows, merges) = merge_exclusive_cols(&rows, &merges, "　");

            let col_count = rows.iter().map(|r| r.len()).max().unwrap_or(0);
            if col_count == 0 {
                return;
            }

            // Step 3: 全幅ヘッダー行で分割して出力
            write_table_with_headers(out, &rows, &merges, col_count);
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

// ─────────────────────────────────────────────────────────────────────────────
// テーブル変換ヘルパー
// ─────────────────────────────────────────────────────────────────────────────

/// `merges` から span_map と covered を構築する共通ヘルパー
fn build_span_and_covered(
    merges: &[(usize, usize, usize, usize)],
) -> (HashMap<(usize, usize), (usize, usize)>, HashSet<(usize, usize)>) {
    let span_map: HashMap<(usize, usize), (usize, usize)> =
        merges.iter().map(|&(r, c, rs, cs)| ((r, c), (rs, cs))).collect();
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
    (span_map, covered)
}

/// 全幅ヘッダー行でテーブルを分割し、ヘッダー行は太字段落として出力する。
///
/// 全幅ヘッダー行 = 可視セルが1つだけ、かつそのセルの colspan が全列数 - 1 以上の行。
fn write_table_with_headers(
    out: &mut String,
    rows: &[Vec<String>],
    merges: &[(usize, usize, usize, usize)],
    col_count: usize,
) {
    let (span_map, covered) = build_span_and_covered(merges);

    // 各行が「全幅ヘッダー行」かどうかを判定
    // 全幅ヘッダー = 非空・非カバーセルが1つだけで、そのspanが列全体 - 1 以上
    let is_header_row = |ri: usize| -> Option<String> {
        let visible_nonempty: Vec<usize> = (0..col_count)
            .filter(|&c| !covered.contains(&(ri, c)))
            .filter(|&c| {
                rows.get(ri)
                    .and_then(|r| r.get(c))
                    .map(|s| !s.is_empty())
                    .unwrap_or(false)
            })
            .collect();
        if visible_nonempty.len() != 1 {
            return None;
        }
        let c = visible_nonempty[0];
        let cs = span_map.get(&(ri, c)).map(|&(_, cs)| cs).unwrap_or(1);
        if cs >= col_count.saturating_sub(1) {
            rows.get(ri)
                .and_then(|r| r.get(c))
                .filter(|v| !v.is_empty())
                .map(|v| v.replace('\r', "").replace('\n', " "))
        } else {
            None
        }
    };

    // 行を「ヘッダー」か「データ」かで分類してセグメント化
    let mut segments: Vec<(Option<String>, Vec<usize>)> = vec![];
    let mut data_rows: Vec<usize> = vec![];

    for ri in 0..rows.len() {
        if let Some(header_text) = is_header_row(ri) {
            if !data_rows.is_empty() {
                segments.push((None, std::mem::take(&mut data_rows)));
            }
            segments.push((Some(header_text), vec![]));
        } else {
            data_rows.push(ri);
        }
    }
    if !data_rows.is_empty() {
        segments.push((None, data_rows));
    }

    // セグメントを出力
    for (header, row_indices) in &segments {
        if let Some(text) = header {
            writeln!(out, "*{}*\n", escape_cell(text)).unwrap();
        }
        if !row_indices.is_empty() {
            let cols: Vec<&str> = vec!["1"; col_count];
            writeln!(out, r#"[cols="{}",options="header"]"#, cols.join(",")).unwrap();
            writeln!(out, "|===").unwrap();

            for (seg_idx, &ri) in row_indices.iter().enumerate() {
                if seg_idx == 1 {
                    writeln!(out).unwrap(); // ヘッダー行とデータ行の間の空行
                }
                for ci in 0..col_count {
                    if covered.contains(&(ri, ci)) {
                        continue;
                    }
                    let text =
                        escape_cell(rows.get(ri).and_then(|r| r.get(ci)).map(|s| s.as_str()).unwrap_or(""));
                    let prefix = span_map
                        .get(&(ri, ci))
                        .map(|&(rs, cs)| cell_span_prefix(cs, rs))
                        .unwrap_or_default();
                    writeln!(out, "{}| {}", prefix, text).unwrap();
                }
            }
            writeln!(out, "|===").unwrap();
            writeln!(out).unwrap();
        }
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// 列圧縮
// ─────────────────────────────────────────────────────────────────────────────

/// 神エクセル等の多列テーブルを「論理列」に圧縮する。
///
/// 1. マージあり: 全マージの開始列を論理列境界として収集し 125列→11列のように縮小する。
/// 2. マージなし: 全行で空の列を除去する。
///
/// `rows` と `merges` を論理座標に変換して返す。
fn compress_columns(
    rows: &[Vec<String>],
    merges: &[(usize, usize, usize, usize)], // (row, col, rowspan, colspan)
) -> (Vec<Vec<String>>, Vec<(usize, usize, usize, usize)>) {
    if merges.is_empty() {
        // マージなし: 全行で空の列を除去
        let max_cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
        let keep: Vec<usize> = (0..max_cols)
            .filter(|&c| rows.iter().any(|r| !r.get(c).map(|s| s.is_empty()).unwrap_or(true)))
            .collect();
        if keep.len() == max_cols {
            return (rows.to_vec(), vec![]);
        }
        let new_rows = rows
            .iter()
            .map(|r| keep.iter().map(|&c| r.get(c).cloned().unwrap_or_default()).collect())
            .collect();
        return (new_rows, vec![]);
    }

    // 1. 全マージ開始列 + 0 を論理列境界として収集
    let mut boundaries: BTreeSet<usize> = BTreeSet::new();
    boundaries.insert(0);
    for &(_, c, _, cs) in merges {
        boundaries.insert(c);
        boundaries.insert(c + cs); // 終端境界（次論理列の先頭）
    }

    // 1b. マージ開始でないスタンドアロンセルの列も境界に追加する。
    //     例: 神エクセルの階層表で "metadata"(col9), "title"(col10) などが
    //     マージなしの単独セルとして複数行に出現する場合、それらを別論理列に保つ。
    {
        // covered はここではまだ合算されていないので仮計算
        let mut tmp_covered: HashSet<(usize, usize)> = HashSet::new();
        let merge_starts: HashSet<(usize, usize)> =
            merges.iter().map(|&(r, c, _, _)| (r, c)).collect();
        for &(r, c, rs, cs) in merges {
            for dr in 0..rs {
                for dc in 0..cs {
                    if dr == 0 && dc == 0 {
                        continue;
                    }
                    tmp_covered.insert((r + dr, c + dc));
                }
            }
        }
        let mut standalone_freq: HashMap<usize, usize> = HashMap::new();
        for (ri, row) in rows.iter().enumerate() {
            for (ci, val) in row.iter().enumerate() {
                if !val.is_empty()
                    && !tmp_covered.contains(&(ri, ci))
                    && !merge_starts.contains(&(ri, ci))
                {
                    *standalone_freq.entry(ci).or_insert(0) += 1;
                }
            }
        }
        // 2行以上で出現するスタンドアロン列を境界として追加
        for (c, &freq) in &standalone_freq {
            if freq >= 2 {
                boundaries.insert(*c);
            }
        }
    }

    let logical_cols: Vec<usize> = boundaries.into_iter().collect(); // ソート済み
    let n_logical = logical_cols.len();

    // 2. 物理列 → 論理列インデックスのマップを構築
    //    論理列 i は [logical_cols[i], logical_cols[i+1]) の物理列を担当
    let phys_to_logical: HashMap<usize, usize> = {
        let max_phys = rows.iter().map(|r| r.len()).max().unwrap_or(0);
        let mut m = HashMap::new();
        for (li, &start) in logical_cols.iter().enumerate() {
            let end = logical_cols.get(li + 1).copied().unwrap_or(max_phys + 1);
            for p in start..end.min(max_phys + 1) {
                m.insert(p, li);
            }
        }
        m
    };

    // 3. rows を論理列数の配列に縮小（各論理列の先頭物理列の値を使用）
    let new_rows: Vec<Vec<String>> = rows
        .iter()
        .map(|row| {
            let mut new_row = vec![String::new(); n_logical];
            for (ci, val) in row.iter().enumerate() {
                if let Some(&li) = phys_to_logical.get(&ci) {
                    // 先勝ち：論理列がまだ空の場合のみ書き込む
                    if new_row[li].is_empty() && !val.is_empty() {
                        new_row[li] = val.clone();
                    }
                }
            }
            new_row
        })
        .collect();

    // 4. merges を論理座標に変換
    //    物理 col → logical_start
    //    物理 col+colspan（排他端）→ 論理列インデックスで何論理列分か
    let new_merges: Vec<(usize, usize, usize, usize)> = merges
        .iter()
        .filter_map(|&(r, c, rs, cs)| {
            let logical_start = *phys_to_logical.get(&c)?;
            // 排他端 c+cs が論理列境界に収まる位置を探す
            let phys_end = c + cs;
            let logical_end = logical_cols.partition_point(|&x| x < phys_end);
            let logical_cs = logical_end.saturating_sub(logical_start);
            if rs > 1 || logical_cs > 1 {
                Some((r, logical_start, rs, logical_cs))
            } else {
                None
            }
        })
        .collect();

    (new_rows, new_merges)
}

// ─────────────────────────────────────────────────────────────────────────────
// 階層インデント統合
// ─────────────────────────────────────────────────────────────────────────────

/// 隣接する「排他的列群」を1列に統合し、位置に応じたインデントを付与する。
///
/// 排他的列群 = いずれの行でも、群の中に非空・非カバーセルが最大1つしかない列の連続。
/// これが成立する場合、左端から何番目のサブ列に値があるかがネスト深さを示す（神エクセル）。
///
/// `indent` はインデント単位文字列（例: "　" 全角スペース）。
fn merge_exclusive_cols(
    rows: &[Vec<String>],
    merges: &[(usize, usize, usize, usize)],
    indent: &str,
) -> (Vec<Vec<String>>, Vec<(usize, usize, usize, usize)>) {
    let n_cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    if n_cols <= 1 {
        return (rows.to_vec(), merges.to_vec());
    }

    let (_, covered) = build_span_and_covered(merges);

    // row r, col c が「アクティブセル」かどうか
    // = カバーされていない かつ 非空
    let is_active = |r: usize, c: usize| -> bool {
        !covered.contains(&(r, c))
            && rows.get(r).and_then(|row| row.get(c)).map(|s| !s.is_empty()).unwrap_or(false)
    };

    // 隣接ペアの排他性を先に計算する
    // col c と col c+1 が「ペアとして排他的」= いずれの行でも両方アクティブにならない
    let pairwise_exclusive: Vec<bool> = (0..n_cols.saturating_sub(1))
        .map(|c| (0..rows.len()).all(|r| !(is_active(r, c) && is_active(r, c + 1))))
        .collect();

    // 隣接ペアが排他的かつグループ全体でも排他性を維持する場合にのみ拡張
    //
    // これにより「項番列（全行でアクティブ）」が深さ指示列と誤って統合されるのを防ぐ。
    // 深さ指示列同士（col8, col9, col10...）はペア排他かつグループ排他なので統合される。
    let mut groups: Vec<Vec<usize>> = vec![vec![0]];

    for c in 1..n_cols {
        // ステップ1: 直前列とのペア排他性チェック（高速フィルタ）
        let pair_ok = pairwise_exclusive.get(c - 1).copied().unwrap_or(false);

        // ステップ2: ペア排他が成立する場合のみグループ全体の排他性を確認
        let can_extend = pair_ok && {
            let last_group = groups.last().unwrap();

            // 2a. グループ全体の排他性: いずれの行でもアクティブセルは最大1つ
            let group_exclusive = (0..rows.len()).all(|r| {
                let active_in_group =
                    last_group.iter().filter(|&&gc| is_active(r, gc)).count();
                active_in_group + usize::from(is_active(r, c)) <= 1
            });

            // 2b. rowspanコンフリクト検査
            //     グループ内の列に rowspan カバーがある行で、追加列 c がアクティブになる場合は不可。
            //     例: col0 に rowspan4 があり col1 がそのカバー範囲でアクティブ → マージ不可。
            let no_rowspan_conflict = (0..rows.len()).all(|r| {
                let group_covers_row =
                    last_group.iter().any(|&gc| covered.contains(&(r, gc)));
                !(group_covers_row && is_active(r, c))
            });

            group_exclusive && no_rowspan_conflict
        };

        if can_extend {
            groups.last_mut().unwrap().push(c);
        } else {
            groups.push(vec![c]);
        }
    }

    // グループが全て単一列なら変換不要
    if groups.iter().all(|g| g.len() == 1) {
        return (rows.to_vec(), merges.to_vec());
    }

    let n_new = groups.len();

    // 旧論理列 → 新グループインデックスのマップ
    let old_to_new: Vec<usize> = {
        let mut m = vec![0usize; n_cols];
        for (gi, group) in groups.iter().enumerate() {
            for &c in group {
                m[c] = gi;
            }
        }
        m
    };

    // 旧論理列 → グループ内深さ
    let col_depth: Vec<usize> = {
        let mut d = vec![0usize; n_cols];
        for group in &groups {
            for (depth, &c) in group.iter().enumerate() {
                d[c] = depth;
            }
        }
        d
    };

    // 新 rows: カバーされていない非空セルを基に、インデント付きで新グループ列へ書き込む
    let new_rows: Vec<Vec<String>> = rows
        .iter()
        .enumerate()
        .map(|(ri, row)| {
            let mut new_row = vec![String::new(); n_new];
            for c in 0..n_cols {
                if covered.contains(&(ri, c)) {
                    continue;
                }
                let val = row.get(c).map(|s| s.as_str()).unwrap_or("");
                if val.is_empty() {
                    continue;
                }
                let gi = old_to_new[c];
                if new_row[gi].is_empty() {
                    let depth = col_depth[c];
                    new_row[gi] = format!("{}{}", indent.repeat(depth), val);
                }
            }
            new_row
        })
        .collect();

    // 新 merges: 旧列座標を新グループ座標に変換
    // colspan は「開始グループ」から「終了グループ」までの範囲に変換
    let new_merges: Vec<(usize, usize, usize, usize)> = merges
        .iter()
        .filter_map(|&(r, c, rs, cs)| {
            if c >= n_cols {
                return None;
            }
            let gi_start = old_to_new[c];
            let end_col = (c + cs).saturating_sub(1).min(n_cols - 1);
            let gi_end = old_to_new[end_col];
            let new_cs = gi_end - gi_start + 1;
            if rs > 1 || new_cs > 1 {
                Some((r, gi_start, rs, new_cs))
            } else {
                None
            }
        })
        .collect();

    (new_rows, new_merges)
}

// ─────────────────────────────────────────────────────────────────────────────
// ユーティリティ
// ─────────────────────────────────────────────────────────────────────────────

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
