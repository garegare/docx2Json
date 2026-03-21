use std::collections::{BTreeSet, HashMap, HashSet};
use std::fmt::Write as _;
use std::path::PathBuf;

use anyhow::{Context, Result};

use crate::models::{CellMerge, Document, Element, Section};

/// テーブル変換関数が返す (rows, merges) のペア型
type TableData = (Vec<Vec<String>>, Vec<CellMerge>);

/// `build_span_and_covered` の戻り値型（span_map, covered）
type SpanInfo = (HashMap<(usize, usize), (usize, usize)>, HashSet<(usize, usize)>);

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
            // NOTE/WARNING/TIP/CodeBlock/Quote 以外は [ をエスケープして出力
            let adoc = match &metadata.role {
                Some(crate::models::SemanticRole::Note) => {
                    format!("NOTE: {}\n", escape_para(text))
                }
                Some(crate::models::SemanticRole::Warning) => {
                    format!("WARNING: {}\n", escape_para(text))
                }
                Some(crate::models::SemanticRole::Tip) => {
                    format!("TIP: {}\n", escape_para(text))
                }
                Some(crate::models::SemanticRole::CodeBlock) => {
                    // listing ブロック内はリテラル扱いのためエスケープ不要
                    format!("[source]\n----\n{}\n----\n", text)
                }
                Some(crate::models::SemanticRole::Quote) => {
                    // quote ブロック内もリテラル扱い
                    format!("[quote]\n____\n{}\n____\n", text)
                }
                _ => format!("{}\n", escape_para(text)),
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

            // Step 2b: 幽霊列（全セルが空またはカバー済みで自身はマージ開始でない）を除去
            let (rows, merges) = remove_phantom_cols(&rows, &merges);

            // Step 3: 全空行を除去（末尾だけでなく中間の rowspan 端数行も対象）
            let (rows, merges) = remove_empty_rows(&rows, &merges);

            let col_count = rows.iter().map(|r| r.len()).max().unwrap_or(0);
            if col_count == 0 {
                return;
            }

            // Step 4: 全幅ヘッダー行で分割して出力
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
) -> SpanInfo {
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
///
/// 分割後の 2 番目以降のサブテーブルには最初のデータセグメントの行を
/// AsciiDoc ヘッダー行として繰り返す（列名の文脈を維持するため）。
fn write_table_with_headers(
    out: &mut String,
    rows: &[Vec<String>],
    merges: &[(usize, usize, usize, usize)],
    col_count: usize,
) {
    let (span_map, covered) = build_span_and_covered(merges);

    // 各行が「全幅ヘッダー行」かどうかを判定。
    //
    // 全幅ヘッダー行の条件:
    //   1. 有効な分割点: 直前の行からの rowspan がこの行をまたがない
    //   2. 非空・非カバーセルが1つだけ
    //   3. テーブルが複数列（単列テーブルでは誤検知が多い）
    //
    // 条件1 は colspan なしで行全体に値が1つだけの場合（例: 「(1) 学校数…」）にも
    // 機能するよう、colspan 要件を廃止した。
    let is_header_row = |ri: usize| -> Option<String> {
        // 1. 有効な分割点チェック: 直前 rowspan がこの行をまたがないこと
        let valid_split = merges.iter().all(|&(r, _, rs, _)| !(r < ri && r + rs > ri));
        if !valid_split {
            return None;
        }
        // 2. 非空・非カバーセルが1つだけ
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
        // 3. 複数列テーブルのみ（単列は誤検知を避ける）
        if col_count <= 1 {
            return None;
        }
        let c = visible_nonempty[0];
        rows.get(ri)
            .and_then(|r| r.get(c))
            .filter(|v| !v.is_empty())
            .map(|v| v.replace('\r', "").replace('\n', " "))
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

    // 最初のデータセグメントの行インデックス（サブテーブルのヘッダー繰り返し用）
    let first_data_rows: &[usize] = segments
        .iter()
        .find(|(h, rows)| h.is_none() && !rows.is_empty())
        .map(|(_, rows)| rows.as_slice())
        .unwrap_or(&[]);

    // セグメントを出力
    let mut data_seg_idx = 0usize;
    for (header, row_indices) in &segments {
        if let Some(text) = header {
            writeln!(out, "*{}*\n", escape_cell(text)).unwrap();
        }
        if !row_indices.is_empty() {
            let cols: Vec<&str> = vec!["1"; col_count];
            writeln!(out, r#"[cols="{}",options="header"]"#, cols.join(",")).unwrap();
            writeln!(out, "|===").unwrap();

            if data_seg_idx == 0 {
                // 最初のデータセグメント: 行[0] が AsciiDoc ヘッダー行
                for (seg_idx, &ri) in row_indices.iter().enumerate() {
                    if seg_idx == 1 {
                        writeln!(out).unwrap(); // ヘッダー行とデータ行の間の空行
                    }
                    write_row(out, rows, ri, col_count, &covered, &span_map);
                }
                data_seg_idx += 1;
            } else {
                // 2 番目以降のデータセグメント:
                //   最初のデータセグメントの行をヘッダーとして繰り返す。
                //   そのセグメント内に閉じる rowspan のみ保持し、
                //   データ行は別の covered/span_map を使う。
                let max_tmpl_ri = first_data_rows.iter().copied().max().unwrap_or(0);
                let tmpl_set: HashSet<usize> = first_data_rows.iter().copied().collect();
                let tmpl_merges: Vec<_> = merges
                    .iter()
                    .filter(|&&(r, _, rs, _)| tmpl_set.contains(&r) && r + rs <= max_tmpl_ri + 1)
                    .copied()
                    .collect();
                let (tmpl_span_map, tmpl_covered) = build_span_and_covered(&tmpl_merges);

                // データ行専用の covered/span_map（テンプレート行のマージを除外）
                let data_set: HashSet<usize> = row_indices.iter().copied().collect();
                let data_merges: Vec<_> = merges
                    .iter()
                    .filter(|&&(r, _, _, _)| data_set.contains(&r))
                    .copied()
                    .collect();
                let (data_span_map, data_covered) = build_span_and_covered(&data_merges);

                // ヘッダーテンプレート行を AsciiDoc ヘッダー行として出力
                for &ri in first_data_rows {
                    write_row(out, rows, ri, col_count, &tmpl_covered, &tmpl_span_map);
                }
                writeln!(out).unwrap(); // ヘッダー行とデータ行の間の空行

                // 実データ行を出力
                for &ri in row_indices {
                    write_row(out, rows, ri, col_count, &data_covered, &data_span_map);
                }
                data_seg_idx += 1;
            }

            writeln!(out, "|===").unwrap();
            writeln!(out).unwrap();
        }
    }
}

/// 1行分のセルを AsciiDoc テーブル記法で出力するヘルパー
fn write_row(
    out: &mut String,
    rows: &[Vec<String>],
    ri: usize,
    col_count: usize,
    covered: &HashSet<(usize, usize)>,
    span_map: &HashMap<(usize, usize), (usize, usize)>,
) {
    for ci in 0..col_count {
        if covered.contains(&(ri, ci)) {
            continue;
        }
        let text = escape_cell(
            rows.get(ri)
                .and_then(|r| r.get(ci))
                .map(|s| s.as_str())
                .unwrap_or(""),
        );
        let prefix = span_map
            .get(&(ri, ci))
            .map(|&(rs, cs)| cell_span_prefix(cs, rs))
            .unwrap_or_default();
        writeln!(out, "{}| {}", prefix, text).unwrap();
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// 全空行の除去
// ─────────────────────────────────────────────────────────────────────────────

/// 「全空行」をテーブルから除去し、rowspan を調整する。
///
/// 全空行 = すべてのセルが「カバー済み」または「空文字列」である行。
///
/// - 末尾だけでなく中間の空行も除去する（rowspan で列がカバーされているが
///   残りの列が空というパターン、例: 複合ヘッダー行の継続行）。
/// - 除去した行にまたがる rowspan はその分だけ短縮する。
fn remove_empty_rows(
    rows: &[Vec<String>],
    merges: &[(usize, usize, usize, usize)],
) -> TableData {
    let n_rows = rows.len();
    let n_cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    if n_rows == 0 {
        return (rows.to_vec(), merges.to_vec());
    }
    let (_, covered) = build_span_and_covered(merges);

    // 各行が「全空行」かどうか（全列が空またはカバー済み）
    let is_empty: Vec<bool> = (0..n_rows)
        .map(|ri| {
            (0..n_cols).all(|ci| {
                covered.contains(&(ri, ci))
                    || rows.get(ri).and_then(|r| r.get(ci)).map(|v| v.is_empty()).unwrap_or(true)
            })
        })
        .collect();

    if is_empty.iter().all(|&e| !e) {
        return (rows.to_vec(), merges.to_vec());
    }

    // 残す行の旧番号 → 新番号マッピング
    let mut old_to_new = vec![usize::MAX; n_rows];
    let mut new_ri = 0usize;
    for (ri, &empty) in is_empty.iter().enumerate() {
        if !empty {
            old_to_new[ri] = new_ri;
            new_ri += 1;
        }
    }

    let new_rows: Vec<Vec<String>> = rows
        .iter()
        .enumerate()
        .filter(|&(ri, _)| !is_empty[ri])
        .map(|(_, row)| row.clone())
        .collect();

    // rowspan 調整:
    //   開始行が除去された場合はマージ自体を削除
    //   それ以外は rowspan を「残る行の数」に縮小し、行番号を再マッピング
    let new_merges: Vec<_> = merges
        .iter()
        .filter_map(|&(r, c, rs, cs)| {
            if is_empty[r] {
                return None; // 開始行が除去
            }
            let new_r = old_to_new[r];
            // [r, r+rs) の範囲内で残す行の数
            let new_rs = (r..r + rs).filter(|&i| i < n_rows && !is_empty[i]).count();
            if new_rs > 1 || cs > 1 {
                Some((new_r, c, new_rs, cs))
            } else {
                None
            }
        })
        .collect();

    (new_rows, new_merges)
}

// ─────────────────────────────────────────────────────────────────────────────
// 幽霊列の除去
// ─────────────────────────────────────────────────────────────────────────────

/// 全セルが「カバー済みか空」である「幽霊列（phantom column）」を除去する。
///
/// 神エクセル処理後、ヘッダー行だけを対象としたマージの境界として生成された
/// 論理列が、データ行では空のまま残る場合がある。これを除去して余分な空列を消す。
///
/// - マージ開始位置が除去対象列の場合は、マージを次の実列に移動してcolspanを調整する。
/// - マージが除去対象列をまたぐ場合は、その分だけcolspanを縮小する。
fn remove_phantom_cols(
    rows: &[Vec<String>],
    merges: &[(usize, usize, usize, usize)],
) -> TableData {
    let n_cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    if n_cols == 0 {
        return (rows.to_vec(), merges.to_vec());
    }

    let (span_map, covered) = build_span_and_covered(merges);

    // 幽霊列: 以下の 2 条件をともに満たす列
    //
    //   条件1: 全行において列 c が次のいずれか
    //     a) カバー済み
    //     b) 空
    //     c) colspan > 1 のマージ開始セル（値が右方向へ広がる）
    //
    //   条件2: rowspan > 1 のマージがこの列から開始しない
    //     ← rowspan を持つ列は「複数行にまたがる見出し」を担う実列
    //
    // 条件1 のみでは「データマージ (cs=2) の開始列」も幽霊と誤判定するため、
    // rowspan チェックで本物のデータ列を除外する。
    //
    // 例: リクエスト表の lc1（パラメータ名）は cs=2 のデータマージ起点だが
    //     (3,1,rs=2,cs=2) という rowspan マージも持つため幽霊とならない。
    //     lc2（app_submit 起点）は rs=1 のみ → 幽霊。
    let has_rowspan_start: Vec<bool> = (0..n_cols)
        .map(|c| merges.iter().any(|&(_, mc, rs, _)| mc == c && rs > 1))
        .collect();

    let is_phantom: Vec<bool> = (0..n_cols)
        .map(|c| {
            // 条件2: rowspan 起点があれば幽霊ではない
            if has_rowspan_start[c] {
                return false;
            }
            // 条件1: 全行が a/b/c のいずれか
            (0..rows.len()).all(|r| {
                if covered.contains(&(r, c)) {
                    return true;
                }
                let empty = rows
                    .get(r)
                    .and_then(|row| row.get(c))
                    .map(|s| s.is_empty())
                    .unwrap_or(true);
                if empty {
                    return true;
                }
                // 非空・非カバー: colspan > 1 のマージ起点の場合のみ幽霊と見なす
                span_map.get(&(r, c)).map(|&(_, cs)| cs > 1).unwrap_or(false)
            })
        })
        .collect();

    if is_phantom.iter().all(|&p| !p) {
        return (rows.to_vec(), merges.to_vec());
    }

    // 旧列 → 新列インデックス（幽霊列は usize::MAX）
    let mut old_to_new = vec![usize::MAX; n_cols];
    let mut ni = 0;
    for (c, &ph) in is_phantom.iter().enumerate() {
        if !ph {
            old_to_new[c] = ni;
            ni += 1;
        }
    }
    let n_new = ni;
    if n_new == n_cols {
        return (rows.to_vec(), merges.to_vec());
    }

    // 新 rows: 幽霊列を除去した値を引き継ぐ
    let mut new_rows: Vec<Vec<String>> = rows
        .iter()
        .map(|row| {
            (0..n_cols)
                .filter(|&c| !is_phantom[c])
                .map(|c| row.get(c).cloned().unwrap_or_default())
                .collect()
        })
        .collect();

    // 幽霊列がマージ開始の場合、その値を次の実列へ移動する。
    // （値は幽霊列にのみ存在し、新 rows では消えてしまうため）
    for &(r, c, _, cs) in merges {
        if !is_phantom[c] {
            continue;
        }
        let val = rows.get(r).and_then(|row| row.get(c)).cloned().unwrap_or_default();
        if val.is_empty() {
            continue;
        }
        // [c, c+cs) の範囲で最初の実列へ移動
        if let Some(first_real) = (c..c + cs).find(|&col| col < n_cols && !is_phantom[col]) {
            let new_c = old_to_new[first_real];
            if r < new_rows.len() && new_c < new_rows[r].len() && new_rows[r][new_c].is_empty() {
                new_rows[r][new_c] = val;
            }
        }
    }

    // 新 merges:
    //   開始列が幽霊の場合は次の実列にシフトして colspan を縮小
    //   開始列が実の場合は colspan の範囲内の実列数に縮小
    let new_merges: Vec<_> = merges
        .iter()
        .filter_map(|&(r, c, rs, cs)| {
            // [c, c+cs) の範囲内で最初の実列を探す
            let first_real = (c..c + cs).find(|&col| col < n_cols && !is_phantom[col])?;
            let new_c = old_to_new[first_real];
            // 実列の数
            let new_cs = (first_real..c + cs)
                .filter(|&col| col < n_cols && !is_phantom[col])
                .count();
            if rs > 1 || new_cs > 1 {
                Some((r, new_c, rs, new_cs))
            } else {
                None
            }
        })
        .collect();

    (new_rows, new_merges)
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
) -> TableData {
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
        // 2行以上で出現するスタンドアロン列を境界として追加する。
        // 閾値を 2 とするのは、1行だけの単独値はノイズ（空きセルや一時的な値）の
        // 可能性が高く、誤った境界挿入でテーブル列が増殖するのを防ぐため。
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
) -> TableData {
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

/// AsciiDoc テーブルセル内の特殊文字をエスケープする
///
/// - `|`  → `\|`  （セル区切り文字）
/// - `[`  → `\[`  （属性リストとして誤解釈されるのを防ぐ）
/// - `\n` → ` +\n`（AsciiDoc ハードラインブレーク記法で改行を保持）
/// - `\r` は除去
fn escape_cell(s: &str) -> String {
    s.replace('\r', "")
        .replace('|', "\\|")
        .replace('[', "\\[")
        .replace('\n', " +\n")
}

/// AsciiDoc 段落テキスト内の特殊文字をエスケープする
///
/// テーブルセル外の段落で `[重要]` などが属性リストとして誤解釈されるのを防ぐ。
fn escape_para(s: &str) -> String {
    s.replace('\r', "").replace('[', "\\[")
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rows(data: &[&[&str]]) -> Vec<Vec<String>> {
        data.iter()
            .map(|r| r.iter().map(|s| s.to_string()).collect())
            .collect()
    }

    // ─── cell_span_prefix ───

    #[test]
    fn test_cell_span_prefix_none() {
        assert_eq!(cell_span_prefix(1, 1), "");
    }

    #[test]
    fn test_cell_span_prefix_colspan() {
        assert_eq!(cell_span_prefix(3, 1), "3+");
    }

    #[test]
    fn test_cell_span_prefix_rowspan() {
        assert_eq!(cell_span_prefix(1, 2), ".2+");
    }

    #[test]
    fn test_cell_span_prefix_both() {
        assert_eq!(cell_span_prefix(2, 3), "2.3+");
    }

    // ─── escape_cell ───

    #[test]
    fn test_escape_cell_pipe() {
        assert_eq!(escape_cell("a|b"), "a\\|b");
    }

    #[test]
    fn test_escape_cell_bracket() {
        assert_eq!(escape_cell("[test]"), "\\[test]");
    }

    #[test]
    fn test_escape_cell_newline() {
        assert_eq!(escape_cell("a\nb"), "a +\nb");
    }

    #[test]
    fn test_escape_cell_cr_removed() {
        assert_eq!(escape_cell("a\r\nb"), "a +\nb");
    }

    // ─── remove_empty_rows ───

    #[test]
    fn test_remove_empty_rows_no_empty() {
        let r = rows(&[&["A", "B"], &["1", "2"]]);
        let (new_rows, new_merges) = remove_empty_rows(&r, &[]);
        assert_eq!(new_rows.len(), 2);
        assert!(new_merges.is_empty());
    }

    #[test]
    fn test_remove_empty_rows_trailing() {
        // 末尾の全空行を除去する
        let r = rows(&[&["A", "B"], &["", ""]]);
        let (new_rows, _) = remove_empty_rows(&r, &[]);
        assert_eq!(new_rows.len(), 1);
        assert_eq!(new_rows[0], vec!["A", "B"]);
    }

    #[test]
    fn test_remove_empty_rows_intermediate() {
        // 中間の全空行（rowspan カバーのみの行）も除去する
        // row0: ["A", "B"]  merges: (0,0,2,1) → row1,row2 の col0 はカバー済み
        // row1: ["",  "X"]  col0=covered, col1=非空 → NOT all empty
        // row2: ["",  ""]   col0=covered, col1=空   → all empty → 除去
        let r = rows(&[&["A", "B"], &["", "X"], &["", ""]]);
        let merges = vec![(0usize, 0usize, 3usize, 1usize)]; // (row,col,rowspan,colspan)
        let (new_rows, new_merges) = remove_empty_rows(&r, &merges);
        assert_eq!(new_rows.len(), 2);
        // rowspan は 2 に縮小される
        assert!(new_merges.iter().any(|&(r, c, rs, cs)| r == 0 && c == 0 && rs == 2 && cs == 1));
    }

    // ─── compress_columns ───

    #[test]
    fn test_compress_columns_no_merges_removes_empty_cols() {
        // 列1が全行空 → 除去
        let r = rows(&[&["A", "", "C"], &["1", "", "3"]]);
        let (new_rows, _) = compress_columns(&r, &[]);
        assert_eq!(new_rows[0], vec!["A", "C"]);
        assert_eq!(new_rows[1], vec!["1", "3"]);
    }

    #[test]
    fn test_compress_columns_no_merges_all_used() {
        let r = rows(&[&["A", "B"], &["1", "2"]]);
        let (new_rows, _) = compress_columns(&r, &[]);
        assert_eq!(new_rows.len(), 2);
        assert_eq!(new_rows[0].len(), 2);
    }

    #[test]
    fn test_compress_columns_with_merges_title_compressed() {
        // タイトルが物理列0〜2 (colspan=3) にまたがるケース。
        // row1 の A,B,C は各1回のみ出現 (freq<2) → 境界に追加されない。
        // 結果: 論理列 [0, 3] の2列, タイトルは論理列0 に収まり colspan=1 → merges なし
        let r = rows(&[&["タイトル", "", ""], &["A", "B", "C"]]);
        let merges = vec![(0usize, 0usize, 1usize, 3usize)]; // (row,col,rowspan,colspan)
        let (new_rows, new_merges) = compress_columns(&r, &merges);
        assert_eq!(new_rows[0][0], "タイトル");
        // 3物理列が1論理列に圧縮されたため colspan=1 → マージ不要
        assert!(new_merges.is_empty());
    }

    #[test]
    fn test_compress_columns_with_merges_rowspan_preserved() {
        // rowspan=2 のマージが論理座標に変換されて保持されるケース
        // row0: "ラベル" at col0, rowspan=2 / "値1" at col1
        // row1: ""       at col0 (カバー済み)   / "値2" at col1
        let r = rows(&[&["ラベル", "値1"], &["", "値2"]]);
        let merges = vec![(0usize, 0usize, 2usize, 1usize)]; // rowspan=2, colspan=1
        let (new_rows, new_merges) = compress_columns(&r, &merges);
        assert_eq!(new_rows[0].len(), 2); // 論理列は2列
        // rowspan=2 は保持される
        assert!(new_merges.iter().any(|&(_, _, rs, _)| rs == 2));
    }

    // ─── remove_phantom_cols ───

    #[test]
    fn test_remove_phantom_cols_removes_all_empty_col() {
        // 列1が全行空 → 幽霊列として除去
        let r = rows(&[&["A", "", "C"], &["1", "", "3"]]);
        let (new_rows, _) = remove_phantom_cols(&r, &[]);
        assert_eq!(new_rows[0], vec!["A", "C"]);
    }

    #[test]
    fn test_remove_phantom_cols_keeps_rowspan_col() {
        // 列0が rowspan>1 のマージ起点 → 幽霊ではない
        let r = rows(&[&["A", "B"], &["", "C"]]);
        let merges = vec![(0usize, 0usize, 2usize, 1usize)]; // rowspan=2
        let (new_rows, _) = remove_phantom_cols(&r, &merges);
        // 列0は rowspan起点なので保持
        assert_eq!(new_rows[0].len(), 2);
    }

    // ─── merge_exclusive_cols ───

    #[test]
    fn test_merge_exclusive_cols_single_col_no_change() {
        let r = rows(&[&["A"], &["B"]]);
        let (new_rows, _) = merge_exclusive_cols(&r, &[], "　");
        assert_eq!(new_rows, r);
    }

    #[test]
    fn test_merge_exclusive_cols_exclusive_pair_merged() {
        // 列0と列1が排他的（各行でどちらか一方のみ非空）→ 1列に統合
        let r = rows(&[&["A", ""], &["", "B"], &["C", ""]]);
        let (new_rows, _) = merge_exclusive_cols(&r, &[], "　");
        assert_eq!(new_rows[0].len(), 1);
        assert_eq!(new_rows[0][0], "A");
        // 列1のものはインデント付き（depth=1）
        assert_eq!(new_rows[1][0], "　B");
    }

    #[test]
    fn test_merge_exclusive_cols_non_exclusive_not_merged() {
        // 両方が非空の行がある場合は統合しない
        let r = rows(&[&["A", "B"], &["C", "D"]]);
        let (new_rows, _) = merge_exclusive_cols(&r, &[], "　");
        assert_eq!(new_rows[0].len(), 2);
    }
}
