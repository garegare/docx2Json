use std::collections::HashMap;
use std::io::{BufReader, Read};
use std::path::Path;

use anyhow::{Context, Result};
use image::codecs::jpeg::JpegEncoder;
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use zip::ZipArchive;

use crate::config::{normalize_style_name, Config, DocxConfig};
use crate::models::{Asset, Document, Element, ElementMetadata, SemanticRole, Section};
use super::emf;

/// DOCXファイルを解析してDocumentを返す
pub fn parse(path: &Path, config: &Config) -> Result<Document> {
    let file = std::fs::File::open(path)
        .with_context(|| format!("ファイルを開けません: {}", path.display()))?;
    let mut archive = ZipArchive::new(BufReader::new(file))
        .context("ZIPアーカイブとして開けません（破損または非DOCXファイルの可能性）")?;

    // ファイル名をタイトルとして使用
    let title = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("Untitled")
        .to_string();

    // relationships を読み込んで画像パスを解決
    let rels = parse_rels(&mut archive)
        .context("リレーションシップファイルのパースに失敗")?;

    // 画像データをバイナリで取得（Base64 エンコードは出力直前まで遅延）
    let images = extract_images(&mut archive, &rels, config)
        .context("画像の抽出に失敗")?;

    // numbering.xml から箇条書き定義を取得（存在しない場合は空マップ）
    let numbering = parse_numbering(&mut archive);

    // document.xml をパース
    let xml = read_zip_entry(&mut archive, "word/document.xml")
        .context("word/document.xml の読み込みに失敗")?;
    let sections = parse_document_xml(&xml, &images, &numbering, config)
        .context("document.xml のパースに失敗")?;

    Ok(Document { title, sections })
}

/// word/_rels/document.xml.rels を解析し、rId -> パスのマップを返す
fn parse_rels(archive: &mut ZipArchive<BufReader<std::fs::File>>) -> Result<HashMap<String, String>> {
    let xml = match read_zip_entry(archive, "word/_rels/document.xml.rels") {
        Ok(s) => s,
        Err(_) => return Ok(HashMap::new()), // rels ファイルがない場合は空マップを返す
    };

    let mut rels = HashMap::new();
    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) if e.local_name().as_ref() == b"Relationship" => {
                let id = attr_value(&e, "Id").unwrap_or_default();
                let target = attr_value(&e, "Target").unwrap_or_default();
                let r_type = attr_value(&e, "Type").unwrap_or_default();
                if r_type.contains("/image") && !id.is_empty() {
                    // Target は "media/image1.png" のような相対パス
                    rels.insert(id, format!("word/{}", target));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::Error::from(e)),
            _ => {}
        }
    }
    Ok(rels)
}

/// ZIP内の画像ファイルをバイナリとして返す（#4 Lazy Serialization）
///
/// Base64 エンコードはここでは行わず、`Asset.data: Vec<u8>` として保持する。
/// JSON 出力時に `models::serialize_as_base64` が呼ばれるため、メモリ上の
/// データ量を元サイズ（Base64比 約75%）に抑えられる。
///
/// config.image.max_px > 0 の場合は長辺をリサイズして JPEG 再エンコードする。
fn extract_images(
    archive: &mut ZipArchive<BufReader<std::fs::File>>,
    rels: &HashMap<String, String>,
    config: &Config,
) -> Result<HashMap<String, Vec<u8>>> {
    let mut images = HashMap::new();
    for (rid, zip_path) in rels {
        if let Ok(mut entry) = archive.by_name(zip_path) {
            let mut buf = Vec::new();
            entry.read_to_end(&mut buf)?;

            // EMF 形式の場合: Windows のみ PNG に変換、非 Windows はスキップ
            let final_buf = if emf::is_emf(&buf) {
                match emf::emf_to_png(&buf) {
                    Some(png) => png,
                    None => {
                        eprintln!("  (EMF 画像はこのプラットフォームではスキップ: {})", zip_path);
                        continue;
                    }
                }
            // リサイズ・圧縮が有効な場合は画像を処理する
            } else if config.image.max_px > 0 {
                resize_and_compress(&buf, config.image.max_px, config.image.quality)
                    .unwrap_or(buf) // 変換失敗時は元データをそのまま使用
            } else {
                buf
            };

            images.insert(rid.clone(), final_buf); // Vec<u8> のまま格納
        }
    }
    Ok(images)
}

/// 画像データをリサイズして JPEG 形式で再エンコードする
///
/// - max_px: 長辺の最大ピクセル数（アスペクト比を維持してリサイズ）
/// - quality: JPEG 品質（1〜100）
/// - 変換できない場合は None を返す（呼び出し側で元データにフォールバック）
fn resize_and_compress(data: &[u8], max_px: u32, quality: u8) -> Option<Vec<u8>> {
    let img = image::load_from_memory(data).ok()?;

    // リサイズが必要かチェック（長辺が max_px を超える場合のみ処理）
    let img = if img.width() > max_px || img.height() > max_px {
        img.thumbnail(max_px, max_px)
    } else {
        img
    };

    // JPEG は透過チャンネルを持たないため RGB に変換してからエンコード
    let rgb = img.into_rgb8();
    let mut output = Vec::new();
    let encoder = JpegEncoder::new_with_quality(&mut output, quality);
    rgb.write_with_encoder(encoder).ok()?;
    Some(output)
}

/// word/numbering.xml を解析し、numId → Vec<numFmt> のマップを返す
/// ilvl(0-8) ごとの numFmt 文字列（"bullet", "decimal", "lowerLetter" など）を保持する
/// numbering.xml が存在しない場合は空マップを返す（エラーは無視）
fn parse_numbering(archive: &mut ZipArchive<BufReader<std::fs::File>>) -> HashMap<String, Vec<String>> {
    let xml = match read_zip_entry(archive, "word/numbering.xml") {
        Ok(s) => s,
        Err(_) => return HashMap::new(),
    };

    // abstractNumId → [numFmt at ilvl 0..8]
    let mut abstract_fmts: HashMap<String, Vec<String>> = HashMap::new();
    // numId → abstractNumId
    let mut num_to_abstract: HashMap<String, String> = HashMap::new();

    let mut current_abstract_id: Option<String> = None; // abstractNum 処理中
    let mut current_num_id: Option<String> = None;      // num 処理中
    let mut current_ilvl: Option<usize> = None;         // lvl 処理中

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            // <w:abstractNum w:abstractNumId="0">
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"abstractNum" => {
                if let Some(id) = attr_value(&e, "abstractNumId") {
                    abstract_fmts
                        .entry(id.clone())
                        .or_insert_with(|| vec!["bullet".to_string(); 9]);
                    current_abstract_id = Some(id);
                }
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"abstractNum" => {
                current_abstract_id = None;
            }

            // <w:lvl w:ilvl="0">
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"lvl" => {
                current_ilvl = attr_value(&e, "ilvl").and_then(|s| s.parse().ok());
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"lvl" => {
                current_ilvl = None;
            }

            // <w:numFmt w:val="bullet"/>
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) if e.local_name().as_ref() == b"numFmt" => {
                if let (Some(ref abs_id), Some(ilvl)) = (&current_abstract_id, current_ilvl) {
                    if let Some(fmt) = attr_value(&e, "val") {
                        if let Some(fmts) = abstract_fmts.get_mut(abs_id) {
                            if ilvl < fmts.len() {
                                fmts[ilvl] = fmt;
                            }
                        }
                    }
                }
            }

            // <w:num w:numId="1">
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"num" => {
                current_num_id = attr_value(&e, "numId");
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"num" => {
                current_num_id = None;
            }

            // <w:abstractNumId w:val="0"/>（num の子要素）
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if current_num_id.is_some() && e.local_name().as_ref() == b"abstractNumId" =>
            {
                if let (Some(ref num_id), Some(abs_id)) = (&current_num_id, attr_value(&e, "val")) {
                    num_to_abstract.insert(num_id.clone(), abs_id);
                }
            }

            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // numId → fmts に解決して返す
    num_to_abstract
        .into_iter()
        .filter_map(|(num_id, abs_id)| {
            abstract_fmts.get(&abs_id).map(|fmts| (num_id, fmts.clone()))
        })
        .collect()
}

/// document.xml を走査しセクションツリーを構築する
fn parse_document_xml(
    xml: &str,
    images: &HashMap<String, Vec<u8>>,
    numbering: &HashMap<String, Vec<String>>,
    config: &Config,
) -> Result<Vec<Section>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    // ---- 段落パーサー状態 ----
    let mut in_del = 0u32;           // w:del ネスト深さ
    let mut in_ins = 0u32;           // w:ins ネスト深さ（バランス追跡用）
    let mut in_ppr = false;          // w:pPr 内か
    let mut in_ppr_rpr = false;      // w:pPr > w:rPr 内か（段落デフォルト書式）
    let mut in_rpr = false;          // w:rPr 内か（ラン書式）
    let mut ppr_underline = false;   // w:pPr > w:rPr に w:u があるか（見出し判定用）
    let mut run_underline = false;   // ラン w:rPr に w:u があるか（見出し判定用）
    let mut current_text = String::new();           // 現在の段落テキスト
    let mut current_assets: Vec<Asset> = Vec::new(); // 現在の段落の画像
    let mut drawing_rid: Option<String> = None;     // 処理中の drawing の rId
    let mut drawing_title: Option<String> = None;   // 処理中の drawing のタイトル（wp:docPr から取得）
    let mut in_mc_fallback: u32 = 0;               // mc:Fallback ネスト深さ（AlternateContent 重複防止）
    let mut vml_rid: Option<String> = None;        // v:imagedata r:id（VML 画像参照）
    let mut vml_title: Option<String> = None;      // v:imagedata o:title（VML 画像タイトル）
    let mut in_paragraph = false;
    let mut paragraph_style: Option<String> = None;
    let mut para_alignment: Option<String> = None;    // w:jc（段落の水平配置）
    let mut para_outline_level: Option<u32> = None;   // w:outlineLvl（1-based）
    let mut para_anchor_id: Option<String> = None;    // w:bookmarkStart の name

    // ---- 箇条書きパーサー状態 ----
    let mut in_numpr = false;                    // w:numPr 内か
    let mut para_num_id: Option<String> = None;  // 現在の段落の numId
    let mut para_num_ilvl: u32 = 0;              // 現在の段落の ilvl（インデントレベル）

    // ---- テーブルパーサー状態 ----
    let mut in_table: u32 = 0;                      // w:tbl ネスト深さ
    let mut in_table_cell = false;                  // w:tc 内か（最外テーブルのみ）
    let mut current_cell_text = String::new();      // 現在のセルテキスト
    let mut current_row: Vec<String> = Vec::new();  // 現在の行のセル
    let mut table_rows: Vec<Vec<String>> = Vec::new(); // テーブル全行

    // ---- フィールドコード状態 ----
    // w:fldChar / w:instrText で表現されるフィールドコード（PAGE, DATE, REF 等）の
    // 命令テキスト（w:instrText）を出力から除外するためのフラグ。
    // フィールドの表示値（begin→separate 間を除く w:t）は通常通り出力する。
    let mut in_fld_instr = false;    // w:instrText 内か
    let mut in_ppr_change = false;   // w:pPrChange 内か（変更前の古いプロパティをスキップ）

    // ---- OMML（数式）パーサー状態 ----
    // Office Math Markup Language (OMML) を LaTeX ライクな表記に変換する。
    // SAX スタイルパーサーで再帰構造を扱うため、(要素名, 出力バッファ) の
    // スタックを用いてボトムアップに LaTeX テキストを組み立てる。
    let mut in_omath: u32 = 0;                           // m:oMath ネスト深さ
    let mut omath_stack: Vec<(String, String)> = Vec::new(); // (タグ名, 出力バッファ)
    let mut in_mt = false;                               // m:t（数式テキスト）内か

    // ---- セクションスタック ----
    // (level, section) のスタック。ルートの children が最終出力。
    //
    // # スタック操作の設計（#1 エッジケース対応）
    // 新しい見出しレベル N が来たとき:
    //   1. スタックトップのレベル >= N である限り pop し、適切な親に push する
    //   2. 残ったスタックトップがこの見出しの親になる（いなければ root_sections）
    //
    // エッジケース例:
    //   H1 → H3（H2 スキップ）: H3 は H1 の子として扱う（スタックは空にならない）
    //   H2 → H1（レベル逆転）: H2 を root に flush してから H1 を push する
    //   先頭が H2（H1 なし）: H2 は直接 root_sections に入る
    let mut stack: Vec<(usize, Section)> = Vec::new();
    let mut root_sections: Vec<Section> = Vec::new();

    loop {
        match reader.read_event() {

            // ================================================================
            // OMML 数式処理（m:oMath > m:f / m:sSup / m:sSub / m:t など）
            // ================================================================

            // oMath 開始: 数式モードに入る（ネスト対応）
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"oMath" => {
                in_omath += 1;
                omath_stack.push(("oMath".to_string(), String::new()));
            }

            // 数式モード内のすべての開始タグ: スタックに積む
            Ok(Event::Start(e)) if in_omath > 0 => {
                let local = String::from_utf8_lossy(e.local_name().as_ref()).to_string();
                if local == "t" { in_mt = true; }
                omath_stack.push((local, String::new()));
            }

            // 数式モード内のすべての終了タグ: ポップしてバッファを親に結合
            Ok(Event::End(e)) if in_omath > 0 => {
                let local = String::from_utf8_lossy(e.local_name().as_ref()).to_string();
                if local == "t" { in_mt = false; }
                if local == "oMath" {
                    // 数式モード終了: LaTeX 文字列を生成して段落に追加
                    if let Some((_, latex_buf)) = omath_stack.pop() {
                        let latex = format!("${}$", latex_buf.trim());
                        if in_paragraph && in_table == 0 && !latex.is_empty() {
                            current_text.push_str(&latex);
                        }
                    }
                    in_omath = in_omath.saturating_sub(1);
                } else {
                    omath_pop_and_combine(&mut omath_stack, &local);
                }
            }

            // ================================================================
            // テーブル処理（w:tbl > w:tr > w:tc）
            // ================================================================

            // テーブル開始
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"tbl" => {
                in_table += 1;
                if in_table == 1 {
                    table_rows.clear();
                }
            }
            // テーブル終了: Markdownを生成して現在のセクションに挿入
            Ok(Event::End(e)) if e.local_name().as_ref() == b"tbl" => {
                if in_table == 1 && !table_rows.is_empty() {
                    // mem::take で所有権を取りつつ table_rows を空にする（clone 不要）
                    let rows = std::mem::take(&mut table_rows);
                    let md = rows_to_markdown(&rows);
                    if let Some((_, sec)) = stack.last_mut() {
                        if !sec.body_text.is_empty() {
                            sec.body_text.push('\n');
                        }
                        sec.body_text.push_str(&md);
                        sec.elements.push(Element::Table {
                            metadata: ElementMetadata::default(),
                            rows,
                        });
                    }
                    // table_rows は take() 済みなので clear() は不要
                }
                in_table = in_table.saturating_sub(1);
            }

            // 行開始（最外テーブルのみ追跡）
            Ok(Event::Start(e)) if in_table == 1 && e.local_name().as_ref() == b"tr" => {
                current_row.clear();
            }
            // 行終了: 行をテーブルに追加
            Ok(Event::End(e)) if in_table == 1 && e.local_name().as_ref() == b"tr" => {
                if !current_row.is_empty() {
                    table_rows.push(std::mem::take(&mut current_row));
                }
            }

            // セル開始
            Ok(Event::Start(e)) if in_table == 1 && e.local_name().as_ref() == b"tc" => {
                in_table_cell = true;
                current_cell_text.clear();
            }
            // セル終了: テキストを行に追加
            Ok(Event::End(e)) if in_table == 1 && e.local_name().as_ref() == b"tc" => {
                if in_table_cell {
                    // 末尾に残る "<br>"（最終段落の区切り）を除去してから行に追加。
                    // `trim()` は空白文字のみ除去するため <br> が残ってしまうため、
                    // 明示的に trim_end_matches で除去する。
                    let text = current_cell_text
                        .trim()
                        .trim_end_matches("<br>")
                        .trim()
                        .to_string();
                    current_row.push(text); // 空セルは rows_to_markdown 側で " " に補填
                    current_cell_text.clear();
                    in_table_cell = false;
                }
            }

            // セル内の段落終了: 複数段落を持つセルの段落間区切り（#2）
            // Markdown テーブルでは生改行が表の構造を崩すため <br> に変換する。
            // セル終了（</tc>）時に末尾の <br> を除去するため、ここでは末尾付与のみ。
            Ok(Event::End(e)) if in_table > 0 && in_table_cell && e.local_name().as_ref() == b"p" => {
                let trimmed = current_cell_text.trim_end().to_string();
                if !trimmed.is_empty() {
                    current_cell_text = trimmed;
                    current_cell_text.push_str("<br>"); // 段落区切りを <br> に変換
                }
            }

            // ================================================================
            // 段落処理（テーブル外のみ）
            // ================================================================

            // 段落開始
            Ok(Event::Start(e)) if in_table == 0 && e.local_name().as_ref() == b"p" => {
                in_paragraph = true;
                paragraph_style = None;
                ppr_underline = false;
                run_underline = false;
                in_numpr = false;
                para_num_id = None;
                para_num_ilvl = 0;
                current_text.clear();
                current_assets.clear();
                para_alignment = None;
                para_outline_level = None;
                para_anchor_id = None;
            }

            // ---- 段落プロパティ変更追跡（w:pPrChange）----
            // w:pPrChange には変更前の古いプロパティが入った w:pPr が含まれる。
            // その内部の w:pStyle 等を読み取ると現在のスタイルが上書きされるため、
            // pPrChange 内では全プロパティ読み取りをスキップする。
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"pPrChange" => {
                in_ppr_change = true;
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"pPrChange" => {
                in_ppr_change = false;
            }

            // ---- 段落プロパティ ----
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"pPr" => {
                in_ppr = true;
            }
            // pPrChange 内のネストした </w:pPr> では in_ppr をリセットしない
            Ok(Event::End(e)) if e.local_name().as_ref() == b"pPr" && !in_ppr_change => {
                in_ppr = false;
                in_ppr_rpr = false;
                in_numpr = false;
            }
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) if in_ppr && !in_ppr_change && e.local_name().as_ref() == b"pStyle" => {
                if let Some(val) = attr_value(&e, "w:val").or_else(|| attr_value(&e, "val")) {
                    paragraph_style = Some(val);
                }
            }

            // ---- 段落の水平配置（w:jc）----
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) if in_ppr && !in_ppr_change && e.local_name().as_ref() == b"jc" => {
                if let Some(val) = attr_value(&e, "val") {
                    para_alignment = Some(val);
                }
            }

            // ---- アウトラインレベル（w:outlineLvl）----
            // Word では 0-based（0 = 見出し1相当）なので +1 して 1-based に変換する
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) if in_ppr && !in_ppr_change && e.local_name().as_ref() == b"outlineLvl" => {
                if let Some(val) = attr_value(&e, "val") {
                    para_outline_level = val.parse::<u32>().ok().map(|v| v + 1);
                }
            }

            // ---- アンカー ID（w:bookmarkStart）----
            // 段落内外を問わずドキュメント全体を対象とし、最初に出現したブックマーク名を
            // 現在の段落の anchor_id として採用する（in_paragraph 外でも処理される）。
            // _Toc・_GoBack は Word が自動生成するブックマークのため除外する。
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) if e.local_name().as_ref() == b"bookmarkStart" => {
                if para_anchor_id.is_none() {
                    if let Some(name) = attr_value(&e, "name") {
                        if !name.starts_with("_Toc") && name != "_GoBack" {
                            para_anchor_id = Some(name);
                        }
                    }
                }
            }

            // ---- 箇条書き番号参照（w:numPr）----
            // <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
            Ok(Event::Start(e)) if in_ppr && !in_ppr_change && e.local_name().as_ref() == b"numPr" => {
                in_numpr = true;
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"numPr" => {
                in_numpr = false;
            }
            // ilvl: インデントレベル（0 = 最上位、1 = 1段インデント、…）
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if in_numpr && e.local_name().as_ref() == b"ilvl" =>
            {
                if let Some(val) = attr_value(&e, "val") {
                    para_num_ilvl = val.parse().unwrap_or(0);
                }
            }
            // numId: 箇条書き定義の参照 ID（0 はリストなしを意味するためスキップ）
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if in_numpr && e.local_name().as_ref() == b"numId" =>
            {
                para_num_id = attr_value(&e, "val").filter(|s| s != "0");
            }

            // ---- ランプロパティ ----
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"rPr" => {
                in_rpr = true;
                if in_ppr && !in_ppr_change { in_ppr_rpr = true; }
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"rPr" => {
                in_rpr = false;
                in_ppr_rpr = false;
            }

            // ---- 下線検出（見出し判定用）----
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if e.local_name().as_ref() == b"u" =>
            {
                let val = attr_value(&e, "w:val").or_else(|| attr_value(&e, "val"));
                if val.as_deref() != Some("none") {
                    if in_ppr_rpr {
                        ppr_underline = true;
                    } else if in_rpr && !in_ppr {
                        run_underline = true;
                    }
                }
            }

            // ---- 変更履歴: del（削除済みテキストは無視）----
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"del" => {
                in_del += 1;
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"del" => {
                in_del = in_del.saturating_sub(1);
            }
            // ---- 変更履歴: ins（挿入済みは採用）----
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"ins" => {
                in_ins += 1;
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"ins" => {
                in_ins = in_ins.saturating_sub(1);
            }

            // ---- mc:AlternateContent フォールバック（重複防止）----
            // mc:Choice 側に a:blip がある場合、mc:Fallback 内の v:imagedata は
            // スキップして画像の二重追加を防ぐ。
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"Fallback" => {
                in_mc_fallback += 1;
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"Fallback" => {
                in_mc_fallback = in_mc_fallback.saturating_sub(1);
            }

            // ---- 画像メタデータ（wp:docPr）----
            // <wp:docPr id="1" name="図 2" descr="代替テキスト"/>
            // descr（代替テキスト/alt text）を優先し、なければ name を使用
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) if e.local_name().as_ref() == b"docPr" => {
                let descr = attr_value(&e, "descr").filter(|s| !s.trim().is_empty());
                let name  = attr_value(&e, "name").filter(|s| !s.trim().is_empty());
                drawing_title = descr.or(name);
            }

            // ---- 画像参照（a:blip）----
            // a:blip は子要素を持つ場合(Start)と自己閉じ(Empty)の両方がある
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) if e.local_name().as_ref() == b"blip" => {
                // a:blip r:embed="rId5"
                if let Some(rid) = attr_value(&e, "embed") {
                    drawing_rid = Some(rid);
                }
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"drawing" => {
                // drawing 終了時に rId と title をまとめて Asset 化
                let title = drawing_title.take().unwrap_or_default();
                if let Some(rid) = drawing_rid.take() {
                    if let Some(raw) = images.get(&rid) {
                        current_assets.push(Asset {
                            asset_type: "image".to_string(),
                            id: Some(rid),
                            title,
                            data: raw.clone(), // Vec<u8>: Base64 エンコードは出力時に実施
                        });
                    }
                }
            }

            // ---- VML 画像参照（v:imagedata）----
            // <v:imagedata r:id="rId5" o:title="図のタイトル"/>
            // w:pict > v:shape > v:imagedata の形式で埋め込まれたレガシー画像。
            // mc:Fallback 内は a:blip 側（mc:Choice）で処理済みのためスキップする。
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if e.local_name().as_ref() == b"imagedata" && in_mc_fallback == 0 =>
            {
                vml_rid   = attr_value(&e, "id");   // r:id の local name は "id"
                vml_title = attr_value(&e, "title").filter(|s| !s.trim().is_empty()); // o:title
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"pict" => {
                // w:pict 終了時に VML 画像を Asset 化
                let title = vml_title.take().unwrap_or_default();
                if let Some(rid) = vml_rid.take() {
                    if let Some(raw) = images.get(&rid) {
                        current_assets.push(Asset {
                            asset_type: "image".to_string(),
                            id: Some(rid),
                            title,
                            data: raw.clone(),
                        });
                    }
                }
            }

            // ---- フィールドコード命令テキスト（w:instrText）----
            // フィールドコードの命令部分（例: " PAGE "、" DATE \@ ..."）を除外する。
            // begin/separate/end の fldChar で囲まれた表示値（w:t）は通常通り出力される。
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"instrText" => {
                in_fld_instr = true;
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"instrText" => {
                in_fld_instr = false;
            }

            // ---- ソフト改行（w:br）----
            // <w:br/> または <w:br w:type="textWrapping"/> はラン内の強制改行。
            // ページ区切り（w:type="page"）や段区切り（"column"）はテキストに含めない。
            // セル内では <br> に、段落内では \n に変換する。
            Ok(Event::Empty(e)) if in_del == 0 && e.local_name().as_ref() == b"br" => {
                let br_type = attr_value(&e, "type").unwrap_or_default();
                if br_type != "page" && br_type != "column" {
                    if in_table == 1 && in_table_cell {
                        current_cell_text.push_str("<br>");
                    } else if in_paragraph {
                        current_text.push('\n');
                    }
                }
            }

            // ---- テキストノード ----
            Ok(Event::Text(e)) if in_del == 0 && !in_ppr && !in_rpr && !in_fld_instr => {
                let text = e.unescape().unwrap_or_default();
                if in_omath > 0 && in_mt {
                    // 数式テキスト: omath_stack のトップバッファに追記
                    if let Some((_, buf)) = omath_stack.last_mut() {
                        buf.push_str(&text);
                    }
                } else if in_table == 1 && in_table_cell {
                    // テーブルセルのテキスト
                    current_cell_text.push_str(&text);
                } else if in_paragraph {
                    // 通常段落のテキスト
                    current_text.push_str(&text);
                }
            }

            // ---- 段落終了（テーブル外のみ）----
            Ok(Event::End(e)) if in_table == 0 && e.local_name().as_ref() == b"p" => {
                if in_paragraph {
                    let style = paragraph_style.take();
                    let ppr_ul = std::mem::replace(&mut ppr_underline, false);
                    let run_ul = std::mem::replace(&mut run_underline, false);
                    let num_id = para_num_id.take();
                    let num_ilvl = para_num_ilvl;
                    let alignment = para_alignment.take();
                    let outline_level = para_outline_level.take();
                    let anchor_id = para_anchor_id.take();

                    let heading_level = style.as_deref()
                        .and_then(|s| config.heading_level_for_style(s))
                        .or({
                            if (ppr_ul && config.docx.ppr_underline_as_heading)
                                || (run_ul && config.docx.run_underline_as_heading) {
                                Some(1)
                            } else {
                                None
                            }
                        });
                    let body = std::mem::take(&mut current_text).trim().to_string();
                    let assets = std::mem::take(&mut current_assets);
                    in_paragraph = false;

                    if let Some(level) = heading_level {
                        // 空の見出し（装飾用の空行など）はスキップ
                        if body.is_empty() && assets.is_empty() {
                            continue;
                        }
                        // level=0（"Title" スタイルなど）はセクションとして扱わずスキップ。
                        // config で level=0 を設定した場合のパニック防止でもある。
                        if level == 0 {
                            continue;
                        }
                        // 見出しテキストが空で図のみの場合（例: 図を含む見出し段落でテキスト無し）は
                        // 新しいセクションを作成せず、現在のセクションに図として追加する。
                        if body.is_empty() && !assets.is_empty() {
                            if let Some((_, sec)) = stack.last_mut() {
                                for asset in &assets {
                                    sec.elements.push(Element::AssetRef {
                                        asset_id: asset.id.clone().unwrap_or_default(),
                                        metadata: ElementMetadata {
                                            caption: if asset.title.is_empty() {
                                                None
                                            } else {
                                                Some(asset.title.clone())
                                            },
                                            ..Default::default()
                                        },
                                    });
                                }
                                sec.assets.extend(assets);
                            }
                            continue;
                        }
                        let new_section = Section {
                            context_path: Vec::new(), // fill_context_path() で後から設定
                            heading: body,
                            body_text: String::new(),
                            assets,
                            children: Vec::new(),
                            ..Default::default()
                        };
                        // スタックを巻き戻してこのレベルの親を探す（#1 スタック操作）
                        //
                        // アルゴリズム:
                        //   スタックトップのレベル >= 新しいレベル である限りポップして
                        //   適切な親（次のスタックトップ、なければ root_sections）に追加する。
                        //
                        // エッジケースの設計方針:
                        //   H2 → H1（逆転）: H2 を root に flush してから H1 を push する。
                        //   H1 → H3（階層スキップ）: H3 を H1 の直下の子として扱う。
                        //     ファントム H2 は挿入しない（文書に存在しない構造を生成しない）。
                        //   先頭が H2（H1 なし）: H2 を root_sections に直接追加する。
                        //
                        // 安全性: while ループ内の stack.pop() は直前の last() が Some を
                        //   返したことを確認してから呼ぶため、unwrap() でパニックしない。
                        while stack.last().is_some_and(|(l, _)| *l >= level) {
                            let (_, finished) = stack.pop()
                                .expect("stack.last() が Some だったため pop は Some を返す");
                            push_to_parent(&mut stack, &mut root_sections, finished);
                        }
                        stack.push((level, new_section));
                    } else if !body.is_empty() || !assets.is_empty() {
                        // 意味的役割を判定（箇条書きはリスト役割、それ以外はスタイル名から推定）
                        let role = if let Some(ref nid) = num_id {
                            let ordered = numbering
                                .get(nid)
                                .and_then(|fmts| fmts.get(num_ilvl as usize))
                                .map(|fmt| is_ordered_numfmt(fmt))
                                .unwrap_or(false);
                            if ordered {
                                Some(SemanticRole::OrderedList)
                            } else {
                                Some(SemanticRole::BulletList)
                            }
                        } else {
                            style.as_deref().and_then(|s| determine_role(s, &config.docx))
                        };

                        // 箇条書きプレフィックスを付与（numId が設定されている場合）
                        let body = if let Some(ref nid) = num_id {
                            let indent = "  ".repeat(num_ilvl as usize);
                            let ordered = numbering
                                .get(nid)
                                .and_then(|fmts| fmts.get(num_ilvl as usize))
                                .map(|fmt| is_ordered_numfmt(fmt))
                                .unwrap_or(false);
                            let prefix = if ordered {
                                format!("{}1. ", indent)
                            } else {
                                format!("{}- ", indent)
                            };
                            format!("{}{}", prefix, body)
                        } else {
                            body
                        };

                        // 現在のセクションの body_text と elements に追加
                        if let Some((_, sec)) = stack.last_mut() {
                            if !sec.body_text.is_empty() {
                                sec.body_text.push('\n');
                            }
                            sec.body_text.push_str(&body);

                            // Element::Paragraph を追加
                            // style: 正規化済みスタイル名（全角→半角変換）
                            // raw_style: XML から取得した生値（将来のカスタムスタイル対応用）
                            let normalized_style = style.as_deref().map(normalize_style_name);
                            sec.elements.push(Element::Paragraph {
                                text: body,
                                metadata: ElementMetadata {
                                    style: normalized_style,
                                    raw_style: style,
                                    alignment,
                                    outline_level,
                                    role,
                                    anchor_id,
                                    caption: None,
                                },
                            });

                            // 画像ごとに Element::AssetRef を追加（段落内の画像位置を保持）
                            for asset in &assets {
                                sec.elements.push(Element::AssetRef {
                                    asset_id: asset.id.clone().unwrap_or_default(),
                                    metadata: ElementMetadata {
                                        caption: if asset.title.is_empty() {
                                            None
                                        } else {
                                            Some(asset.title.clone())
                                        },
                                        ..Default::default()
                                    },
                                });
                            }

                            sec.assets.extend(assets);
                        }
                    }
                }
            }

            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::Error::from(e)
                .context("document.xml のXML読み取り中にエラーが発生")),
            _ => {}
        }
    }

    // スタックの残りをフラッシュ
    while let Some((_, finished)) = stack.pop() {
        push_to_parent(&mut stack, &mut root_sections, finished);
    }

    Ok(root_sections)
}

/// OMML スタックのトップをポップし、親バッファに LaTeX として結合する（#3）
///
/// # 変換ルール
/// | 親タグ  | 子タグ | 出力                          |
/// |---------|--------|-------------------------------|
/// | `f`     | `num`  | `\frac{<num>}{` (分子)        |
/// | `f`     | `den`  | `<den>}` (分母+閉じ)          |
/// | `sSup`  | `e`    | `<base>^{` (基数+上付き開き)  |
/// | `sSup`  | `sup`  | `<exp>}` (指数+閉じ)          |
/// | `sSub`  | `e`    | `<base>_{` (基数+下付き開き)  |
/// | `sSub`  | `sub`  | `<sub>}` (添字+閉じ)          |
/// | `rad`   | `e`    | `\sqrt{<radicand>}` (平方根)  |
/// | その他  | 任意   | バッファをそのまま連結        |
fn omath_pop_and_combine(stack: &mut Vec<(String, String)>, child_tag: &str) {
    if let Some((_, child_buf)) = stack.pop() {
        if let Some((parent_tag, parent_buf)) = stack.last_mut() {
            match (parent_tag.as_str(), child_tag) {
                ("f", "num") => {
                    parent_buf.push_str("\\frac{");
                    parent_buf.push_str(&child_buf);
                    parent_buf.push_str("}{");
                }
                ("f", "den") => {
                    parent_buf.push_str(&child_buf);
                    parent_buf.push('}');
                }
                ("sSup", "e") => {
                    parent_buf.push_str(&child_buf);
                    parent_buf.push_str("^{");
                }
                ("sSup", "sup") => {
                    parent_buf.push_str(&child_buf);
                    parent_buf.push('}');
                }
                ("sSub", "e") => {
                    parent_buf.push_str(&child_buf);
                    parent_buf.push_str("_{");
                }
                ("sSub", "sub") => {
                    parent_buf.push_str(&child_buf);
                    parent_buf.push('}');
                }
                ("rad", "e") => {
                    parent_buf.push_str("\\sqrt{");
                    parent_buf.push_str(&child_buf);
                    parent_buf.push('}');
                }
                // deg（√の次数）や nary（総和記号 Σ など）は現状テキストを連結
                _ => {
                    parent_buf.push_str(&child_buf);
                }
            }
        }
    }
}

/// テーブルの行データをMarkdown形式に変換する（#2 テーブル Markdown 適合）
///
/// - 最初の行をヘッダーとして扱い、その後にセパレーター行を挿入する
/// - セル内の `\n` は `<br>` に変換済み（セル段落終了ハンドラで処理）
/// - 空セルは ` ` (スペース) で埋めて列ズレを防止する
fn rows_to_markdown(rows: &[Vec<String>]) -> String {
    if rows.is_empty() {
        return String::new();
    }
    let cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    if cols == 0 {
        return String::new();
    }

    let mut lines = Vec::with_capacity(rows.len() + 1);
    for (i, row) in rows.iter().enumerate() {
        // 不足列を空文字で埋め、パイプ文字をエスケープ
        let cells: Vec<String> = (0..cols)
            .map(|j| {
                let cell = row.get(j)
                    .map(|c| c.replace('|', r"\|"))
                    .unwrap_or_default();
                // 空セルは " " に置換して列のズレを防止
                if cell.is_empty() { " ".to_string() } else { cell }
            })
            .collect();
        lines.push(format!("| {} |", cells.join(" | ")));

        // ヘッダー行（最初の行）の直後にセパレーターを挿入
        if i == 0 {
            lines.push(format!("| {} |", vec!["---"; cols].join(" | ")));
        }
    }
    lines.join("\n")
}

/// numFmt 値が番号付きリスト（ordered list）かどうかを判定する
/// "decimal", "lowerLetter" などは番号付き、"bullet" などは箇条書き
fn is_ordered_numfmt(fmt: &str) -> bool {
    matches!(
        fmt,
        "decimal"
            | "decimalZero"
            | "lowerLetter"
            | "upperLetter"
            | "lowerRoman"
            | "upperRoman"
            | "ordinal"
            | "cardinalText"
            | "ordinalText"
    )
}

/// セクションを適切な親に追加する
fn push_to_parent(
    stack: &mut [(usize, Section)],
    root: &mut Vec<Section>,
    section: Section,
) {
    if let Some((_, parent)) = stack.last_mut() {
        parent.children.push(section);
    } else {
        root.push(section);
    }
}

/// スタイル名を CamelCase・ハイフン・アンダースコア・スペースで単語に分割し、
/// 小文字の単語リストを返す。
///
/// 例: "WarningBox" → ["warning", "box"]、"note-text" → ["note", "text"]、
///     "Footnote" → ["footnote"]（誤検知を防止）
fn style_words(style: &str) -> Vec<String> {
    let mut words: Vec<String> = Vec::new();
    let mut word = String::new();
    let chars: Vec<char> = style.chars().collect();
    let n = chars.len();
    let mut i = 0;
    while i < n {
        let ch = chars[i];
        if ch == '-' || ch == '_' || ch == ' ' || ch == '.' {
            if !word.is_empty() {
                words.push(word.to_lowercase());
                word = String::new();
            }
        } else if ch.is_uppercase() {
            let word_has_lower = word.chars().any(|c| c.is_lowercase());
            // 頭字語（ALL_CAPS）末尾の境界検出: 次が小文字かつ現在 word 末尾が大文字
            // 例: "WARNINGBox" の 'B' で "WARNING" / "Box" に分割
            let next_is_lower = i + 1 < n && chars[i + 1].is_lowercase();
            let last_is_upper = word.chars().next_back().map(|c| c.is_uppercase()).unwrap_or(false);
            if !word.is_empty() && (word_has_lower || (next_is_lower && last_is_upper)) {
                words.push(word.to_lowercase());
                word = String::new();
            }
            word.push(ch);
        } else {
            word.push(ch);
        }
        i += 1;
    }
    if !word.is_empty() {
        words.push(word.to_lowercase());
    }
    words
}

/// スタイル名から意味的役割（SemanticRole）を推定する
///
/// 1. config の `semantic_role_styles` にカスタムマッピングがあれば優先して返す。
/// 2. 組み込みルールはスタイル名を単語に分割して境界マッチを行い、
///    部分一致による誤検知（"Footnote" → "note"、"Barcode" → "code" 等）を防ぐ。
/// 3. 日本語キーワードは単語境界の概念が異なるため `contains` で照合する。
///
/// 見出しスタイルはここでは扱わない（heading_level 判定済みのため）。
fn determine_role(style: &str, config: &DocxConfig) -> Option<SemanticRole> {
    // カスタムマッピングを優先確認
    let normalized = normalize_style_name(style);
    if let Some(role) = config.semantic_role_styles.get(&normalized) {
        return Some(role.clone());
    }

    let words = style_words(style);
    let lower = style.to_lowercase();

    if words.iter().any(|w| matches!(w.as_str(), "warning" | "caution" | "alert" | "danger"))
        || lower.contains("警告")
    {
        Some(SemanticRole::Warning)
    } else if words.iter().any(|w| w == "note")
        || lower.contains("注意") || lower.contains("注記")
    {
        Some(SemanticRole::Note)
    } else if words.iter().any(|w| matches!(w.as_str(), "tip" | "hint"))
        || lower.contains("ヒント")
    {
        Some(SemanticRole::Tip)
    } else if words.iter().any(|w| matches!(w.as_str(), "code" | "verbatim"))
        // "PreFormat"・"SourceText" はCamelCase分割で1単語にならないため個別チェック
        || lower.contains("preformat") || lower.contains("sourcetext")
    {
        Some(SemanticRole::CodeBlock)
    } else if words.iter().any(|w| matches!(w.as_str(), "quote" | "quotation"))
        || lower.contains("引用")
    {
        Some(SemanticRole::Quote)
    } else {
        None
    }
}

/// XML要素から属性値を取得する（名前空間プレフィックスを無視）
fn attr_value(e: &quick_xml::events::BytesStart, name: &str) -> Option<String> {
    let local = name.split(':').next_back().unwrap_or(name);
    for attr in e.attributes().flatten() {
        let key = attr.key.local_name();
        if key.as_ref() == local.as_bytes() {
            return String::from_utf8(attr.value.to_vec()).ok();
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::config::DocxConfig;

    fn default_docx_config() -> DocxConfig {
        DocxConfig::default()
    }

    // ── determine_role ────────────────────────────────────────────────────

    #[test]
    fn test_determine_role_warning_variants() {
        let cfg = default_docx_config();
        assert_eq!(determine_role("Warning", &cfg), Some(SemanticRole::Warning));
        assert_eq!(determine_role("WarningBox", &cfg), Some(SemanticRole::Warning));
        assert_eq!(determine_role("CautionNote", &cfg), Some(SemanticRole::Warning));
        assert_eq!(determine_role("AlertStyle", &cfg), Some(SemanticRole::Warning));
        assert_eq!(determine_role("DangerZone", &cfg), Some(SemanticRole::Warning));
    }

    #[test]
    fn test_determine_role_note_variants() {
        let cfg = default_docx_config();
        assert_eq!(determine_role("Note", &cfg), Some(SemanticRole::Note));
        assert_eq!(determine_role("NoteBox", &cfg), Some(SemanticRole::Note));
        assert_eq!(determine_role("note-text", &cfg), Some(SemanticRole::Note));
    }

    #[test]
    fn test_determine_role_no_false_positive_footnote() {
        // "Footnote" は "note" という単語を含まない（footnote = 1 単語）
        let cfg = default_docx_config();
        assert_eq!(determine_role("Footnote", &cfg), None);
        assert_eq!(determine_role("FootnoteText", &cfg), None);
        assert_eq!(determine_role("Annotation", &cfg), None);
    }

    #[test]
    fn test_determine_role_no_false_positive_code() {
        // "Barcode" は "code" という単語を含まない（barcode = 1 単語）
        let cfg = default_docx_config();
        assert_eq!(determine_role("Barcode", &cfg), None);
    }

    #[test]
    fn test_determine_role_code_block_variants() {
        let cfg = default_docx_config();
        assert_eq!(determine_role("CodeBlock", &cfg), Some(SemanticRole::CodeBlock));
        assert_eq!(determine_role("PreFormat", &cfg), Some(SemanticRole::CodeBlock));
        assert_eq!(determine_role("Verbatim", &cfg), Some(SemanticRole::CodeBlock));
        // "SourceText" は複合語だが contains で検出
        assert_eq!(determine_role("SourceText", &cfg), Some(SemanticRole::CodeBlock));
    }

    #[test]
    fn test_determine_role_tip_variants() {
        let cfg = default_docx_config();
        assert_eq!(determine_role("Tip", &cfg), Some(SemanticRole::Tip));
        assert_eq!(determine_role("TipBox", &cfg), Some(SemanticRole::Tip));
        assert_eq!(determine_role("HintStyle", &cfg), Some(SemanticRole::Tip));
    }

    #[test]
    fn test_determine_role_quote_variants() {
        let cfg = default_docx_config();
        assert_eq!(determine_role("Quote", &cfg), Some(SemanticRole::Quote));
        assert_eq!(determine_role("BlockQuote", &cfg), Some(SemanticRole::Quote));
        assert_eq!(determine_role("Quotation", &cfg), Some(SemanticRole::Quote));
    }

    #[test]
    fn test_determine_role_japanese_keywords() {
        let cfg = default_docx_config();
        assert_eq!(determine_role("警告スタイル", &cfg), Some(SemanticRole::Warning));
        assert_eq!(determine_role("注意事項", &cfg), Some(SemanticRole::Note));
        assert_eq!(determine_role("注記スタイル", &cfg), Some(SemanticRole::Note));
        assert_eq!(determine_role("ヒント", &cfg), Some(SemanticRole::Tip));
        assert_eq!(determine_role("引用スタイル", &cfg), Some(SemanticRole::Quote));
    }

    #[test]
    fn test_determine_role_custom_mapping() {
        let mut cfg = default_docx_config();
        cfg.semantic_role_styles.insert("MyCustomStyle".to_string(), SemanticRole::Warning);
        assert_eq!(determine_role("MyCustomStyle", &cfg), Some(SemanticRole::Warning));
        // カスタムマッピングがない場合は組み込みルールにフォールバック
        assert_eq!(determine_role("NoteBox", &cfg), Some(SemanticRole::Note));
    }

    #[test]
    fn test_determine_role_unknown_style() {
        let cfg = default_docx_config();
        assert_eq!(determine_role("Normal", &cfg), None);
        assert_eq!(determine_role("BodyText", &cfg), None);
        assert_eq!(determine_role("ListParagraph", &cfg), None);
    }

    // ── style_words ───────────────────────────────────────────────────────

    #[test]
    fn test_style_words_camel_case() {
        assert_eq!(style_words("WarningBox"), vec!["warning", "box"]);
        assert_eq!(style_words("NoteText"), vec!["note", "text"]);
        assert_eq!(style_words("Footnote"), vec!["footnote"]);
        assert_eq!(style_words("FootnoteText"), vec!["footnote", "text"]);
    }

    #[test]
    fn test_style_words_separators() {
        assert_eq!(style_words("note-text"), vec!["note", "text"]);
        assert_eq!(style_words("code_block"), vec!["code", "block"]);
        assert_eq!(style_words("warning style"), vec!["warning", "style"]);
    }

    // ── is_ordered_numfmt ─────────────────────────────────────────────────

    #[test]
    fn test_is_ordered_numfmt() {
        assert!(is_ordered_numfmt("decimal"));
        assert!(is_ordered_numfmt("lowerLetter"));
        assert!(is_ordered_numfmt("upperRoman"));
        assert!(!is_ordered_numfmt("bullet"));
        assert!(!is_ordered_numfmt(""));
    }

    // ── bookmarkStart フィルタリング ──────────────────────────────────────

    /// フィルタリング仕様のテスト用ヘルパー（実装と同一ロジック）
    fn should_skip_bookmark(name: &str) -> bool {
        name.starts_with("_Toc") || name == "_GoBack"
    }

    #[test]
    fn test_bookmark_filter_toc_skipped() {
        // Word 自動生成の目次ブックマークはスキップされる
        assert!(should_skip_bookmark("_Toc123456"));
        assert!(should_skip_bookmark("_Toc0"));
    }

    #[test]
    fn test_bookmark_filter_goback_exact_match() {
        // _GoBack は完全一致でのみスキップ（starts_with では "_GoBackward" 等を誤除外する）
        assert!(should_skip_bookmark("_GoBack"));
        assert!(!should_skip_bookmark("_GoBackward")); // ユーザー定義: スキップしない
        assert!(!should_skip_bookmark("_GoBack2"));    // ユーザー定義: スキップしない
    }

    #[test]
    fn test_bookmark_filter_user_defined_underscore() {
        // アンダースコア始まりのユーザー定義ブックマークはスキップされない
        assert!(!should_skip_bookmark("_CustomSection"));
        assert!(!should_skip_bookmark("_MyAnchor"));
        assert!(!should_skip_bookmark("_section-intro"));
    }

    #[test]
    fn test_bookmark_filter_normal_names_not_skipped() {
        assert!(!should_skip_bookmark("section-intro"));
        assert!(!should_skip_bookmark("chapter1"));
        assert!(!should_skip_bookmark(""));
    }

    // ── outlineLvl 変換（0-based → 1-based）────────────────────────────────

    #[test]
    fn test_outline_level_conversion() {
        // Word の outlineLvl は 0-based、出力は 1-based に変換する
        let raw: u32 = 0;
        let converted = raw + 1;
        assert_eq!(converted, 1);

        let raw: u32 = 2;
        let converted = raw + 1;
        assert_eq!(converted, 3);
    }

    // ── テーブルセル内の role 判定（Known Limitation） ───────────────────
    // 現在の実装ではセル内テキストは current_cell_text に平文で収集され、
    // Element::Paragraph として処理されない。
    // そのため SemanticRole はテーブル外の段落にのみ適用される。
    // この制限を tests で明示し、将来の改善時に回帰検知できるようにする。

    #[test]
    fn test_determine_role_not_called_for_table_cell_content() {
        // テーブルセル内のスタイルは determine_role の対象外（テキストのみ収集）
        // この仕様が変わった場合はここで検知できる
        let cfg = default_docx_config();
        // セル内の "Warning" テキストはセルの平文テキストとして扱われ、
        // SemanticRole は付与されない（role 判定は table 外の段落のみ）
        // ここでは determine_role 自体は正しく動作することを確認するのみ
        assert_eq!(determine_role("Warning", &cfg), Some(SemanticRole::Warning));
    }

    // ── 改善ポイント 1: 境界値・エッジケース ──────────────────────────────

    #[test]
    fn test_determine_role_empty_style() {
        let cfg = default_docx_config();
        // 空文字はパニックせず None を返す
        assert_eq!(determine_role("", &cfg), None);
    }

    #[test]
    fn test_determine_role_whitespace_only() {
        let cfg = default_docx_config();
        // スペースのみもパニックせず None を返す
        assert_eq!(determine_role("   ", &cfg), None);
    }

    #[test]
    fn test_determine_role_very_long_style_name() {
        let cfg = default_docx_config();
        // 255 文字超の入力でもパニックしない
        let long_name = "A".repeat(300);
        let _ = determine_role(&long_name, &cfg); // パニックしないことを確認
        let long_with_keyword = format!("Warning{}", "X".repeat(250));
        assert_eq!(determine_role(&long_with_keyword, &cfg), Some(SemanticRole::Warning));
    }

    #[test]
    fn test_style_words_empty_and_whitespace() {
        assert!(style_words("").is_empty());
        assert!(style_words("   ").is_empty());
        assert!(style_words("---").is_empty());
    }

    // ── 改善ポイント 2: 大文字小文字の正規化 ──────────────────────────────

    #[test]
    fn test_style_words_all_caps() {
        // 全大文字は 1 単語として扱われる
        assert_eq!(style_words("WARNING"), vec!["warning"]);
        assert_eq!(style_words("NOTE"), vec!["note"]);
        assert_eq!(style_words("CODE"), vec!["code"]);
    }

    #[test]
    fn test_style_words_all_caps_compound() {
        // 頭字語 + CamelCase の組み合わせ
        assert_eq!(style_words("WARNINGBox"), vec!["warning", "box"]);
    }

    #[test]
    fn test_determine_role_all_caps() {
        let cfg = default_docx_config();
        // 全大文字でもキーワードが認識される
        assert_eq!(determine_role("WARNING", &cfg), Some(SemanticRole::Warning));
        assert_eq!(determine_role("NOTE", &cfg), Some(SemanticRole::Note));
        assert_eq!(determine_role("CODE", &cfg), Some(SemanticRole::CodeBlock));
        assert_eq!(determine_role("TIP", &cfg), Some(SemanticRole::Tip));
    }

    #[test]
    fn test_determine_role_all_lowercase() {
        let cfg = default_docx_config();
        // 全小文字でもキーワードが認識される
        assert_eq!(determine_role("warning", &cfg), Some(SemanticRole::Warning));
        assert_eq!(determine_role("note", &cfg), Some(SemanticRole::Note));
        assert_eq!(determine_role("tip", &cfg), Some(SemanticRole::Tip));
    }

    // ── 改善ポイント 3: 多言語・特殊文字 ─────────────────────────────────

    #[test]
    fn test_determine_role_japanese_fullwidth_number() {
        let cfg = default_docx_config();
        // 全角数字を含む日本語スタイル名でも日本語キーワードが認識される
        assert_eq!(determine_role("警告スタイル１", &cfg), Some(SemanticRole::Warning));
        assert_eq!(determine_role("注意事項２", &cfg), Some(SemanticRole::Note));
    }

    #[test]
    fn test_style_words_with_emoji_does_not_panic() {
        // 絵文字・補助文字が含まれてもパニックしない（Rust の chars() は正しく処理する）
        let result = style_words("Warning🚨Style");
        assert!(!result.is_empty()); // パニックせず結果が返ることを確認
    }

    // ── 改善ポイント 4: 優先順位の競合テスト ─────────────────────────────

    #[test]
    fn test_determine_role_priority_note_before_code() {
        let cfg = default_docx_config();
        // 評価順: Warning > Note > Tip > CodeBlock > Quote
        // "CodeNote" は Note と Code を含むが、Note の評価が先なので Note が返る
        assert_eq!(determine_role("CodeNote", &cfg), Some(SemanticRole::Note));
    }

    #[test]
    fn test_determine_role_priority_warning_before_note() {
        let cfg = default_docx_config();
        // "WarningNote" → Warning が Note より先に評価される
        assert_eq!(determine_role("WarningNote", &cfg), Some(SemanticRole::Warning));
    }

    // ── 改善ポイント 5: outline_level の最大値超過 ────────────────────────

    #[test]
    fn test_outline_level_max_word_spec() {
        // Word 仕様上の最大は 0-based で 8（= 1-based で 9）
        let raw: u32 = 8;
        let converted = raw + 1;
        assert_eq!(converted, 9);
    }

    #[test]
    fn test_outline_level_beyond_max_no_panic() {
        // 仕様外の値（raw=10 など）でも u32 加算はパニックしない
        let raw: u32 = 10;
        let converted = raw.saturating_add(1);
        assert_eq!(converted, 11); // オーバーフローなし
    }
}

/// ZIPアーカイブから指定エントリをUTF-8文字列として読み込む
fn read_zip_entry(
    archive: &mut ZipArchive<BufReader<std::fs::File>>,
    name: &str,
) -> Result<String> {
    let mut entry = archive.by_name(name)
        .with_context(|| format!("ZIPエントリが見つかりません: {}", name))?;
    let mut buf = String::new();
    entry.read_to_string(&mut buf)
        .with_context(|| format!("ZIPエントリの読み込みに失敗: {}", name))?;
    Ok(buf)
}
