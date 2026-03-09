use std::collections::HashMap;
use std::io::{BufReader, Read};
use std::path::Path;

use anyhow::{Context, Result};
use image::codecs::jpeg::JpegEncoder;
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use zip::ZipArchive;

use crate::config::Config;
use crate::models::{Asset, Document, Section};

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
/// config.image_max_px > 0 の場合は長辺をリサイズして JPEG 再エンコードする。
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

            // リサイズ・圧縮が有効な場合は画像を処理する
            let final_buf = if config.image_max_px > 0 {
                resize_and_compress(&buf, config.image_max_px, config.image_quality)
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
    let mut in_paragraph = false;
    let mut paragraph_style: Option<String> = None;

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
                    let md = rows_to_markdown(&table_rows);
                    if let Some((_, sec)) = stack.last_mut() {
                        if !sec.body_text.is_empty() {
                            sec.body_text.push('\n');
                        }
                        sec.body_text.push_str(&md);
                    }
                    table_rows.clear();
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
            }

            // ---- 段落プロパティ ----
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"pPr" => {
                in_ppr = true;
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"pPr" => {
                in_ppr = false;
                in_ppr_rpr = false;
                in_numpr = false;
            }
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) if in_ppr && e.local_name().as_ref() == b"pStyle" => {
                if let Some(val) = attr_value(&e, "w:val").or_else(|| attr_value(&e, "val")) {
                    paragraph_style = Some(val);
                }
            }

            // ---- 箇条書き番号参照（w:numPr）----
            // <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
            Ok(Event::Start(e)) if in_ppr && e.local_name().as_ref() == b"numPr" => {
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
                if in_ppr { in_ppr_rpr = true; }
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
                            title,
                            data: raw.clone(), // Vec<u8>: Base64 エンコードは出力時に実施
                        });
                    }
                }
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
            Ok(Event::Text(e)) if in_del == 0 && !in_ppr && !in_rpr => {
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

                    let heading_level = style.as_deref()
                        .and_then(|s| config.heading_level_for_style(s))
                        .or({
                            if (ppr_ul && config.ppr_underline_as_heading)
                                || (run_ul && config.run_underline_as_heading) {
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

                        // 現在のセクションの body_text に追加
                        if let Some((_, sec)) = stack.last_mut() {
                            if !sec.body_text.is_empty() {
                                sec.body_text.push('\n');
                            }
                            sec.body_text.push_str(&body);
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
