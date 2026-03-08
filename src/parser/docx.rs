use std::collections::HashMap;
use std::io::{BufReader, Read};
use std::path::Path;

use base64::Engine;
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use zip::ZipArchive;

use crate::config::Config;
use crate::models::{Asset, Document, Section};

type Error = Box<dyn std::error::Error + Send + Sync>;

/// DOCXファイルを解析してDocumentを返す
pub fn parse(path: &Path, config: &Config) -> Result<Document, Error> {
    let file = std::fs::File::open(path)?;
    let mut archive = ZipArchive::new(BufReader::new(file))?;

    // ファイル名をタイトルとして使用
    let title = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("Untitled")
        .to_string();

    // relationships を読み込んで画像パスを解決
    let rels = parse_rels(&mut archive)?;

    // 画像データをBase64で取得
    let images = extract_images(&mut archive, &rels)?;

    // document.xml をパース
    let xml = read_zip_entry(&mut archive, "word/document.xml")?;
    let sections = parse_document_xml(&xml, &images, config)?;

    Ok(Document { title, sections })
}

/// word/_rels/document.xml.rels を解析し、rId -> パスのマップを返す
fn parse_rels(archive: &mut ZipArchive<BufReader<std::fs::File>>) -> Result<HashMap<String, String>, Error> {
    let xml = match read_zip_entry(archive, "word/_rels/document.xml.rels") {
        Ok(s) => s,
        Err(_) => return Ok(HashMap::new()),
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
            Err(e) => return Err(e.into()),
            _ => {}
        }
    }
    Ok(rels)
}

/// ZIP内の画像ファイルをBase64エンコードして返す
fn extract_images(
    archive: &mut ZipArchive<BufReader<std::fs::File>>,
    rels: &HashMap<String, String>,
) -> Result<HashMap<String, String>, Error> {
    let mut images = HashMap::new();
    for (rid, zip_path) in rels {
        if let Ok(mut entry) = archive.by_name(zip_path) {
            let mut buf = Vec::new();
            entry.read_to_end(&mut buf)?;
            let encoded = base64::engine::general_purpose::STANDARD.encode(&buf);
            images.insert(rid.clone(), encoded);
        }
    }
    Ok(images)
}

/// document.xml を走査しセクションツリーを構築する
fn parse_document_xml(xml: &str, images: &HashMap<String, String>, config: &Config) -> Result<Vec<Section>, Error> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    // パーサー状態
    let mut in_del = 0u32;           // w:del ネスト深さ
    let mut in_ins = 0u32;           // w:ins ネスト深さ（未使用だがバランス追跡用）
    let mut in_ppr = false;          // w:pPr 内か
    let mut in_ppr_rpr = false;      // w:pPr > w:rPr 内か（段落デフォルト書式）
    let mut in_rpr = false;          // w:rPr 内か（ラン書式）
    let mut ppr_underline = false;   // w:pPr > w:rPr に w:u があるか（見出し判定用）
    let mut run_underline = false;   // ラン w:rPr に w:u があるか（見出し判定用）
    let mut current_text = String::new();           // 現在の段落テキスト
    let mut current_assets: Vec<Asset> = Vec::new(); // 現在の段落の画像
    let mut drawing_rid: Option<String> = None;     // 処理中の drawing の rId

    // セクションスタック: (level, section)
    // level 0 = ルート（Heading1）、level 1 = Heading2、...
    // ルートの "children" が最終出力
    let mut stack: Vec<(usize, Section)> = Vec::new();
    // まだどのセクションにも属さない段落を受け取るルートセクション
    let mut root_sections: Vec<Section> = Vec::new();

    let mut in_paragraph = false;
    let mut paragraph_style: Option<String> = None;

    loop {
        match reader.read_event() {
            // ---- 段落開始 ----
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"p" => {
                in_paragraph = true;
                paragraph_style = None;
                ppr_underline = false;
                run_underline = false;
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
            }
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) if in_ppr && e.local_name().as_ref() == b"pStyle" => {
                if let Some(val) = attr_value(&e, "w:val").or_else(|| attr_value(&e, "val")) {
                    paragraph_style = Some(val);
                }
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

            // ---- 下線検出 ----
            // Word で「見出し」スタイルを使わず下線で見出しを表現するケース
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if e.local_name().as_ref() == b"u" =>
            {
                let val = attr_value(&e, "w:val").or_else(|| attr_value(&e, "val"));
                // w:val="none" は下線なし
                if val.as_deref() != Some("none") {
                    if in_ppr_rpr {
                        ppr_underline = true;  // 段落デフォルト書式（pPr > rPr）
                    } else if in_rpr && !in_ppr {
                        run_underline = true;  // ラン書式（r > rPr）
                    }
                }
            }

            // ---- 変更履歴: del（削除済みテキストは無視）----
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"del" => {
                in_del += 1;
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"del" => {
                if in_del > 0 { in_del -= 1; }
            }
            // ---- 変更履歴: ins（挿入済みは採用）----
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"ins" => {
                in_ins += 1;
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"ins" => {
                if in_ins > 0 { in_ins -= 1; }
            }

            // ---- 画像参照 ----
            Ok(Event::Empty(e)) if e.local_name().as_ref() == b"blip" => {
                // a:blip r:embed="rId5"
                if let Some(rid) = attr_value(&e, "r:embed").or_else(|| attr_value(&e, "embed")) {
                    drawing_rid = Some(rid);
                }
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"drawing" => {
                if let Some(rid) = drawing_rid.take() {
                    if let Some(b64) = images.get(&rid) {
                        current_assets.push(Asset {
                            asset_type: "image".to_string(),
                            title: String::new(),
                            data: b64.clone(),
                        });
                    }
                }
            }

            // ---- テキストノード ----
            Ok(Event::Text(e)) if in_paragraph && in_del == 0 && !in_ppr && !in_rpr => {
                let text = e.unescape().unwrap_or_default();
                current_text.push_str(&text);
            }

            // ---- 段落終了 ----
            Ok(Event::End(e)) if e.local_name().as_ref() == b"p" => {
                if in_paragraph {
                    let style = paragraph_style.take();
                    let ppr_ul = std::mem::replace(&mut ppr_underline, false);
                    let run_ul = std::mem::replace(&mut run_underline, false);
                    // pStyleによる検出を優先し、なければ下線フォールバック（設定で有効な場合）
                    let heading_level = style.as_deref()
                        .and_then(|s| config.heading_level_for_style(s))
                        .or_else(|| {
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
                        // 空の見出しテキスト（装飾用の空行など）はスキップ
                        if body.is_empty() && assets.is_empty() {
                            in_paragraph = false;
                            continue;
                        }
                        // 新しいセクションを開始
                        let new_section = Section {
                            heading: body,
                            body_text: String::new(),
                            assets,
                            children: Vec::new(),
                        };

                        // スタックを巻き戻してこのレベルの親を探す
                        while stack.last().map_or(false, |(l, _)| *l >= level) {
                            let (_, finished) = stack.pop().unwrap();
                            push_to_parent(&mut stack, &mut root_sections, finished);
                        }
                        stack.push((level, new_section));
                    } else if !body.is_empty() || !assets.is_empty() {
                        // 通常段落: 現在のセクションの body_text に追加
                        if let Some((_, sec)) = stack.last_mut() {
                            if !sec.body_text.is_empty() {
                                sec.body_text.push('\n');
                            }
                            sec.body_text.push_str(&body);
                            sec.assets.extend(assets);
                        } else {
                            // セクション外の段落は無視（またはルートに追加）
                        }
                    }
                }
            }

            Ok(Event::Eof) => break,
            Err(e) => return Err(e.into()),
            _ => {}
        }
    }

    // スタックの残りをフラッシュ
    while let Some((_, finished)) = stack.pop() {
        push_to_parent(&mut stack, &mut root_sections, finished);
    }

    Ok(root_sections)
}

/// セクションを適切な親に追加する
fn push_to_parent(
    stack: &mut Vec<(usize, Section)>,
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
    // name は "w:val" や "val" のどちらでも可
    let local = name.split(':').last().unwrap_or(name);
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
) -> Result<String, Error> {
    let mut entry = archive.by_name(name)?;
    let mut buf = String::new();
    entry.read_to_string(&mut buf)?;
    Ok(buf)
}
