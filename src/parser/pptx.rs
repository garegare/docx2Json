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

/// PPTXファイルを解析してDocumentを返す
///
/// 各スライドを1つのSectionとして生成する。
/// - heading: スライドタイトル（title/ctrTitle プレースホルダー）
/// - body_text: テキストボックスを Y 座標順に結合したテキスト
/// - assets: スライド内の画像
/// - notes があれば body_text 末尾に "[ノート]" として付加
pub fn parse(path: &Path, config: &Config) -> Result<Document> {
    let file = std::fs::File::open(path)
        .with_context(|| format!("ファイルを開けません: {}", path.display()))?;
    let mut archive = ZipArchive::new(BufReader::new(file))
        .context("ZIPアーカイブとして開けません（破損または非PPTXファイルの可能性）")?;

    // プレゼンテーションのタイトルを取得（core.xml → ファイル名の順でフォールバック）
    let title = read_doc_title(&mut archive).unwrap_or_else(|_| {
        path.file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("Untitled")
            .to_string()
    });

    // スライドのZIPパスを順序付きで取得
    let slide_paths =
        read_slide_order(&mut archive).context("スライド順序の取得に失敗")?;

    // notesSlide のマップを構築: slide ZIP パス → notes テキスト
    let notes_map = build_notes_map(&mut archive).unwrap_or_default();

    // 各スライドをセクションとしてパース
    let mut sections = Vec::new();
    for (idx, slide_path) in slide_paths.iter().enumerate() {
        match parse_slide(&mut archive, slide_path, idx + 1, config, &notes_map) {
            Ok(section) => sections.push(section),
            Err(e) => eprintln!("  警告: スライド {} のパースに失敗: {:#}", idx + 1, e),
        }
    }

    Ok(Document { title, sections })
}

// ─── スライド順序取得 ──────────────────────────────────────────────────────────

/// docProps/core.xml の dc:title を返す
fn read_doc_title(archive: &mut ZipArchive<BufReader<std::fs::File>>) -> Result<String> {
    let xml = read_zip_entry(archive, "docProps/core.xml")?;
    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    let mut in_title = false;
    let mut title = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"title" => in_title = true,
            Ok(Event::End(e)) if e.local_name().as_ref() == b"title" => in_title = false,
            Ok(Event::Text(e)) if in_title => {
                title.push_str(e.unescape().unwrap_or_default().as_ref())
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::Error::from(e)),
            _ => {}
        }
    }

    if title.is_empty() {
        anyhow::bail!("タイトルなし");
    }
    Ok(title)
}

/// presentation.xml と presentation.xml.rels からスライドの ZIP パスを順序付きで返す
fn read_slide_order(
    archive: &mut ZipArchive<BufReader<std::fs::File>>,
) -> Result<Vec<String>> {
    // rId → Target マップを構築（type = ".../slide" のみ）
    let rels_xml = read_zip_entry(archive, "ppt/_rels/presentation.xml.rels")?;
    let rid_to_target = parse_rels_by_type(&rels_xml, "/slide");

    // presentation.xml の sldIdLst から rId の順序を取得
    let pres_xml = read_zip_entry(archive, "ppt/presentation.xml")?;
    let mut reader = Reader::from_str(&pres_xml);
    reader.config_mut().trim_text(true);

    let mut in_sldidlst = false;
    let mut ordered_rids: Vec<String> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"sldIdLst" => in_sldidlst = true,
                b"sldId" if in_sldidlst => {
                    // r:id 属性（ローカル名 "id" で取得後、同名属性の衝突を避けるため
                    // "r:id" を指定して名前空間プレフィックスも考慮して検索）
                    if let Some(rid) = attr_value_with_prefix(&e, "r", "id") {
                        ordered_rids.push(rid);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) if e.local_name().as_ref() == b"sldId" && in_sldidlst => {
                if let Some(rid) = attr_value_with_prefix(&e, "r", "id") {
                    ordered_rids.push(rid);
                }
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"sldIdLst" => {
                in_sldidlst = false;
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::Error::from(e)),
            _ => {}
        }
    }

    // rId をパスに変換し "ppt/" プレフィックスを付ける
    let paths: Vec<String> = ordered_rids
        .iter()
        .filter_map(|rid| rid_to_target.get(rid))
        .map(|target| format!("ppt/{}", target))
        .collect();

    Ok(paths)
}

// ─── ノートマップ ───────────────────────────────────────────────────────────────

/// notesSlide ↔ slide の対応を調べ、slide ZIP パス → notes テキストのマップを返す
fn build_notes_map(
    archive: &mut ZipArchive<BufReader<std::fs::File>>,
) -> Result<HashMap<String, String>> {
    let mut map = HashMap::new();

    // ppt/notesSlides/_rels/*.rels を列挙
    let rels_paths: Vec<String> = archive
        .file_names()
        .filter(|n| n.starts_with("ppt/notesSlides/_rels/") && n.ends_with(".rels"))
        .map(|s| s.to_string())
        .collect();

    for rels_path in rels_paths {
        // "ppt/notesSlides/_rels/notesSlide5.xml.rels" → "ppt/notesSlides/notesSlide5.xml"
        let notes_xml_path = match rels_path
            .strip_prefix("ppt/notesSlides/_rels/")
            .and_then(|s| s.strip_suffix(".rels"))
        {
            Some(name) => format!("ppt/notesSlides/{}", name),
            None => continue,
        };

        // rels からスライドへのパスを取得
        let rels_xml = match read_zip_entry(archive, &rels_path) {
            Ok(s) => s,
            Err(_) => continue,
        };
        let slide_targets = parse_rels_by_type(&rels_xml, "/slide");
        let slide_target = match slide_targets.values().next() {
            Some(t) => t.clone(),
            None => continue,
        };

        // "../slides/slide79.xml" → "ppt/slides/slide79.xml"
        let slide_zip_path = resolve_ppt_relative_path("ppt/notesSlides", &slide_target);

        // notes XML をパースしてボディテキストを取得
        let notes_xml = match read_zip_entry(archive, &notes_xml_path) {
            Ok(s) => s,
            Err(_) => continue,
        };
        let shapes = parse_shapes(&notes_xml);
        let notes_text: String = shapes
            .iter()
            .filter(|s| s.ph_type.as_deref() == Some("body"))
            .map(|s| s.text.as_str())
            .filter(|t| !t.is_empty())
            .collect::<Vec<_>>()
            .join("\n");

        if !notes_text.is_empty() {
            map.insert(slide_zip_path, notes_text);
        }
    }

    Ok(map)
}

// ─── スライドパース ───────────────────────────────────────────────────────────

/// 1枚のスライドをパースして Section を返す
fn parse_slide(
    archive: &mut ZipArchive<BufReader<std::fs::File>>,
    slide_path: &str,
    slide_num: usize,
    config: &Config,
    notes_map: &HashMap<String, String>,
) -> Result<Section> {
    // スライドの rels からイメージ rId → Target マップを取得
    let slide_file = slide_path.rsplit('/').next().unwrap_or("");
    let rels_path = format!("ppt/slides/_rels/{}.rels", slide_file);
    let image_rids: HashMap<String, String> = read_zip_entry(archive, &rels_path)
        .map(|xml| parse_rels_by_type(&xml, "/image"))
        .unwrap_or_default();

    // スライド XML をパースしてシェイプ一覧を取得
    let slide_xml = read_zip_entry(archive, slide_path)
        .with_context(|| format!("スライドXMLの読み込みに失敗: {}", slide_path))?;
    let shapes = parse_shapes(&slide_xml);

    // タイトルシェイプを特定（title / ctrTitle プレースホルダー）
    let heading = shapes
        .iter()
        .find(|s| matches!(s.ph_type.as_deref(), Some("title") | Some("ctrTitle")))
        .map(|s| s.text.clone())
        .filter(|t| !t.is_empty())
        .unwrap_or_else(|| format!("スライド {}", slide_num));

    // ボディシェイプ: メタ系プレースホルダーを除いて Y 座標順にソート
    let mut body_shapes: Vec<&Shape> = shapes
        .iter()
        .filter(|s| {
            !matches!(
                s.ph_type.as_deref(),
                Some("title") | Some("ctrTitle") | Some("sldNum") | Some("dt") | Some("ftr")
            )
        })
        .collect();
    body_shapes.sort_by_key(|s| s.y_offset);

    let mut body_parts: Vec<String> = body_shapes
        .iter()
        .map(|s| s.text.clone())
        .filter(|t| !t.is_empty())
        .collect();

    // スライドノートを補足として末尾に付加
    if let Some(notes) = notes_map.get(slide_path) {
        if !notes.is_empty() {
            body_parts.push(format!("[ノート]\n{}", notes));
        }
    }

    let body_text = body_parts.join("\n\n");

    // 画像アセットを収集
    let assets = collect_assets(archive, &shapes, &image_rids, "ppt/slides", config);

    Ok(Section {
        heading,
        body_text,
        assets,
        ..Default::default()
    })
}

/// スライド内の画像シェイプからアセットを収集する
fn collect_assets(
    archive: &mut ZipArchive<BufReader<std::fs::File>>,
    shapes: &[Shape],
    image_rids: &HashMap<String, String>,
    base_dir: &str,
    config: &Config,
) -> Vec<Asset> {
    let mut assets = Vec::new();
    for shape in shapes {
        let rid = match &shape.image_rid {
            Some(r) => r,
            None => continue,
        };
        let target = match image_rids.get(rid) {
            Some(t) => t,
            None => continue,
        };
        let zip_path = resolve_ppt_relative_path(base_dir, target);
        let buf = match read_zip_entry_bytes(archive, &zip_path) {
            Ok(b) => b,
            Err(_) => continue,
        };
        let data = if config.image_max_px > 0 {
            resize_and_compress(&buf, config.image_max_px, config.image_quality)
                .unwrap_or(buf)
        } else {
            buf
        };
        let title = zip_path.rsplit('/').next().unwrap_or("image").to_string();
        assets.push(Asset {
            asset_type: "image".to_string(),
            title,
            data,
        });
    }
    assets
}

// ─── シェイプパーサー ──────────────────────────────────────────────────────────

/// スライドXML（またはノートXML）内のシェイプ情報
struct Shape {
    /// プレースホルダータイプ（"title", "ctrTitle", "body", "sldNum" など）
    ph_type: Option<String>,
    /// スライド上のY座標（EMU単位）。テキストボックスのソートに使用
    y_offset: i64,
    /// シェイプ内の全テキスト（改行で段落を区切る）
    text: String,
    /// 画像シェイプの場合の rId（p:pic 要素の a:blip r:embed 属性）
    image_rid: Option<String>,
}

/// スライド/ノートXMLをパースしてシェイプ一覧を返す
///
/// SAX ストリームで走査し、p:sp（テキストシェイプ）と p:pic（画像シェイプ）を収集する。
fn parse_shapes(xml: &str) -> Vec<Shape> {
    let mut shapes: Vec<Shape> = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    // ─ p:sp / p:pic のネスト深度（グループシェイプ内でも動作するよう深度管理）
    let mut sp_depth: i32 = 0;
    let mut pic_depth: i32 = 0;

    // ─ 各シェイプの作業用状態
    let mut in_nvpr = false;     // inside p:nvPr
    let mut in_sppr = false;     // inside p:spPr
    let mut in_xfrm = false;     // inside a:xfrm (in p:spPr)
    let mut in_blipfill = false; // inside p:blipFill (in p:pic)
    let mut in_txbody = false;   // inside p:txBody
    let mut in_para = false;     // inside a:p
    let mut in_text = false;     // inside a:t

    let mut ph_type: Option<String> = None;
    let mut y_offset: i64 = 0;
    let mut image_rid: Option<String> = None;
    let mut shape_text = String::new();
    let mut para_level: u8 = 0;
    let mut para_text = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                match e.local_name().as_ref() {
                    b"sp" if sp_depth == 0 && pic_depth == 0 => {
                        sp_depth = 1;
                        ph_type = None;
                        y_offset = 0;
                        image_rid = None;
                        shape_text.clear();
                    }
                    b"sp" if sp_depth > 0 => sp_depth += 1,
                    b"pic" if sp_depth == 0 && pic_depth == 0 => {
                        pic_depth = 1;
                        ph_type = None;
                        y_offset = 0;
                        image_rid = None;
                        shape_text.clear();
                    }
                    b"pic" if pic_depth > 0 => pic_depth += 1,
                    b"nvPr" if sp_depth == 1 => in_nvpr = true,
                    b"ph" if in_nvpr => {
                        ph_type = attr_value(e, "type");
                    }
                    b"spPr" if sp_depth == 1 || pic_depth == 1 => in_sppr = true,
                    b"xfrm" if in_sppr => in_xfrm = true,
                    b"blipFill" if pic_depth == 1 => in_blipfill = true,
                    b"blip" if in_blipfill => {
                        image_rid = attr_value(e, "r:embed");
                    }
                    b"txBody" if sp_depth == 1 => in_txbody = true,
                    b"p" if in_txbody => {
                        in_para = true;
                        para_level = 0;
                        para_text.clear();
                    }
                    b"pPr" if in_para => {
                        para_level =
                            attr_value(e, "lvl").and_then(|v| v.parse().ok()).unwrap_or(0);
                    }
                    b"t" if in_para => in_text = true,
                    b"br" if in_para => {
                        // 改行（a:br）: テキストに改行を追加
                        para_text.push('\n');
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                match e.local_name().as_ref() {
                    b"ph" if in_nvpr => {
                        ph_type = attr_value(e, "type");
                    }
                    b"off" if in_xfrm => {
                        y_offset =
                            attr_value(e, "y").and_then(|v| v.parse().ok()).unwrap_or(0);
                    }
                    b"blip" if in_blipfill => {
                        image_rid = attr_value(e, "r:embed");
                    }
                    b"br" if in_para => {
                        para_text.push('\n');
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                match e.local_name().as_ref() {
                    b"sp" if sp_depth > 0 => {
                        sp_depth -= 1;
                        if sp_depth == 0 {
                            shapes.push(Shape {
                                ph_type: ph_type.take(),
                                y_offset,
                                text: shape_text.clone(),
                                image_rid: image_rid.take(),
                            });
                            in_nvpr = false;
                            in_sppr = false;
                            in_xfrm = false;
                            in_txbody = false;
                            in_para = false;
                            in_text = false;
                        }
                    }
                    b"pic" if pic_depth > 0 => {
                        pic_depth -= 1;
                        if pic_depth == 0 {
                            shapes.push(Shape {
                                ph_type: None,
                                y_offset,
                                text: String::new(),
                                image_rid: image_rid.take(),
                            });
                            in_sppr = false;
                            in_xfrm = false;
                            in_blipfill = false;
                        }
                    }
                    b"nvPr" => in_nvpr = false,
                    b"spPr" => {
                        in_sppr = false;
                        in_xfrm = false;
                    }
                    b"xfrm" => in_xfrm = false,
                    b"blipFill" => in_blipfill = false,
                    b"txBody" => in_txbody = false,
                    b"p" if in_txbody && in_para => {
                        let trimmed = para_text.trim().to_string();
                        if !trimmed.is_empty() {
                            if !shape_text.is_empty() {
                                shape_text.push('\n');
                            }
                            // インデント（bullet レベルに応じて 2 スペース）
                            for _ in 0..para_level {
                                shape_text.push_str("  ");
                            }
                            shape_text.push_str(&trimmed);
                        }
                        in_para = false;
                        in_text = false;
                    }
                    b"t" => in_text = false,
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) if in_text => {
                if let Ok(s) = e.unescape() {
                    para_text.push_str(s.as_ref());
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    shapes
}

// ─── ユーティリティ ───────────────────────────────────────────────────────────

/// rels XML から指定 type suffix に一致する Id → Target マップを返す
///
/// type_suffix 例: "/slide", "/image", "/notes"
fn parse_rels_by_type(xml: &str, type_suffix: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.local_name().as_ref() == b"Relationship" =>
            {
                let rel_type = attr_value(e, "Type").unwrap_or_default();
                if rel_type.ends_with(type_suffix) {
                    let id = attr_value(e, "Id").unwrap_or_default();
                    let target = attr_value(e, "Target").unwrap_or_default();
                    if !id.is_empty() {
                        map.insert(id, target);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
    map
}

/// ppt/xxx/ からの相対パスを ZIP 上の絶対パスに変換する
///
/// 例: base_dir="ppt/slides", target="../media/image1.png" → "ppt/media/image1.png"
fn resolve_ppt_relative_path(base_dir: &str, target: &str) -> String {
    let mut parts: Vec<&str> = base_dir.split('/').collect();
    for segment in target.split('/') {
        match segment {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            s => parts.push(s),
        }
    }
    parts.join("/")
}

/// 画像データをリサイズして JPEG 形式で再エンコードする
///
/// 変換できない場合は None を返す（呼び出し側で元データにフォールバック）
fn resize_and_compress(data: &[u8], max_px: u32, quality: u8) -> Option<Vec<u8>> {
    let img = image::load_from_memory(data).ok()?;
    let img = if img.width() > max_px || img.height() > max_px {
        img.thumbnail(max_px, max_px)
    } else {
        img
    };
    let rgb = img.into_rgb8();
    let mut output = Vec::new();
    let encoder = JpegEncoder::new_with_quality(&mut output, quality);
    rgb.write_with_encoder(encoder).ok()?;
    Some(output)
}

/// XML要素から属性値を取得する（名前空間プレフィックスを無視してローカル名で検索）
///
/// name に ":" が含まれる場合は最後の部分をローカル名として使用する。
/// ただし "r:embed" など異なる prefix を持つ同名属性が存在しないことが前提。
fn attr_value(e: &quick_xml::events::BytesStart, name: &str) -> Option<String> {
    let local = name.split(':').last().unwrap_or(name);
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == local.as_bytes() {
            return String::from_utf8(attr.value.to_vec()).ok();
        }
    }
    None
}

/// 名前空間プレフィックスとローカル名を両方指定して属性値を取得する
///
/// 例: attr_value_with_prefix(e, "r", "id") は r:id="rId2" を取得する。
/// id="867" と r:id="rId2" が共存するケースで使用。
fn attr_value_with_prefix(
    e: &quick_xml::events::BytesStart,
    prefix: &str,
    local: &str,
) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = attr.key;
        let matches_local = key.local_name().as_ref() == local.as_bytes();
        let matches_prefix = match key.prefix() {
            Some(p) => p.as_ref() == prefix.as_bytes(),
            None => false,
        };
        if matches_local && matches_prefix {
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
    let mut entry = archive
        .by_name(name)
        .with_context(|| format!("ZIPエントリが見つかりません: {}", name))?;
    let mut buf = String::new();
    entry
        .read_to_string(&mut buf)
        .with_context(|| format!("ZIPエントリの読み込みに失敗: {}", name))?;
    Ok(buf)
}

/// ZIPアーカイブから指定エントリをバイト列として読み込む（画像等のバイナリ用）
fn read_zip_entry_bytes(
    archive: &mut ZipArchive<BufReader<std::fs::File>>,
    name: &str,
) -> Result<Vec<u8>> {
    let mut entry = archive
        .by_name(name)
        .with_context(|| format!("ZIPエントリが見つかりません: {}", name))?;
    let mut buf = Vec::new();
    entry
        .read_to_end(&mut buf)
        .with_context(|| format!("ZIPエントリの読み込みに失敗: {}", name))?;
    Ok(buf)
}
