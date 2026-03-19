use std::collections::HashMap;
use std::io::BufReader;
use std::path::PathBuf;

use anyhow::{Context, Result};
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use serde::Serialize;
use zip::ZipArchive;

/// `inspect-styles` サブコマンドの引数
#[derive(clap::Args)]
pub struct Args {
    /// 解析する DOCX ファイルのパス
    #[arg(long)]
    pub input: PathBuf,

    /// 出力先ファイルパス（省略時は標準出力）
    #[arg(long)]
    pub output: Option<PathBuf>,
}

#[derive(Debug)]
struct StyleEntry {
    style_id: String,
    name: String,
    outline_lvl: usize, // 0-based (0 = レベル1)
}

/// stdout / ファイルに出力する JSON 構造
#[derive(Serialize)]
struct InspectOutput {
    /// docx2json.json にそのまま貼り付けられる設定スニペット
    docx: DocxSnippet,
    /// スタイル詳細（参考情報）
    styles: Vec<StyleDetail>,
}

#[derive(Serialize)]
struct DocxSnippet {
    heading_styles: HashMap<String, usize>,
}

#[derive(Serialize)]
struct StyleDetail {
    style_id: String,
    name: String,
    level: usize,
}

pub fn run(args: Args) -> Result<()> {
    let path = &args.input;
    let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("");
    if !matches!(ext, "docx") {
        anyhow::bail!("DOCX ファイルを指定してください（拡張子が .docx であること）: {}", path.display());
    }

    let styles = parse_styles_xml(path)
        .with_context(|| format!("styles.xml のパースに失敗: {}", path.display()))?;

    if styles.is_empty() {
        eprintln!("見出しスタイル（outlineLvl 付き）は検出されませんでした。");
        eprintln!("ヒント: このファイルは Word 標準の見出しスタイルを使用していない可能性があります。");
        eprintln!("       代わりに --config で ppr_underline_as_heading や run_underline_as_heading を試してください。");
        return Ok(());
    }

    // heading_styles マップ: styleId と name の両方をキーとして登録（重複時は1エントリ）
    let mut heading_styles: HashMap<String, usize> = HashMap::new();
    let mut style_details: Vec<StyleDetail> = Vec::new();

    eprintln!("見出しスタイルを {} 件検出: {}", styles.len(), path.display());
    for entry in &styles {
        let level = entry.outline_lvl + 1;
        eprintln!("  Level {}: {:?}  (style_id: {:?})", level, entry.name, entry.style_id);

        heading_styles.insert(entry.style_id.clone(), level);
        if entry.name != entry.style_id {
            heading_styles.insert(entry.name.clone(), level);
        }
        style_details.push(StyleDetail {
            style_id: entry.style_id.clone(),
            name: entry.name.clone(),
            level,
        });
    }

    let out = InspectOutput {
        docx: DocxSnippet { heading_styles },
        styles: style_details,
    };
    let json = serde_json::to_string_pretty(&out)?;

    match &args.output {
        Some(out_path) => {
            std::fs::write(out_path, &json)
                .with_context(|| format!("出力ファイルの書き込みに失敗: {}", out_path.display()))?;
            eprintln!("設定スニペットを書き出しました → {}", out_path.display());
        }
        None => println!("{}", json),
    }

    Ok(())
}

/// word/styles.xml をパースして outlineLvl を持つ段落スタイルを返す
fn parse_styles_xml(docx_path: &PathBuf) -> Result<Vec<StyleEntry>> {
    let file = std::fs::File::open(docx_path)
        .with_context(|| format!("ファイルを開けません: {}", docx_path.display()))?;
    let mut archive = ZipArchive::new(BufReader::new(file))
        .context("ZIPアーカイブとして開けません（破損または非 DOCX ファイルの可能性）")?;

    let xml = {
        let mut entry = archive
            .by_name("word/styles.xml")
            .context("word/styles.xml が見つかりません")?;
        let mut buf = String::new();
        std::io::Read::read_to_string(&mut entry, &mut buf)
            .context("word/styles.xml の読み込みに失敗")?;
        buf
    };

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    let mut styles: Vec<StyleEntry> = Vec::new();
    let mut current_id: Option<String> = None;
    let mut is_paragraph: bool = false;
    let mut current_name: Option<String> = None;
    let mut current_lvl: Option<usize> = None;
    let mut in_style = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                match e.local_name().as_ref() {
                    b"style" => {
                        in_style = true;
                        current_id = attr_value(e, "styleId");
                        is_paragraph = attr_value(e, "type").as_deref() == Some("paragraph");
                        current_name = None;
                        current_lvl = None;
                    }
                    b"name" if in_style => {
                        if let Some(val) = attr_value(e, "val") {
                            current_name = Some(val);
                        }
                    }
                    b"outlineLvl" if in_style => {
                        if let Some(val) = attr_value(e, "val") {
                            if let Ok(n) = val.parse::<usize>() {
                                // outlineLvl 9 はボディテキスト扱い（見出しではない）
                                if n < 9 {
                                    current_lvl = Some(n);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"style" => {
                if in_style && is_paragraph {
                    if let (Some(id), Some(name), Some(lvl)) =
                        (current_id.take(), current_name.take(), current_lvl.take())
                    {
                        styles.push(StyleEntry {
                            style_id: id,
                            name,
                            outline_lvl: lvl,
                        });
                    }
                }
                // いずれにせよリセット
                current_id = None;
                current_name = None;
                current_lvl = None;
                is_paragraph = false;
                in_style = false;
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::Error::from(e)),
            _ => {}
        }
    }

    styles.sort_by_key(|s| s.outline_lvl);
    Ok(styles)
}

/// XML 要素から属性値を取得する（名前空間プレフィックスを無視）
fn attr_value(e: &quick_xml::events::BytesStart, name: &str) -> Option<String> {
    let local = name.split(':').next_back().unwrap_or(name);
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == local.as_bytes() {
            return String::from_utf8(attr.value.to_vec()).ok();
        }
    }
    None
}
