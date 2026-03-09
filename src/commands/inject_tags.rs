use std::fs;
use std::path::PathBuf;

use anyhow::{Context, Result};
use serde::Deserialize;

/// `inject-tags` サブコマンドの引数
#[derive(clap::Args)]
pub struct Args {
    /// 更新対象の document.json のパス
    #[arg(long)]
    pub input: PathBuf,

    /// タグを注入するセクションの ID
    #[arg(long)]
    pub section_id: String,

    /// 注入するタグの JSON 配列文字列（例: '["認証", "API設計"]'）
    #[arg(long)]
    pub tags: String,

    /// バリデーション用キーワードリスト（keywords.json）のパス
    /// 省略時または --init 指定時はバリデーションをスキップ
    #[arg(long)]
    pub keywords: Option<PathBuf>,

    /// 初回モード: バリデーションをスキップ（keywords.json 未作成時）
    #[arg(long)]
    pub init: bool,

    /// 更新済み document.json の出力パス（入力ファイルの上書きも可）
    #[arg(long)]
    pub output: PathBuf,
}

/// keywords.json のフォーマット
#[derive(Deserialize)]
struct Keywords {
    keywords: Vec<String>,
}

pub fn run(args: Args) -> Result<()> {
    // document.json を読み込む
    let content = fs::read_to_string(&args.input)
        .with_context(|| format!("document.json の読み込みに失敗: {}", args.input.display()))?;
    let mut doc: crate::models::Document = serde_json::from_str(&content)
        .with_context(|| format!("JSON パースに失敗: {}", args.input.display()))?;

    // --tags を JSON 配列としてパース
    let tags: Vec<String> = serde_json::from_str(&args.tags)
        .with_context(|| format!("--tags のパースに失敗（JSON 配列が必要）: {}", args.tags))?;

    // keywords.json を読み込んでバリデーション（--init でなく、--keywords が指定された場合）
    let valid_tags: Vec<String> = if !args.init {
        if let Some(ref kw_path) = args.keywords {
            let kw_content = fs::read_to_string(kw_path)
                .with_context(|| format!("keywords.json の読み込みに失敗: {}", kw_path.display()))?;
            let kw: Keywords = serde_json::from_str(&kw_content)
                .with_context(|| format!("keywords.json のパースに失敗: {}", kw_path.display()))?;

            // バリデーション: keywords に含まれないタグを警告して排除
            let mut validated = Vec::new();
            for tag in &tags {
                if kw.keywords.contains(tag) {
                    validated.push(tag.clone());
                } else {
                    eprintln!(
                        "Warning: タグ \"{}\" は keywords.json に存在しないため排除されます",
                        tag
                    );
                }
            }
            validated
        } else {
            // --keywords 未指定 → バリデーションスキップ
            tags
        }
    } else {
        // --init モード → バリデーションスキップ
        tags
    };

    // 対象セクションを再帰的に検索してタグを注入
    let found = inject_tags_recursive(&mut doc.sections, &args.section_id, &valid_tags);
    if !found {
        anyhow::bail!(
            "セクション ID \"{}\" が document.json に見つかりません",
            args.section_id
        );
    }

    // 更新済み document.json を出力
    if let Some(parent) = args.output.parent() {
        if !parent.as_os_str().is_empty() {
            fs::create_dir_all(parent)
                .with_context(|| format!("出力ディレクトリの作成に失敗: {}", parent.display()))?;
        }
    }
    let out_content = serde_json::to_string_pretty(&doc)
        .context("document.json のシリアライズに失敗")?;
    fs::write(&args.output, out_content)
        .with_context(|| format!("出力ファイルの書き込みに失敗: {}", args.output.display()))?;

    eprintln!(
        "inject-tags: section_id={} に {} 個のタグを注入しました → {}",
        args.section_id,
        valid_tags.len(),
        args.output.display()
    );
    Ok(())
}

/// セクションツリーを再帰的に検索し、指定 ID のセクションに ai_tags を注入する。
/// 見つかった場合は `true` を返す。
fn inject_tags_recursive(
    sections: &mut Vec<crate::models::Section>,
    section_id: &str,
    tags: &[String],
) -> bool {
    for section in sections {
        if section.id == section_id {
            section.metadata.ai_tags = tags.to_vec();
            return true;
        }
        if inject_tags_recursive(&mut section.children, section_id, tags) {
            return true;
        }
    }
    false
}
