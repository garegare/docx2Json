/// RAG向けチャンク分割
///
/// セクションツリーを指定した深さで分割し、各チャンクを個別 JSON ファイルとして出力する。
/// 子セクションは `children` フィールドに再帰的に保持される（body_text へのフラット化は行わない）。
use std::path::Path;

use anyhow::{Context, Result};
use serde::Serialize;

use crate::models::{Asset, Document, Section};

/// 分割後の単一チャンク（RAG インジェスト向け構造）
#[derive(Debug, Serialize)]
pub struct Chunk {
    /// 元のドキュメントタイトル（ファイル名）
    pub source: String,
    /// ルートから自身への見出しパスリスト
    pub context_path: Vec<String>,
    /// このチャンクのトップ見出し
    pub heading: String,
    /// このセクション直下の本文（子セクションの内容は含まない）
    pub body_text: String,
    /// このセクション直下のアセット
    pub assets: Vec<Asset>,
    /// 子セクションのチャンク（再帰的）
    pub children: Vec<Chunk>,
}

/// Document をチャンクに分割して個別 JSON ファイルへ書き出す
///
/// # 引数
/// - `split_level`: 分割する深さ（1 = 最上位セクション単位、2 = 2階層目単位、…）
///   指定した深さに満たないリーフセクションは深さに関わらずチャンク化される
pub fn write_chunks(
    doc: &Document,
    input_path: &Path,
    output_dir: Option<&Path>,
    split_level: usize,
) -> Result<()> {
    let stem = input_path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("output");

    let chunks = collect_chunks(&doc.sections, &doc.title, split_level, 1);
    let total = chunks.len();

    if total == 0 {
        println!("  (チャンクなし: セクションが存在しません)");
        return Ok(());
    }

    for (i, chunk) in chunks.iter().enumerate() {
        let filename = format!("{}_chunk_{:03}.json", stem, i + 1);
        let out_path = if let Some(dir) = output_dir {
            dir.join(&filename)
        } else {
            input_path
                .parent()
                .unwrap_or(Path::new("."))
                .join(&filename)
        };

        let json = serde_json::to_string_pretty(chunk)
            .context("チャンクのJSONシリアライズに失敗")?;
        std::fs::write(&out_path, &json)
            .with_context(|| format!("チャンクファイルへの書き込みに失敗: {}", out_path.display()))?;
    }

    Ok(())
}

/// セクションツリーを再帰的に走査し、指定深さのチャンクを収集する
///
/// `current_depth >= split_level` または葉ノード（children が空）のセクションを
/// チャンク化の起点とする。その配下の子セクションは `children` フィールドに再帰的に保持する。
fn collect_chunks(
    sections: &[Section],
    source: &str,
    split_level: usize,
    current_depth: usize,
) -> Vec<Chunk> {
    let mut chunks = Vec::new();
    for section in sections {
        if current_depth >= split_level || section.children.is_empty() {
            // このセクションをチャンクの起点とする
            // 子セクションは children フィールドに再帰的に保持（body_text へのフラット化は行わない）
            chunks.push(section_to_chunk(section, source));
        } else {
            // まだ目標の深さに達していない: 子を再帰的に収集
            chunks.extend(collect_chunks(
                &section.children,
                source,
                split_level,
                current_depth + 1,
            ));
        }
    }
    chunks
}

/// Section を Chunk に変換する（子セクションも再帰的に Chunk 化）
fn section_to_chunk(section: &Section, source: &str) -> Chunk {
    Chunk {
        source: source.to_string(),
        context_path: section.context_path.clone(),
        heading: section.heading.clone(),
        body_text: section.body_text.clone(),
        assets: section.assets.clone(),
        children: section.children.iter()
            .map(|c| section_to_chunk(c, source))
            .collect(),
    }
}
