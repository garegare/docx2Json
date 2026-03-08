/// RAG向けチャンク分割
///
/// セクションツリーを指定した深さで分割し、各チャンクを個別 JSON ファイルとして出力する。
/// 子セクションの本文は Markdown 見出しとして上位チャンクに統合される。
use std::path::Path;

use anyhow::{Context, Result};
use serde::Serialize;

use crate::models::{Asset, Document, Section};

/// 分割後の単一チャンク（RAG インジェスト向けフラット構造）
#[derive(Debug, Serialize)]
pub struct Chunk {
    /// 元のドキュメントタイトル（ファイル名）
    pub source: String,
    /// ルートから自身への見出しパスリスト
    pub context_path: Vec<String>,
    /// このチャンクのトップ見出し
    pub heading: String,
    /// 子セクションの本文を Markdown 見出しとして統合済みの本文
    pub body_text: String,
    /// 子セクションのアセットを含む全アセット
    pub assets: Vec<Asset>,
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
        println!("  -> {} ({}/{})", out_path.display(), i + 1, total);
    }

    Ok(())
}

/// セクションツリーを再帰的に走査し、指定深さのチャンクを収集する
///
/// `current_depth >= split_level` または葉ノード（children が空）のセクションを
/// チャンク化の起点とする。その配下の子セクションは `flatten_into` で本文に統合する。
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
            let mut body_text = section.body_text.clone();
            let mut assets = section.assets.clone();
            // 子セクションを Markdown 見出し付きで body_text に統合
            flatten_into(&section.children, &mut body_text, &mut assets, 2);
            chunks.push(Chunk {
                source: source.to_string(),
                context_path: section.context_path.clone(),
                heading: section.heading.clone(),
                body_text,
                assets,
            });
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

/// セクションリストを再帰的に走査して body_text・assets に平坦化する
///
/// 各セクションの見出しは Markdown 見出し記法（## / ### …）で挿入する
fn flatten_into(
    sections: &[Section],
    body_text: &mut String,
    assets: &mut Vec<Asset>,
    heading_depth: usize,
) {
    for section in sections {
        // 見出しを Markdown 形式で挿入
        if !section.heading.is_empty() {
            let hashes = "#".repeat(heading_depth.min(6));
            if !body_text.is_empty() {
                body_text.push('\n');
            }
            body_text.push_str(&format!("\n{} {}\n", hashes, section.heading));
        }
        // 本文を追記
        if !section.body_text.is_empty() {
            if !body_text.is_empty() && !body_text.ends_with('\n') {
                body_text.push('\n');
            }
            body_text.push_str(&section.body_text);
        }
        // アセットを収集
        assets.extend_from_slice(&section.assets);
        // 孫以下を再帰的に処理
        flatten_into(&section.children, body_text, assets, heading_depth + 1);
    }
}
