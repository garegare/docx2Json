pub mod docx;
pub mod xlsx;

use std::path::Path;

use anyhow::{Context, Result};

use crate::config::Config;
use crate::models::{Document, Section};

pub fn parse_file(path: &Path, config: &Config) -> Result<Document> {
    let mut doc = match path.extension().and_then(|e| e.to_str()) {
        Some("docx") => docx::parse(path, config)
            .with_context(|| format!("DOCXパース失敗: {}", path.display()))?,
        Some("xlsx") => xlsx::parse(path)
            .with_context(|| format!("XLSXパース失敗: {}", path.display()))?,
        ext => anyhow::bail!("未対応のファイル形式: {:?}", ext),
    };

    // 全セクションに context_path（ルートから自身への見出しパス）を付与
    fill_context_path(&mut doc.sections, &[]);
    Ok(doc)
}

/// セクションツリーを再帰的に走査し、各セクションの context_path を設定する。
///
/// context_path はルートからそのセクション自身の見出しまでを含むリストで、
/// チャンク分割後も文書内の位置をAIが把握できるようにする。
///
/// 例: 第1章 > 1.1節 > 1.1.1項 の場合
///   context_path = ["第1章 導入", "1.1 背景", "1.1.1 詳細"]
fn fill_context_path(sections: &mut Vec<Section>, parent_path: &[String]) {
    for section in sections {
        let mut path = parent_path.to_vec();
        path.push(section.heading.clone());
        section.context_path = path.clone();
        fill_context_path(&mut section.children, &path);
    }
}
