pub mod docx;
pub mod emf;
pub mod pptx;
pub mod xlsx;
pub mod xlsx_advanced;

use std::path::Path;

use anyhow::{Context, Result};

use crate::config::Config;
use crate::models::{Document, Section};

pub fn parse_file(path: &Path, config: &Config) -> Result<Document> {
    let mut doc = match path.extension().and_then(|e| e.to_str()) {
        Some("docx") => docx::parse(path, config)
            .with_context(|| format!("DOCXパース失敗: {}", path.display()))?,
        Some("pptx") => pptx::parse(path, config)
            .with_context(|| format!("PPTXパース失敗: {}", path.display()))?,
        Some("xlsx") | Some("xlsm") => {
            // xlsx_heading.enabled == true のとき神エクセル対応パーサーに切り替え
            // xlsm（マクロ有効ブック）は xlsx と同じ構造のため同一パーサーで処理する
            if config.xlsx.heading.as_ref().is_some_and(|h| h.enabled) {
                xlsx_advanced::parse(path, config)
                    .with_context(|| format!("XLSX(advanced)パース失敗: {}", path.display()))?
            } else {
                xlsx::parse(path, config)
                    .with_context(|| format!("XLSXパース失敗: {}", path.display()))?
            }
        }
        ext => anyhow::bail!("未対応のファイル形式: {:?}", ext),
    };

    // 全セクションに context_path（ルートから自身への見出しパス）を付与
    fill_context_path(&mut doc.sections, &[]);
    // 全セクションに安定 ID（FNV-1a ハッシュ）を付与
    fill_section_id(&mut doc.sections, &doc.title);
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

/// セクションツリーを再帰的に走査し、各セクションに安定 ID を付与する。
///
/// ID は「文書タイトル + context_path を連結した文字列」の FNV-1a 64bit ハッシュの 16進数表現。
/// 追加クレート不要で、同一ドキュメント・同一パスであれば実行間で安定する。
/// 64bit を採用することで 32bit（数千セクションで衝突期待値 1）より衝突リスクを大幅に低減。
fn fill_section_id(sections: &mut Vec<Section>, title: &str) {
    for section in sections {
        let key = format!("{}\x00{}", title, section.context_path.join("\x00"));
        section.id = fnv1a_hex(&key);
        fill_section_id(&mut section.children, title);
    }
}

/// XML 要素から属性値を取得する（名前空間プレフィックスを無視してローカル名で検索）。
///
/// `name` に ":" が含まれる場合は最後の部分をローカル名として使用する。
pub(crate) fn attr_value(e: &quick_xml::events::BytesStart, name: &str) -> Option<String> {
    let local = name.split(':').next_back().unwrap_or(name);
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == local.as_bytes() {
            return String::from_utf8(attr.value.to_vec()).ok();
        }
    }
    None
}

/// 文字列の FNV-1a 64bit ハッシュを 16文字の 16進数文字列として返す。
fn fnv1a_hex(s: &str) -> String {
    const FNV_OFFSET: u64 = 14695981039346656037;
    const FNV_PRIME: u64 = 1099511628211;
    let hash = s.bytes().fold(FNV_OFFSET, |acc, b| {
        acc.wrapping_mul(FNV_PRIME) ^ (b as u64)
    });
    format!("{:016x}", hash)
}
