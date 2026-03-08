use std::path::Path;
use crate::models::{Document, Section};

type Error = Box<dyn std::error::Error + Send + Sync>;

/// XLSXファイルを解析してDocumentを返す（現在はスタブ実装）
pub fn parse(path: &Path) -> Result<Document, Error> {
    let title = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("Untitled")
        .to_string();

    // TODO: XLSXパーサーの本実装
    // 1. ZIPを展開して xl/workbook.xml からシート名を取得
    // 2. xl/worksheets/sheet*.xml を走査してセル値を取得
    // 3. 各シートを Section として構造化
    let section = Section {
        heading: format!("(XLSX stub: {})", title),
        body_text: "XLSXパーサーは未実装です。".to_string(),
        assets: Vec::new(),
        children: Vec::new(),
    };

    Ok(Document {
        title,
        sections: vec![section],
    })
}
