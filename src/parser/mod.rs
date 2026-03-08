pub mod docx;
pub mod xlsx;

use std::path::Path;

use anyhow::{Context, Result};

use crate::config::Config;
use crate::models::Document;

pub fn parse_file(path: &Path, config: &Config) -> Result<Document> {
    match path.extension().and_then(|e| e.to_str()) {
        Some("docx") => docx::parse(path, config)
            .with_context(|| format!("DOCXパース失敗: {}", path.display())),
        Some("xlsx") => xlsx::parse(path)
            .with_context(|| format!("XLSXパース失敗: {}", path.display())),
        ext => anyhow::bail!("未対応のファイル形式: {:?}", ext),
    }
}
