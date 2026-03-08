pub mod docx;
pub mod xlsx;

use std::path::Path;
use crate::config::Config;
use crate::models::Document;

pub fn parse_file(path: &Path, config: &Config) -> Result<Document, Box<dyn std::error::Error + Send + Sync>> {
    match path.extension().and_then(|e| e.to_str()) {
        Some("docx") => docx::parse(path, config),
        Some("xlsx") => xlsx::parse(path),
        ext => Err(format!("Unsupported file type: {:?}", ext).into()),
    }
}
