use std::path::Path;

use anyhow::{Context, Result};

use crate::models::Document;

/// DocumentをJSONファイルに書き出す
/// 出力パスは入力パスの拡張子を .json に置換したもの
pub fn write_json(doc: &Document, input_path: &Path, output_dir: Option<&Path>) -> Result<()> {
    let stem = input_path
        .file_stem()
        .and_then(|s| s.to_str())
        .with_context(|| format!("無効なファイル名: {}", input_path.display()))?;

    let out_path = if let Some(dir) = output_dir {
        dir.join(format!("{}.json", stem))
    } else {
        input_path.with_extension("json")
    };

    let json = serde_json::to_string_pretty(doc)
        .context("JSONシリアライズに失敗")?;
    std::fs::write(&out_path, &json)
        .with_context(|| format!("ファイルへの書き込みに失敗: {}", out_path.display()))?;
    println!("  -> {}", out_path.display());
    Ok(())
}
