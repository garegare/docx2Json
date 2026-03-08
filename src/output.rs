use std::path::Path;
use crate::models::Document;

type Error = Box<dyn std::error::Error + Send + Sync>;

/// DocumentをJSONファイルに書き出す
/// 出力パスは入力パスの拡張子を .json に置換したもの
pub fn write_json(doc: &Document, input_path: &Path, output_dir: Option<&Path>) -> Result<(), Error> {
    let stem = input_path
        .file_stem()
        .and_then(|s| s.to_str())
        .ok_or("Invalid file name")?;

    let out_path = if let Some(dir) = output_dir {
        dir.join(format!("{}.json", stem))
    } else {
        input_path.with_extension("json")
    };

    let json = serde_json::to_string_pretty(doc)?;
    std::fs::write(&out_path, json)?;
    println!("  -> {}", out_path.display());
    Ok(())
}
