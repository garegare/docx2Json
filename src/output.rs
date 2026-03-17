use std::path::Path;

use anyhow::{Context, Result};

use crate::config::Config;
use crate::models::{Document, Section};

/// DocumentをJSONファイルに書き出す
/// 出力パスは入力パスの拡張子を .json に置換したもの
pub fn write_json(doc: &Document, input_path: &Path, output_dir: Option<&Path>, config: &Config) -> Result<()> {
    let stem = input_path
        .file_stem()
        .and_then(|s| s.to_str())
        .with_context(|| format!("無効なファイル名: {}", input_path.display()))?;

    let out_path = if let Some(dir) = output_dir {
        dir.join(format!("{}.json", stem))
    } else {
        input_path.with_extension("json")
    };

    let json = if !config.output.include_body_text || !config.output.include_base64 {
        // オプション無効時はドキュメントを複製してフィールドを除去してからシリアライズ
        let mut doc = doc.clone();
        apply_output_config(&mut doc.sections, config);
        serde_json::to_string_pretty(&doc).context("JSONシリアライズに失敗")?
    } else {
        serde_json::to_string_pretty(doc).context("JSONシリアライズに失敗")?
    };

    std::fs::write(&out_path, &json)
        .with_context(|| format!("ファイルへの書き込みに失敗: {}", out_path.display()))?;
    Ok(())
}

/// 出力設定に従いセクションツリーのフィールドを除去する（in-place）
fn apply_output_config(sections: &mut Vec<Section>, config: &Config) {
    for sec in sections.iter_mut() {
        if !config.output.include_body_text {
            sec.body_text.clear();
        }
        if !config.output.include_base64 {
            for asset in &mut sec.assets {
                asset.data.clear();
            }
        }
        apply_output_config(&mut sec.children, config);
    }
}
