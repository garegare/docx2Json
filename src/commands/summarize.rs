use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

use anyhow::{Context, Result};
use serde::Serialize;

/// `summarize` サブコマンドの引数
#[derive(clap::Args)]
pub struct Args {
    /// document.json 単体またはそれを含むディレクトリのパス
    #[arg(long)]
    pub input: PathBuf,

    /// タグ使用統計 JSON（tags_summary.json）の出力パス
    #[arg(long)]
    pub output: PathBuf,
}

/// tags_summary.json のフォーマット
#[derive(Serialize)]
struct TagsSummary {
    generated_at: String,
    total_sections: usize,
    tagged_sections: usize,
    untagged_sections: usize,
    /// タグ名 → 出現セクション数
    tag_counts: HashMap<String, usize>,
    /// 出現頻度上位のタグ（降順）
    top_tags: Vec<String>,
}

pub fn run(args: Args) -> Result<()> {
    // 対象 JSON ファイルを収集
    let json_files = collect_json_files(&args.input);
    if json_files.is_empty() {
        anyhow::bail!(
            "document.json が見つかりません: {}",
            args.input.display()
        );
    }

    // 統計を集計
    let mut total_sections = 0usize;
    let mut tagged_sections = 0usize;
    let mut tag_counts: HashMap<String, usize> = HashMap::new();

    for path in &json_files {
        let content = fs::read_to_string(path)
            .with_context(|| format!("ファイルの読み込みに失敗: {}", path.display()))?;
        let doc: crate::models::Document = match serde_json::from_str(&content) {
            Ok(d) => d,
            Err(e) => {
                eprintln!("Warning: JSON パースに失敗（スキップ）: {} - {}", path.display(), e);
                continue;
            }
        };

        accumulate_stats(
            &doc.sections,
            &mut total_sections,
            &mut tagged_sections,
            &mut tag_counts,
        );
    }

    let untagged_sections = total_sections - tagged_sections;

    // top_tags: 出現頻度降順でソート（同数の場合は辞書順）
    let mut top_tags: Vec<(String, usize)> = tag_counts
        .iter()
        .map(|(k, v)| (k.clone(), *v))
        .collect();
    top_tags.sort_by(|a, b| b.1.cmp(&a.1).then_with(|| a.0.cmp(&b.0)));
    let top_tags: Vec<String> = top_tags.into_iter().map(|(k, _)| k).collect();

    let summary = TagsSummary {
        generated_at: current_iso8601(),
        total_sections,
        tagged_sections,
        untagged_sections,
        tag_counts,
        top_tags,
    };

    // 出力ファイルを書き出す
    if let Some(parent) = args.output.parent() {
        if !parent.as_os_str().is_empty() {
            fs::create_dir_all(parent)
                .with_context(|| format!("出力ディレクトリの作成に失敗: {}", parent.display()))?;
        }
    }
    let out_content = serde_json::to_string_pretty(&summary)
        .context("tags_summary.json のシリアライズに失敗")?;
    fs::write(&args.output, out_content)
        .with_context(|| format!("出力ファイルの書き込みに失敗: {}", args.output.display()))?;

    eprintln!(
        "summarize: {} ファイル / {} セクション集計 → {}",
        json_files.len(),
        total_sections,
        args.output.display()
    );
    Ok(())
}

/// セクションツリーを再帰的に走査してタグ統計を集計する
fn accumulate_stats(
    sections: &[crate::models::Section],
    total_sections: &mut usize,
    tagged_sections: &mut usize,
    tag_counts: &mut HashMap<String, usize>,
) {
    for section in sections {
        *total_sections += 1;
        if !section.metadata.ai_tags.is_empty() {
            *tagged_sections += 1;
            for tag in &section.metadata.ai_tags {
                *tag_counts.entry(tag.clone()).or_insert(0) += 1;
            }
        }
        accumulate_stats(&section.children, total_sections, tagged_sections, tag_counts);
    }
}

/// ディレクトリを再帰的にスキャンして .json ファイルを収集する
/// input がファイルの場合はそのファイル単体を返す
fn collect_json_files(input: &PathBuf) -> Vec<PathBuf> {
    if input.is_file() {
        let ext = input.extension().and_then(|e| e.to_str()).unwrap_or("");
        if ext == "json" {
            return vec![input.clone()];
        }
        return Vec::new();
    }

    let mut files = Vec::new();
    if let Ok(entries) = fs::read_dir(input) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                files.extend(collect_json_files(&path));
            } else {
                let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("");
                if ext == "json" {
                    files.push(path);
                }
            }
        }
    }
    files
}

/// 現在時刻を ISO 8601 形式（UTC）で返す（外部クレート不使用）
fn current_iso8601() -> String {
    // std::time を使った簡易実装。外部クレート (chrono 等) を使わずに生成する。
    // 秒精度で十分なためミリ秒は省略。
    let secs = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();

    // UNIX タイムスタンプから年月日時分秒を計算
    let (year, month, day, hour, min, sec) = unix_to_datetime(secs);
    format!(
        "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
        year, month, day, hour, min, sec
    )
}

/// UNIX タイムスタンプ（秒）を (year, month, day, hour, min, sec) に変換する
fn unix_to_datetime(ts: u64) -> (u32, u32, u32, u32, u32, u32) {
    let sec = (ts % 60) as u32;
    let min = ((ts / 60) % 60) as u32;
    let hour = ((ts / 3600) % 24) as u32;

    let days = ts / 86400;
    // グレゴリオ暦の計算（1970-01-01 からの通算日数）
    let mut year = 1970u32;
    let mut remaining = days;
    loop {
        let days_in_year = if is_leap(year) { 366 } else { 365 };
        if remaining < days_in_year {
            break;
        }
        remaining -= days_in_year;
        year += 1;
    }
    let months = [31u64, if is_leap(year) { 29 } else { 28 }, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    let mut month = 1u32;
    for &m in &months {
        if remaining < m {
            break;
        }
        remaining -= m;
        month += 1;
    }
    let day = remaining as u32 + 1;

    (year, month, day, hour, min, sec)
}

fn is_leap(year: u32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)
}
