use std::fs;
use std::io::{BufWriter, Write};
use std::path::PathBuf;

use anyhow::{Context, Result};
use serde::Serialize;

/// `extract-candidates` サブコマンドの引数
#[derive(clap::Args)]
pub struct Args {
    /// パース済み document.json のパス
    #[arg(long)]
    pub input: PathBuf,

    /// 出力 JSONL ファイルのパス
    #[arg(long)]
    pub output: PathBuf,

    /// body_text の最大文字数（0 = 制限なし、デフォルト: 0）
    #[arg(long, default_value = "0", value_name = "CHARS")]
    pub max_body_chars: usize,
}

/// JSONL 1行分のレコード（assets / children を除いたコンパクト形式）
#[derive(Serialize)]
struct CandidateRecord<'a> {
    id: &'a str,
    context_path: &'a [String],
    heading: &'a str,
    body_text: std::borrow::Cow<'a, str>,
}

pub fn run(args: Args) -> Result<()> {
    // document.json を読み込む
    let content = fs::read_to_string(&args.input)
        .with_context(|| format!("document.json の読み込みに失敗: {}", args.input.display()))?;
    let doc: crate::models::Document = serde_json::from_str(&content)
        .with_context(|| format!("JSON パースに失敗: {}", args.input.display()))?;

    // 出力ファイルを開く
    let out_file = fs::File::create(&args.output)
        .with_context(|| format!("出力ファイルの作成に失敗: {}", args.output.display()))?;
    let mut writer = BufWriter::new(out_file);

    // セクションを再帰的に走査して JSONL を書き出す
    let mut count = 0usize;
    write_sections(&doc.sections, &mut writer, args.max_body_chars, &mut count)?;
    writer.flush()?;

    eprintln!("extract-candidates: {} セクションを書き出しました → {}", count, args.output.display());
    Ok(())
}

/// セクションツリーを再帰的に走査して JSONL を書き出す
fn write_sections(
    sections: &[crate::models::Section],
    writer: &mut impl Write,
    max_body_chars: usize,
    count: &mut usize,
) -> Result<()> {
    for section in sections {
        // body_text を必要に応じて切り詰める
        let body_text: std::borrow::Cow<str> =
            if max_body_chars > 0 && section.body_text.chars().count() > max_body_chars {
                // Unicode 境界で安全に切り詰める
                std::borrow::Cow::Owned(truncate_unicode(&section.body_text, max_body_chars))
            } else {
                std::borrow::Cow::Borrowed(&section.body_text)
            };

        let record = CandidateRecord {
            id: &section.id,
            context_path: &section.context_path,
            heading: &section.heading,
            body_text,
        };

        // serde_json はデフォルトで UTF-8 を直接出力する（日本語等の非 ASCII 文字を
        // \uXXXX にエスケープしない）。Anthropic API を含む主要 LLM API は
        // UTF-8 JSONL をそのまま受け付けるため互換性に問題はない。
        let line = serde_json::to_string(&record)
            .with_context(|| format!("JSON シリアライズに失敗: id={}", section.id))?;
        writeln!(writer, "{}", line)?;
        *count += 1;

        // 子セクションを再帰処理
        write_sections(&section.children, writer, max_body_chars, count)?;
    }
    Ok(())
}

/// UTF-8 文字境界を考慮して最大 `max_chars` 文字で切り詰める
fn truncate_unicode(s: &str, max_chars: usize) -> String {
    s.char_indices()
        .nth(max_chars)
        .map(|(byte_idx, _)| s[..byte_idx].to_string())
        .unwrap_or_else(|| s.to_string())
}
