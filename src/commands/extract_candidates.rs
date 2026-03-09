use std::path::PathBuf;

use anyhow::Result;

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

pub fn run(_args: Args) -> Result<()> {
    anyhow::bail!("extract-candidates は Phase 3 で実装予定です")
}
