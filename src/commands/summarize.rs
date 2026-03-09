use std::path::PathBuf;

use anyhow::Result;

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

pub fn run(_args: Args) -> Result<()> {
    anyhow::bail!("summarize は Phase 5 で実装予定です")
}
