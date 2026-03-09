use std::path::PathBuf;

use anyhow::Result;

/// `inject-tags` サブコマンドの引数
#[derive(clap::Args)]
pub struct Args {
    /// 更新対象の document.json のパス
    #[arg(long)]
    pub input: PathBuf,

    /// タグを注入するセクションの ID
    #[arg(long)]
    pub section_id: String,

    /// 注入するタグの JSON 配列文字列（例: '["認証", "API設計"]'）
    #[arg(long)]
    pub tags: String,

    /// バリデーション用キーワードリスト（keywords.json）のパス
    /// 省略時または --init 指定時はバリデーションをスキップ
    #[arg(long)]
    pub keywords: Option<PathBuf>,

    /// 初回モード: バリデーションをスキップ（keywords.json 未作成時）
    #[arg(long)]
    pub init: bool,

    /// 更新済み document.json の出力パス（入力ファイルの上書きも可）
    #[arg(long)]
    pub output: PathBuf,
}

pub fn run(_args: Args) -> Result<()> {
    anyhow::bail!("inject-tags は Phase 4 で実装予定です")
}
