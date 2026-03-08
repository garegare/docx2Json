mod ai;
mod config;
mod models;
mod output;
mod parser;

use std::path::PathBuf;

use clap::Parser;
use rayon::prelude::*;

#[derive(Parser)]
#[command(name = "docx2json", about = "DOCX/XLSX を AI向け構造化JSON に変換する")]
struct Cli {
    /// 入力ディレクトリ（.docx / .xlsx を再帰的にスキャン）
    #[arg(short, long, default_value = ".")]
    input: PathBuf,

    /// 出力ディレクトリ（省略時は入力ファイルと同じ場所に出力）
    #[arg(short, long)]
    output: Option<PathBuf>,

    /// AI変換を有効化（ANTHROPIC_API_KEY 環境変数が必要）
    #[arg(long)]
    ai: bool,

    /// 設定ファイルのパス（省略時は入力ディレクトリ内の docx2json.json を自動検索）
    #[arg(long)]
    config: Option<PathBuf>,
}

fn main() {
    let cli = Cli::parse();

    // 設定ファイルを読み込む
    let input_dir = if cli.input.is_file() {
        cli.input.parent().unwrap_or(&cli.input).to_path_buf()
    } else {
        cli.input.clone()
    };
    let cfg = config::Config::load(cli.config.as_deref(), &input_dir);

    // 出力ディレクトリを作成
    if let Some(ref out) = cli.output {
        if let Err(e) = std::fs::create_dir_all(out) {
            eprintln!("Error creating output directory: {}", e);
            std::process::exit(1);
        }
    }

    // 対象ファイルを収集
    let files = collect_files(&cli.input);
    if files.is_empty() {
        eprintln!("No .docx or .xlsx files found in: {}", cli.input.display());
        std::process::exit(1);
    }

    println!("Processing {} file(s)...", files.len());

    // Rayon で並列処理
    let results: Vec<_> = files
        .par_iter()
        .map(|path| {
            println!("Parsing: {}", path.display());
            let result = parser::parse_file(path, &cfg)
                .map(|doc| if cli.ai { ai::transform(doc) } else { doc })
                .and_then(|doc| output::write_json(&doc, path, cli.output.as_deref()));
            (path, result)
        })
        .collect();

    // 結果サマリー
    let (ok, err): (Vec<_>, Vec<_>) = results.iter().partition(|(_, r)| r.is_ok());
    println!("\nDone: {} succeeded, {} failed.", ok.len(), err.len());
    for (path, e) in err.iter() {
        if let Err(e) = e {
            eprintln!("  FAILED {}: {}", path.display(), e);
        }
    }
}

/// ディレクトリを再帰的にスキャンして .docx/.xlsx ファイルを返す
fn collect_files(dir: &std::path::Path) -> Vec<PathBuf> {
    if dir.is_file() {
        let ext = dir.extension().and_then(|e| e.to_str()).unwrap_or("");
        if matches!(ext, "docx" | "xlsx") {
            return vec![dir.to_path_buf()];
        }
        return Vec::new();
    }

    let mut files = Vec::new();
    if let Ok(entries) = std::fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                files.extend(collect_files(&path));
            } else {
                let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("");
                if matches!(ext, "docx" | "xlsx") {
                    files.push(path);
                }
            }
        }
    }
    files
}
