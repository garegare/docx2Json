mod commands;
mod config;
mod models;
mod output;
mod parser;
mod splitter;

use std::path::PathBuf;

use clap::{Args, Parser, Subcommand};
use indicatif::{ProgressBar, ProgressStyle};
use rayon::prelude::*;

#[derive(Parser)]
#[command(name = "docx2json", about = "DOCX/XLSX/PPTX を AI向け構造化JSON に変換する")]
struct Cli {
    /// サブコマンド（省略時は parse として動作）
    #[command(subcommand)]
    command: Option<Commands>,

    // ---- 後方互換: サブコマンドなしのとき parse として動作するオプション群 ----
    #[command(flatten)]
    args: ParseArgs,
}

#[derive(Subcommand)]
enum Commands {
    /// DOCX/XLSX/PPTX ファイルをパースして document.json を生成する（デフォルト動作と同一）
    Parse(ParseArgs),
    /// document.json から LLM 向け候補テキストを JSONL 形式で抽出する
    ExtractCandidates(commands::extract_candidates::Args),
    /// セクションに AI タグを注入してバリデーションする
    InjectTags(commands::inject_tags::Args),
    /// DOCX ファイルの見出しスタイルを走査して docx2json.json 用設定スニペットを出力する
    InspectStyles(commands::inspect_styles::Args),
    /// 複数の document.json からタグ使用統計を集計する
    Summarize(commands::summarize::Args),
    /// document.json を AsciiDoc 形式に変換する
    ToAsciidoc(commands::to_asciidoc::Args),
}

/// `parse` サブコマンド（またはサブコマンドなし時の後方互換）の引数
#[derive(Args)]
struct ParseArgs {
    /// 入力ディレクトリまたはファイル（.docx / .xlsx / .pptx を再帰的にスキャン）
    #[arg(short, long, default_value = ".")]
    input: PathBuf,

    /// 出力ディレクトリ（省略時は入力ファイルと同じ場所に出力）
    #[arg(short, long)]
    output: Option<PathBuf>,

    /// 設定ファイルのパス（省略時は入力ディレクトリ内の docx2json.json を自動検索）
    #[arg(long)]
    config: Option<PathBuf>,

    /// セクション単位のチャンク分割: 指定した深さ（1 = 最上位）でセクションを分割し
    /// セクションごとに個別 JSON ファイルを出力する（RAG 向け）
    #[arg(long, value_name = "LEVEL")]
    split: Option<usize>,

    /// 画像の最大辺長（ピクセル）。超過する画像をこのサイズにリサイズし JPEG 再エンコードする。
    /// 設定ファイルの `image.max_px` より優先される。
    #[arg(long, value_name = "PIXELS")]
    image_max_px: Option<u32>,

    /// JPEG 再エンコード品質（1〜100）。設定ファイルの `image.quality` より優先される。
    #[arg(long, value_name = "QUALITY")]
    image_quality: Option<u8>,

    /// XLSX 1シートあたりの最大データ行数（超過時に子 Section に分割）。
    /// 設定ファイルの `xlsx.max_rows` より優先される。
    #[arg(long, value_name = "ROWS")]
    xlsx_max_rows: Option<usize>,

    /// 実効設定（設定ファイル + CLI 引数の適用後）を JSON で標準出力に表示して終了する。
    /// 設定ファイルの内容確認や `docx2json.json` の雛形生成に利用できる。
    #[arg(long)]
    dump_config: bool,
}

fn main() {
    let cli = Cli::parse();

    match cli.command {
        None => {
            // 後方互換: サブコマンドなし → parse として動作
            run_parse(cli.args);
        }
        Some(Commands::Parse(args)) => {
            run_parse(args);
        }
        Some(Commands::ExtractCandidates(args)) => {
            if let Err(e) = commands::extract_candidates::run(args) {
                eprintln!("Error: {e:#}");
                std::process::exit(1);
            }
        }
        Some(Commands::InjectTags(args)) => {
            if let Err(e) = commands::inject_tags::run(args) {
                eprintln!("Error: {e:#}");
                std::process::exit(1);
            }
        }
        Some(Commands::InspectStyles(args)) => {
            if let Err(e) = commands::inspect_styles::run(args) {
                eprintln!("Error: {e:#}");
                std::process::exit(1);
            }
        }
        Some(Commands::Summarize(args)) => {
            if let Err(e) = commands::summarize::run(args) {
                eprintln!("Error: {e:#}");
                std::process::exit(1);
            }
        }
        Some(Commands::ToAsciidoc(args)) => {
            if let Err(e) = commands::to_asciidoc::run(args) {
                eprintln!("Error: {e:#}");
                std::process::exit(1);
            }
        }
    }
}

/// `parse` サブコマンド（またはサブコマンドなし時の後方互換）の実処理
fn run_parse(args: ParseArgs) {
    // 設定ファイルを読み込む
    let input_dir = if args.input.is_file() {
        args.input.parent().unwrap_or(&args.input).to_path_buf()
    } else {
        args.input.clone()
    };
    let mut cfg = config::Config::load(args.config.as_deref(), &input_dir);
    // CLI 引数で設定を上書き（設定ファイルより優先）
    if let Some(px) = args.image_max_px    { cfg.image.max_px = px; }
    if let Some(q)  = args.image_quality   { cfg.image.quality = q.clamp(1, 100); }
    if let Some(r)  = args.xlsx_max_rows   { cfg.xlsx.max_rows = r; }

    // --dump-config: 実効設定を JSON 出力して終了
    if args.dump_config {
        match serde_json::to_string_pretty(&cfg) {
            Ok(json) => println!("{}", json),
            Err(e) => {
                eprintln!("Error serializing config: {}", e);
                std::process::exit(1);
            }
        }
        return;
    }

    // 出力ディレクトリを作成
    if let Some(ref out) = args.output {
        if let Err(e) = std::fs::create_dir_all(out) {
            eprintln!("Error creating output directory: {}", e);
            std::process::exit(1);
        }
    }

    // 対象ファイルを収集
    let files = collect_files(&args.input);
    if files.is_empty() {
        eprintln!("No .docx, .xlsx, or .pptx files found in: {}", args.input.display());
        std::process::exit(1);
    }

    // ---- プログレスバーを初期化 ----
    let pb = ProgressBar::new(files.len() as u64);
    pb.set_style(
        ProgressStyle::with_template(
            "{spinner:.cyan} [{elapsed_precise}] [{bar:35.green/white}] {pos}/{len}  {wide_msg}",
        )
        .unwrap_or_else(|_| ProgressStyle::default_bar())
        .progress_chars("█▓░"),
    );
    pb.enable_steady_tick(std::time::Duration::from_millis(100));

    // Rayon で並列処理
    let results: Vec<_> = files
        .par_iter()
        .map(|path| {
            let filename = path
                .file_name()
                .map(|n| n.to_string_lossy().to_string())
                .unwrap_or_default();
            pb.set_message(filename);

            let result = parser::parse_file(path, &cfg)
                .and_then(|doc| {
                    if let Some(level) = args.split {
                        splitter::write_chunks(&doc, path, args.output.as_deref(), level)
                    } else {
                        output::write_json(&doc, path, args.output.as_deref(), &cfg)
                    }
                });

            pb.inc(1);
            (path, result)
        })
        .collect();

    pb.finish_and_clear();

    // ---- 結果サマリーを表示 ----
    let (ok, err): (Vec<_>, Vec<_>) = results.iter().partition(|(_, r)| r.is_ok());

    for (path, _) in ok.iter() {
        let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("?");
        let out_path = if let Some(ref dir) = args.output {
            dir.join(format!("{}.json", stem)).display().to_string()
        } else {
            path.with_extension("json").display().to_string()
        };
        println!("  ✓ {}", out_path);
    }

    println!(
        "\n完了: {} 件成功, {} 件失敗  (経過 {})",
        ok.len(),
        err.len(),
        format_elapsed(pb.elapsed()),
    );

    for (path, e) in err.iter() {
        if let Err(e) = e {
            eprintln!("\n  ✗ FAILED: {}", path.display());
            for (i, cause) in e.chain().enumerate() {
                eprintln!("    {}{}", "  ".repeat(i), cause);
            }
        }
    }

    if !err.is_empty() {
        std::process::exit(1);
    }
}

/// 経過時間を人間が読みやすい形式に変換する（例: "1m23s", "45s"）
fn format_elapsed(dur: std::time::Duration) -> String {
    let secs = dur.as_secs();
    if secs >= 60 {
        format!("{}m{}s", secs / 60, secs % 60)
    } else {
        format!("{}.{:02}s", secs, dur.subsec_millis() / 10)
    }
}

/// ディレクトリを再帰的にスキャンして .docx/.xlsx/.pptx ファイルを返す
fn collect_files(dir: &std::path::Path) -> Vec<PathBuf> {
    if dir.is_file() {
        let ext = dir.extension().and_then(|e| e.to_str()).unwrap_or("");
        if matches!(ext, "docx" | "xlsx" | "pptx") {
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
                if matches!(ext, "docx" | "xlsx" | "pptx") {
                    files.push(path);
                }
            }
        }
    }
    files
}
