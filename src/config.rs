use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::Path;

/// 変換設定（docx2json.json から読み込む）
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Config {
    /// 見出しとして扱うスタイル名とそのレベル
    /// 例: {"Heading1": 1, "Heading2": 2, "見出し1": 1}
    #[serde(default)]
    pub heading_styles: HashMap<String, usize>,

    /// w:pPr > w:rPr に w:u（下線）がある段落を見出し（level 1）として扱うか
    #[serde(default = "default_true")]
    pub ppr_underline_as_heading: bool,

    /// ラン（w:r > w:rPr）に w:u（下線）がある段落を見出し（level 1）として扱うか
    /// ppr_underline_as_heading と同時に使う。直接書式設定で見出しを表現する文書向け。
    #[serde(default = "default_true")]
    pub run_underline_as_heading: bool,

    /// 画像の最大辺長（ピクセル）。この値を超える辺がある場合にリサイズする。
    /// 0 の場合はリサイズ・圧縮を行わない（デフォルト）。
    #[serde(default)]
    pub image_max_px: u32,

    /// JPEG 再エンコード時の品質（1〜100）。image_max_px > 0 のときのみ有効。
    /// デフォルト 80。
    #[serde(default = "default_image_quality")]
    pub image_quality: u8,
}

fn default_true() -> bool {
    true
}

fn default_image_quality() -> u8 {
    80
}

impl Default for Config {
    fn default() -> Self {
        // デフォルト: 標準的な英語・日本語スタイル名
        let mut heading_styles = HashMap::new();
        for (name, level) in [
            ("Heading1", 1usize), ("Heading2", 2), ("Heading3", 3),
            ("heading1", 1), ("heading2", 2), ("heading3", 3),
            ("見出し1", 1), ("見出し2", 2), ("見出し3", 3),
            ("1", 1), ("2", 2), ("3", 3),  // 数値IDスタイル
        ] {
            heading_styles.insert(name.to_string(), level);
        }
        Self {
            heading_styles,
            ppr_underline_as_heading: true,
            run_underline_as_heading: false,  // デフォルトはオフ（誤検出防止）
            image_max_px: 0,   // デフォルトはリサイズなし
            image_quality: 80, // JPEG品質デフォルト
        }
    }
}

impl Config {
    /// 設定ファイルを探して読み込む。見つからない場合はデフォルトを返す。
    /// 探索順: `--config` で指定されたパス → 入力ディレクトリ内の docx2json.json
    pub fn load(config_path: Option<&Path>, input_dir: &Path) -> Self {
        let candidates = [
            config_path.map(|p| p.to_path_buf()),
            Some(input_dir.join("docx2json.json")),
            Some(std::env::current_dir().unwrap_or_default().join("docx2json.json")),
        ];

        for path in candidates.iter().flatten() {
            if path.exists() {
                match std::fs::read_to_string(path) {
                    Ok(content) => match serde_json::from_str::<Config>(&content) {
                        Ok(cfg) => {
                            eprintln!("Config loaded: {}", path.display());
                            return cfg;
                        }
                        Err(e) => {
                            eprintln!("Warning: failed to parse config {}: {}", path.display(), e);
                        }
                    },
                    Err(_) => {}
                }
            }
        }
        Config::default()
    }

    /// スタイル名からレベルを返す
    pub fn heading_level_for_style(&self, style: &str) -> Option<usize> {
        self.heading_styles.get(style).copied()
    }
}
