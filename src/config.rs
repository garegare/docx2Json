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
        // スタイル名は normalize_style_name() で正規化して格納する（#5）
        let mut heading_styles = HashMap::new();
        for (name, level) in [
            ("Heading1", 1usize), ("Heading2", 2), ("Heading3", 3),
            ("heading1", 1), ("heading2", 2), ("heading3", 3),
            ("見出し1", 1), ("見出し2", 2), ("見出し3", 3),
            // 全角数字バリアント（例: 「見出し１」→ 正規化で「見出し1」に統合済み）
            ("見出し１", 1), ("見出し２", 2), ("見出し３", 3),
            ("1", 1), ("2", 2), ("3", 3),  // 数値IDスタイル
        ] {
            heading_styles.insert(normalize_style_name(name), level);
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
    ///
    /// 探索順（#5 暗黙的ロード対応）:
    ///   1. `--config` で指定されたパス
    ///   2. 入力ディレクトリ内の `docx2json.json`
    ///   3. カレントディレクトリの `docx2json.json`
    ///   4. 実行バイナリと同じディレクトリの `docx2json.json`
    pub fn load(config_path: Option<&Path>, input_dir: &Path) -> Self {
        let exe_dir = std::env::current_exe()
            .ok()
            .and_then(|p| p.parent().map(|d| d.join("docx2json.json")));

        let candidates = [
            config_path.map(|p| p.to_path_buf()),
            Some(input_dir.join("docx2json.json")),
            Some(std::env::current_dir().unwrap_or_default().join("docx2json.json")),
            exe_dir,
        ];

        for path in candidates.iter().flatten() {
            if path.exists() {
                match std::fs::read_to_string(path) {
                    Ok(content) => match serde_json::from_str::<Config>(&content) {
                        Ok(cfg) => {
                            eprintln!("Config loaded: {}", path.display());
                            // heading_styles のキーを正規化して返す（#5）
                            let normalized_styles = cfg.heading_styles
                                .into_iter()
                                .map(|(k, v)| (normalize_style_name(&k), v))
                                .collect();
                            return Config {
                                heading_styles: normalized_styles,
                                ..cfg
                            };
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

    /// スタイル名からレベルを返す（#5 正規化マッチング）
    ///
    /// DOCX に記録されるスタイル名は環境によって全角・半角が混在する場合がある。
    /// 例: 「見出し１」と「見出し1」は同一スタイルとして扱う。
    pub fn heading_level_for_style(&self, style: &str) -> Option<usize> {
        let normalized = normalize_style_name(style);
        self.heading_styles.get(&normalized).copied()
    }
}

/// スタイル名を正規化する（#5 全角・半角の揺れを吸収）
///
/// 全角英数字（ＡＢＣ、０１２ など）を半角（ABC, 012 など）に変換する。
/// これにより「見出し１」「見出し1」などの表記揺れを統一して検索できる。
pub fn normalize_style_name(s: &str) -> String {
    s.chars()
        .map(|c| match c {
            // 全角数字 → 半角数字
            '０'..='９' => char::from_u32(c as u32 - '０' as u32 + '0' as u32).unwrap_or(c),
            // 全角大文字英字 → 半角大文字
            'Ａ'..='Ｚ' => char::from_u32(c as u32 - 'Ａ' as u32 + 'A' as u32).unwrap_or(c),
            // 全角小文字英字 → 半角小文字
            'ａ'..='ｚ' => char::from_u32(c as u32 - 'ａ' as u32 + 'a' as u32).unwrap_or(c),
            _ => c,
        })
        .collect()
}
