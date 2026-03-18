use regex::Regex;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::Path;

use crate::models::SemanticRole;

// ── デフォルト値ヘルパー ────────────────────────────────────────────────

fn default_true() -> bool {
    true
}

fn default_image_quality() -> u8 {
    80
}

fn default_heading_cell_ratio() -> f32 {
    0.5
}

// ── サブ設定構造体 ──────────────────────────────────────────────────────

/// 画像処理設定
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ImageConfig {
    /// 画像の最大辺長（px）。超過する辺がある場合にリサイズして JPEG 再エンコードする。
    /// 0 の場合はリサイズ・圧縮を行わない（デフォルト）。
    /// CLI: `--image-max-px`
    #[serde(default)]
    pub max_px: u32,

    /// JPEG 再エンコード時の品質（1〜100）。`max_px > 0` のときのみ有効。
    /// デフォルト 80。CLI: `--image-quality`
    #[serde(default = "default_image_quality")]
    pub quality: u8,
}

impl Default for ImageConfig {
    fn default() -> Self {
        Self {
            max_px: 0,
            quality: 80,
        }
    }
}

/// DOCX パーサー設定
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DocxConfig {
    /// 見出しとして扱うスタイル名とそのレベル（#12 前方一致・正規表現に対応）
    ///
    /// キーの記法:
    ///   - 通常文字列       → 完全一致（正規化後）例: `"Heading1": 1`
    ///   - `"prefix:<文字列>"` → 前方一致              例: `"prefix:My Heading": 1`
    ///   - `"regex:<パターン>"` → 正規表現マッチ        例: `"regex:^Heading\\d+$": 1`
    ///
    /// 優先順位: 完全一致 > 前方一致 > 正規表現（設定ファイル内の順序）
    #[serde(default)]
    pub heading_styles: HashMap<String, usize>,

    /// w:pPr > w:rPr に w:u（下線）がある段落を見出し（level 1）として扱うか（デフォルト: true）
    #[serde(default = "default_true")]
    pub ppr_underline_as_heading: bool,

    /// ラン（w:r > w:rPr）に w:u（下線）がある段落を見出し（level 1）として扱うか（デフォルト: false）
    /// 直接書式設定で見出しを表現する文書向け。誤検出防止のためデフォルトはオフ。
    #[serde(default)]
    pub run_underline_as_heading: bool,

    /// カスタム意味的役割マッピング: スタイル名（正規化済み） → SemanticRole
    ///
    /// 組み込みルール（単語境界マッチ）より優先される。
    /// 値は SemanticRole の snake_case 文字列: "note" / "warning" / "tip" /
    /// "code_block" / "quote" / "bullet_list" / "ordered_list"
    ///
    /// 例: `{ "MyCustomNote": "note", "SpecialAlert": "warning" }`
    #[serde(default)]
    pub semantic_role_styles: HashMap<String, SemanticRole>,
}

impl Default for DocxConfig {
    fn default() -> Self {
        let mut heading_styles = HashMap::new();
        for (name, level) in [
            ("Heading1", 1usize), ("Heading2", 2), ("Heading3", 3),
            ("Heading4", 4), ("Heading5", 5), ("Heading6", 6),
            ("heading1", 1), ("heading2", 2), ("heading3", 3),
            ("heading4", 4), ("heading5", 5), ("heading6", 6),
            ("見出し1", 1), ("見出し2", 2), ("見出し3", 3),
            ("見出し4", 4), ("見出し5", 5), ("見出し6", 6),
            // 全角数字バリアント（正規化で「見出し1」に統合済み）
            ("見出し１", 1), ("見出し２", 2), ("見出し３", 3),
            ("見出し４", 4), ("見出し５", 5), ("見出し６", 6),
            ("1", 1), ("2", 2), ("3", 3), // 数値IDスタイル
            ("4", 4), ("5", 5), ("6", 6),
        ] {
            heading_styles.insert(normalize_style_name(name), level);
        }
        Self {
            heading_styles,
            ppr_underline_as_heading: true,
            run_underline_as_heading: false,
            semantic_role_styles: HashMap::new(),
        }
    }
}

/// XLSX パーサー設定
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct XlsxConfig {
    /// 1シートあたりの最大データ行数（ヘッダー行を除く）。
    /// 超過した場合、ヘッダー行を引き継いだ子 Section に分割する。
    /// 0 = 制限なし（デフォルト）。CLI: `--xlsx-max-rows`
    #[serde(default)]
    pub max_rows: usize,

    /// 神エクセル対応の書式ベース見出し判定設定（#10）。
    /// `null` または省略時は従来モード（先頭行ヘッダー、xlsx.rs を使用）。
    /// `enabled: true` のときのみ xlsx_advanced パーサーに切り替わる。
    #[serde(default)]
    pub heading: Option<XlsxHeadingConfig>,
}


/// XLSX 書式ベース見出し判定の詳細設定（#10 神エクセル対応）
///
/// `xlsx.heading.enabled: true` のときのみ有効。
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XlsxHeadingConfig {
    /// 書式ベース見出し判定を有効にするか（デフォルト: false）
    #[serde(default)]
    pub enabled: bool,

    /// 太字セルを見出し条件にするか（デフォルト: true）
    #[serde(default = "default_true")]
    pub detect_bold: bool,

    /// 非白・非透明の背景色セルを見出し条件にするか（デフォルト: true）
    #[serde(default = "default_true")]
    pub detect_fill: bool,

    /// 見出し判定の最小フォントサイズ（pt）。0.0 = 無効（デフォルト: 0.0）
    #[serde(default)]
    pub heading_font_size_threshold: f32,

    /// 行内で「見出し書式」セルが占める割合の閾値（0.0〜1.0）。
    /// この割合以上なら行全体を見出し行と判定する。デフォルト: 0.5
    #[serde(default = "default_heading_cell_ratio")]
    pub heading_cell_ratio: f32,
}

/// JSON 出力設定
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OutputConfig {
    /// body_text フィールドを JSON に含めるか（デフォルト: false）。
    /// true にすると後方互換用フラットテキストも出力される。
    #[serde(default)]
    pub include_body_text: bool,

    /// 画像データを Base64 エンコードして assets.data に含めるか（デフォルト: true）。
    /// false にするとデータフィールドを省略し JSON サイズを削減できる。
    #[serde(default = "default_true")]
    pub include_base64: bool,
}

impl Default for OutputConfig {
    fn default() -> Self {
        Self {
            include_body_text: false,
            include_base64: true,
        }
    }
}

// ── メイン Config 構造体 ────────────────────────────────────────────────

/// 変換設定（docx2json.json から読み込む）
///
/// 設定ファイルは `docx2json.json` として以下の順で自動探索される:
///   1. `--config` で指定されたパス
///   2. 入力ディレクトリ内の `docx2json.json`
///   3. カレントディレクトリの `docx2json.json`
///   4. 実行バイナリと同じディレクトリの `docx2json.json`
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Config {
    /// 画像処理設定
    #[serde(default)]
    pub image: ImageConfig,

    /// DOCX パーサー設定
    #[serde(default)]
    pub docx: DocxConfig,

    /// XLSX パーサー設定
    #[serde(default)]
    pub xlsx: XlsxConfig,

    /// JSON 出力設定
    #[serde(default)]
    pub output: OutputConfig,

    /// ロード時にコンパイル済みのマッチングルール群（serde には含まない）
    #[serde(skip)]
    heading_rules: Vec<HeadingRule>,
}

/// 見出しスタイルのマッチングルール（#12）
///
/// `docx.heading_styles` のキー記法に基づいてロード時に生成される。
#[derive(Debug, Clone)]
enum HeadingRule {
    /// 完全一致（正規化済み）
    Exact(String, usize),
    /// 前方一致（正規化済み）: キーが `"prefix:<str>"` の形式
    Prefix(String, usize),
    /// 正規表現マッチ: キーが `"regex:<pattern>"` の形式
    Regex(Regex, usize),
}

impl Default for Config {
    fn default() -> Self {
        let docx = DocxConfig::default();
        let rules = compile_heading_rules(&docx.heading_styles);
        Self {
            image: ImageConfig::default(),
            docx,
            xlsx: XlsxConfig::default(),
            output: OutputConfig::default(),
            heading_rules: rules,
        }
    }
}

impl Config {
    /// 設定ファイルを探して読み込む。見つからない場合はデフォルトを返す。
    ///
    /// 探索順:
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
                if let Ok(content) = std::fs::read_to_string(path) {
                    match serde_json::from_str::<Config>(&content) {
                        Ok(mut cfg) => {
                            eprintln!("Config loaded: {}", path.display());
                            // heading_styles のキーを正規化
                            cfg.docx.heading_styles = cfg.docx.heading_styles
                                .into_iter()
                                .map(|(k, v)| (normalize_style_key(&k), v))
                                .collect();
                            // semantic_role_styles のキーを正規化
                            cfg.docx.semantic_role_styles = cfg.docx.semantic_role_styles
                                .into_iter()
                                .map(|(k, v)| (normalize_style_name(&k), v))
                                .collect();
                            // マッチングルールをコンパイル
                            cfg.heading_rules = compile_heading_rules(&cfg.docx.heading_styles);
                            return cfg;
                        }
                        Err(e) => {
                            eprintln!("Warning: failed to parse config {}: {}", path.display(), e);
                        }
                    }
                }
            }
        }
        Config::default()
    }

    /// スタイル名からレベルを返す（#12 前方一致・正規表現マッチ対応）
    ///
    /// マッチ優先順位（`heading_rules` の構築順）:
    ///   1. 完全一致（正規化済み）
    ///   2. 前方一致（`prefix:` キー）
    ///   3. 正規表現（`regex:` キー）
    pub fn heading_level_for_style(&self, style: &str) -> Option<usize> {
        let normalized = normalize_style_name(style);
        for rule in &self.heading_rules {
            match rule {
                HeadingRule::Exact(key, level) => {
                    if *key == normalized {
                        return Some(*level);
                    }
                }
                HeadingRule::Prefix(prefix, level) => {
                    if normalized.starts_with(prefix.as_str()) {
                        return Some(*level);
                    }
                }
                HeadingRule::Regex(re, level) => {
                    if re.is_match(&normalized) {
                        return Some(*level);
                    }
                }
            }
        }
        None
    }
}

/// `docx.heading_styles` の HashMap から `HeadingRule` のリストを構築する（#12）
///
/// 優先度順: Exact → Prefix → Regex の順にソートして返す。
/// 正規表現が無効な場合は警告を出して当該ルールをスキップする。
fn compile_heading_rules(styles: &HashMap<String, usize>) -> Vec<HeadingRule> {
    let mut exact   = Vec::new();
    let mut prefix  = Vec::new();
    let mut patterns = Vec::new();

    for (key, &level) in styles {
        if let Some(pat) = key.strip_prefix("regex:") {
            match Regex::new(pat) {
                Ok(re) => patterns.push(HeadingRule::Regex(re, level)),
                Err(e) => eprintln!("Warning: invalid regex pattern '{pat}': {e}"),
            }
        } else if let Some(pfx) = key.strip_prefix("prefix:") {
            prefix.push(HeadingRule::Prefix(normalize_style_name(pfx), level));
        } else {
            exact.push(HeadingRule::Exact(key.clone(), level));
        }
    }

    // 優先度順に結合: 完全一致 → 前方一致 → 正規表現
    exact.extend(prefix);
    exact.extend(patterns);
    exact
}

/// `heading_styles` のキーを正規化する。
///
/// `"prefix:"` / `"regex:"` プレフィックスは保持し、それ以外の部分のみ
/// `normalize_style_name()` で全角→半角変換する。
fn normalize_style_key(key: &str) -> String {
    if let Some(rest) = key.strip_prefix("regex:") {
        format!("regex:{}", rest)
    } else if let Some(rest) = key.strip_prefix("prefix:") {
        format!("prefix:{}", normalize_style_name(rest))
    } else {
        normalize_style_name(key)
    }
}

/// スタイル名を正規化する（全角英数字 → 半角）
pub fn normalize_style_name(s: &str) -> String {
    s.chars()
        .map(|c| match c {
            '０'..='９' => char::from_u32(c as u32 - '０' as u32 + '0' as u32).unwrap_or(c),
            'Ａ'..='Ｚ' => char::from_u32(c as u32 - 'Ａ' as u32 + 'A' as u32).unwrap_or(c),
            'ａ'..='ｚ' => char::from_u32(c as u32 - 'ａ' as u32 + 'a' as u32).unwrap_or(c),
            _ => c,
        })
        .collect()
}
