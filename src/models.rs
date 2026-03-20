use base64::Engine;
use serde::{Deserialize, Deserializer, Serialize, Serializer};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Document {
    pub title: String,
    pub sections: Vec<Section>,
}

/// セクションに付与されるメタデータ。AI タグ等の追加情報を格納する。
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct SectionMetadata {
    /// AI が付与したタグの配列。初期値は空配列。
    #[serde(default)]
    pub ai_tags: Vec<String>,
}

/// 段落・テーブル・アセット参照の意味的役割
///
/// Note: `Heading` は Section ツリーで表現するため Element には現れない。
/// `InlineCode` はパーサーが未対応のため将来拡張用として予約済み。
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "snake_case")]
pub enum SemanticRole {
    Note,
    Warning,
    Tip,
    CodeBlock,
    Quote,
    BulletList,
    OrderedList,
}

/// 要素に付与されるメタデータ（スタイル・アラインメント・役割・アンカー等）
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ElementMetadata {
    /// Word スタイル名（正規化済み）。例: "Heading1", "Normal"
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
    /// Word スタイル名（生値）。style と同じ値だが将来的なカスタムスタイル対応用
    #[serde(skip_serializing_if = "Option::is_none")]
    pub raw_style: Option<String>,
    /// 段落の水平配置。"left" / "center" / "right" / "both"（両端揃え）
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<String>,
    /// 段落のアウトラインレベル（1 = 最上位見出し）。w:outlineLvl から 1-based に変換
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u32>,
    /// 意味的役割（SemanticRole）。スタイル名・リスト属性から自動判定
    #[serde(skip_serializing_if = "Option::is_none")]
    pub role: Option<SemanticRole>,
    /// w:bookmarkStart の name 属性から取得したアンカー ID
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_id: Option<String>,
    /// asset_ref 要素のキャプション（画像の代替テキストまたは名前）
    #[serde(skip_serializing_if = "Option::is_none")]
    pub caption: Option<String>,
}

/// セクション内の構造化要素
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum Element {
    /// テキスト段落
    Paragraph {
        text: String,
        #[serde(default)]
        metadata: ElementMetadata,
    },
    /// テーブル（Markdown 互換の行列表現）
    Table {
        #[serde(default)]
        metadata: ElementMetadata,
        rows: Vec<Vec<String>>,
        /// セル結合情報: [(row, col, rowspan, colspan)] ブロックローカル 0-based。
        /// 結合なしの場合は省略される（後方互換）。
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        merges: Vec<(usize, usize, usize, usize)>,
    },
    /// 画像等のアセット参照（assets 配列の id と対応）
    AssetRef {
        asset_id: String,
        #[serde(default)]
        metadata: ElementMetadata,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Section {
    /// 文書タイトル + context_path を連結した文字列の FNV-1a 16進数ハッシュ。
    /// 実行間で安定した ID として使用する（追加クレート不要）。
    #[serde(default)]
    pub id: String,
    /// このセクション自身の見出しを含む、ルートからの見出しパスリスト。
    /// RAG でチャンク分割した後も文書内の位置を保持するために使用する。
    /// 例: ["第1章 導入", "1.1 背景", "1.1.1 詳細"]
    pub context_path: Vec<String>,
    pub heading: String,
    /// 後方互換用フラットテキスト（Markdown 形式）。
    /// 新規コードは elements を参照することを推奨する。
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub body_text: String,
    /// 構造化要素配列（paragraph / table / asset_ref）。
    /// body_text と並行して生成される。
    #[serde(default)]
    pub elements: Vec<Element>,
    pub assets: Vec<Asset>,
    pub children: Vec<Section>,
    /// AI タグ等のメタデータ。初期値: { ai_tags: [] }
    #[serde(default)]
    pub metadata: SectionMetadata,
}


/// 画像等のバイナリアセット。
/// 内部では `Vec<u8>` として保持し、JSON 出力時にのみ Base64 エンコードする
/// （Lazy Serialization）。Base64 文字列は元データより約 33% 大きいため、
/// メモリ上はバイナリで保持することで並列処理時のメモリ使用量を抑える。
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Asset {
    #[serde(rename = "type")]
    pub asset_type: String,
    /// リレーションシップ ID（rId）。element の asset_ref.asset_id と対応する。
    /// 旧形式の JSON（id フィールドなし）との後方互換のため default を許容する。
    /// id を持たないアセット（PPTX 画像等）は None。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    pub title: String,
    /// バイナリデータ。JSON では Base64 文字列として出力される。
    #[serde(
        default,
        serialize_with = "serialize_as_base64",
        skip_serializing_if = "Vec::is_empty",
        deserialize_with = "deserialize_from_base64"
    )]
    pub data: Vec<u8>,
}

/// `Vec<u8>` を Base64 文字列としてシリアライズ（JSON 出力用）
fn serialize_as_base64<S>(bytes: &Vec<u8>, serializer: S) -> Result<S::Ok, S::Error>
where
    S: Serializer,
{
    serializer.serialize_str(
        &base64::engine::general_purpose::STANDARD.encode(bytes),
    )
}

/// Base64 文字列を `Vec<u8>` としてデシリアライズ（既存 JSON 読み込み互換）
fn deserialize_from_base64<'de, D>(deserializer: D) -> Result<Vec<u8>, D::Error>
where
    D: Deserializer<'de>,
{
    let s = String::deserialize(deserializer)?;
    base64::engine::general_purpose::STANDARD
        .decode(s)
        .map_err(serde::de::Error::custom)
}
