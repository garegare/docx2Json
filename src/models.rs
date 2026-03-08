use base64::Engine;
use serde::{Deserialize, Deserializer, Serialize, Serializer};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Document {
    pub title: String,
    pub sections: Vec<Section>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Section {
    /// このセクション自身の見出しを含む、ルートからの見出しパスリスト。
    /// RAG でチャンク分割した後も文書内の位置を保持するために使用する。
    /// 例: ["第1章 導入", "1.1 背景", "1.1.1 詳細"]
    pub context_path: Vec<String>,
    pub heading: String,
    pub body_text: String,
    pub assets: Vec<Asset>,
    pub children: Vec<Section>,
}

/// 画像等のバイナリアセット。
/// 内部では `Vec<u8>` として保持し、JSON 出力時にのみ Base64 エンコードする
/// （Lazy Serialization）。Base64 文字列は元データより約 33% 大きいため、
/// メモリ上はバイナリで保持することで並列処理時のメモリ使用量を抑える。
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Asset {
    #[serde(rename = "type")]
    pub asset_type: String,
    pub title: String,
    /// バイナリデータ。JSON では Base64 文字列として出力される。
    #[serde(serialize_with = "serialize_as_base64", deserialize_with = "deserialize_from_base64")]
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
