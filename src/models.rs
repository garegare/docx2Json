use serde::{Deserialize, Serialize};

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

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Asset {
    #[serde(rename = "type")]
    pub asset_type: String,
    pub title: String,
    pub data: String,
}

