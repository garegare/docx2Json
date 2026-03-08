use crate::models::Document;

/// AI変換を適用してDocumentを返す
/// `--features ai` でビルドした場合のみAPIを呼び出す
pub fn transform(doc: Document) -> Document {
    #[cfg(feature = "ai")]
    {
        transform_with_api(doc)
    }
    #[cfg(not(feature = "ai"))]
    {
        // 開発用モック: 変換なしでそのまま返す
        doc
    }
}

#[cfg(feature = "ai")]
fn transform_with_api(doc: Document) -> Document {
    use std::env;

    let api_key = match env::var("ANTHROPIC_API_KEY") {
        Ok(k) => k,
        Err(_) => {
            eprintln!("Warning: ANTHROPIC_API_KEY not set, skipping AI transformation");
            return doc;
        }
    };

    // TODO: Anthropic API への実際のリクエスト実装
    // ureq を使って各セクションを整形する
    let _ = api_key;
    doc
}
