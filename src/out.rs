//! Structured command output. Every command produces a `Report` which renders
//! either as grep-friendly text (default) or JSON (`--json`), mirroring
//! OfficeCLI's `path (type) "text" key=val ...` line format.

use serde::ser::SerializeMap;
use serde::{Serialize, Serializer};
use serde_json::Value;

fn pairs_as_object<S: Serializer>(
    pairs: &[(String, String)],
    serializer: S,
) -> Result<S::Ok, S::Error> {
    let mut map = serializer.serialize_map(Some(pairs.len()))?;
    for (k, v) in pairs {
        map.serialize_entry(k, v)?;
    }
    map.end()
}

/// One addressable document element, as returned by `get`/`add`/`view`.
#[derive(Debug, Clone, Serialize)]
pub struct NodeInfo {
    pub path: String,
    #[serde(rename = "type")]
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(
        skip_serializing_if = "Vec::is_empty",
        serialize_with = "pairs_as_object",
        default
    )]
    pub attributes: Vec<(String, String)>,
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub children: Vec<NodeInfo>,
}

impl NodeInfo {
    pub fn new(path: impl Into<String>, kind: impl Into<String>) -> NodeInfo {
        NodeInfo {
            path: path.into(),
            kind: kind.into(),
            text: None,
            attributes: Vec::new(),
            children: Vec::new(),
        }
    }

    pub fn attr(&mut self, key: impl Into<String>, value: impl Into<String>) -> &mut Self {
        self.attributes.push((key.into(), value.into()));
        self
    }

    fn render_lines(&self, out: &mut String, indent: usize) {
        let pad = "  ".repeat(indent);
        let mut line = format!("{pad}{} ({})", self.path, self.kind);
        if let Some(text) = &self.text {
            line.push_str(&format!(" {:?}", text));
        }
        for (k, v) in &self.attributes {
            line.push_str(&format!(" {k}={v}"));
        }
        out.push_str(&line);
        out.push('\n');
        for child in &self.children {
            child.render_lines(out, indent + 1);
        }
    }
}

#[derive(Debug)]
pub enum Report {
    /// Free-form text body (view modes); JSON form wraps it as {"text": ...}.
    Text(String),
    /// A list of nodes (get/add results).
    Nodes(Vec<NodeInfo>),
    /// Arbitrary structured data with a matching text rendering.
    Data { text: String, data: Value },
}

impl Report {
    pub fn message(text: impl Into<String>) -> Report {
        let text = text.into();
        Report::Data {
            data: serde_json::json!({ "message": text }),
            text,
        }
    }

    pub fn render(&self, json: bool) -> String {
        if json {
            let data = match self {
                Report::Text(t) => serde_json::json!({ "text": t }),
                Report::Nodes(nodes) => serde_json::json!({ "results": nodes }),
                Report::Data { data, .. } => data.clone(),
            };
            let envelope = serde_json::json!({ "ok": true, "data": data });
            serde_json::to_string_pretty(&envelope).unwrap()
        } else {
            match self {
                Report::Text(t) => t.trim_end().to_string(),
                Report::Nodes(nodes) => {
                    let mut out = String::new();
                    for n in nodes {
                        n.render_lines(&mut out, 0);
                    }
                    if nodes.is_empty() {
                        out.push_str("(no results)");
                    }
                    out.trim_end().to_string()
                }
                Report::Data { text, .. } => text.trim_end().to_string(),
            }
        }
    }
}

pub fn error_json(message: &str) -> String {
    serde_json::to_string_pretty(&serde_json::json!({
        "ok": false,
        "error": message,
    }))
    .unwrap()
}
