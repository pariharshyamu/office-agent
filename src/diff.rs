//! `officecli diff a.docx b.docx` — a content-level comparison of two
//! same-format documents, built on the NodeInfo trees the handlers already
//! produce. Designed for agent verify loops: "did my edit do what I meant,
//! and nothing else?"

use crate::out::NodeInfo;

/// Compare two NodeInfo forests; returns change records as NodeInfo rows
/// (kind = added/removed/changed) so they render like any other result.
pub fn diff(a: &[NodeInfo], b: &[NodeInfo]) -> Vec<NodeInfo> {
    let mut changes = Vec::new();
    diff_level(a, b, &mut changes);
    changes
}

fn diff_level(a: &[NodeInfo], b: &[NodeInfo], changes: &mut Vec<NodeInfo>) {
    let max = a.len().max(b.len());
    for i in 0..max {
        match (a.get(i), b.get(i)) {
            (Some(x), None) => {
                let mut c = NodeInfo::new(&x.path, "removed");
                c.text = summarize(x);
                changes.push(c);
            }
            (None, Some(y)) => {
                let mut c = NodeInfo::new(&y.path, "added");
                c.text = summarize(y);
                changes.push(c);
            }
            (Some(x), Some(y)) => {
                if x.kind != y.kind {
                    let mut c = NodeInfo::new(&y.path, "changed");
                    c.attr("from-type", &x.kind);
                    c.attr("to-type", &y.kind);
                    c.text = summarize(y);
                    changes.push(c);
                    // Different element kinds: comparing children is noise.
                    continue;
                }
                let text_changed = x.text != y.text;
                let attr_changes = attr_diff(x, y);
                if text_changed || !attr_changes.is_empty() {
                    let mut c = NodeInfo::new(&y.path, "changed");
                    if text_changed {
                        c.attr("from", x.text.clone().unwrap_or_default());
                        c.text = Some(y.text.clone().unwrap_or_default());
                    }
                    for (k, from, to) in attr_changes {
                        c.attr(format!("{k}-from"), from);
                        c.attr(format!("{k}-to"), to);
                    }
                    changes.push(c);
                }
                diff_level(&x.children, &y.children, changes);
            }
            (None, None) => unreachable!(),
        }
    }
}

/// (name, from, to) for every attribute that differs.
fn attr_diff(x: &NodeInfo, y: &NodeInfo) -> Vec<(String, String, String)> {
    let mut out = Vec::new();
    let get = |n: &NodeInfo, key: &str| -> Option<String> {
        n.attributes
            .iter()
            .find(|(k, _)| k == key)
            .map(|(_, v)| v.clone())
    };
    let mut keys: Vec<&String> = x.attributes.iter().map(|(k, _)| k).collect();
    for (k, _) in &y.attributes {
        if !keys.contains(&k) {
            keys.push(k);
        }
    }
    for key in keys {
        let from = get(x, key);
        let to = get(y, key);
        if from != to {
            out.push((
                key.clone(),
                from.unwrap_or_else(|| "(none)".into()),
                to.unwrap_or_else(|| "(none)".into()),
            ));
        }
    }
    out
}

fn summarize(n: &NodeInfo) -> Option<String> {
    match &n.text {
        Some(t) if !t.is_empty() => Some(format!("{} {:?}", n.kind, truncate(t, 80))),
        _ => Some(n.kind.clone()),
    }
}

fn truncate(s: &str, max: usize) -> String {
    if s.chars().count() <= max {
        s.to_string()
    } else {
        let cut: String = s.chars().take(max).collect();
        format!("{cut}…")
    }
}
