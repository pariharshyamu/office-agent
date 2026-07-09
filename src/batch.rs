//! `officecli batch` — run multiple operations from JSON in one save cycle.
//!
//! Accepts the OfficeCLI batch item shape:
//!   [{"command":"set","path":"/Sheet1/A1","props":{"value":"Name"}},
//!    {"op":"add","parent":"/body","type":"paragraph","props":{"text":"hi"}}]

use anyhow::{bail, Context, Result};
use serde::Deserialize;
use serde_json::{json, Value};

use crate::handler::{Handler, Position};
use crate::out::Report;
use crate::props::Props;

#[derive(Debug, Deserialize)]
struct BatchItem {
    #[serde(alias = "op")]
    command: String,
    #[serde(default)]
    path: Option<String>,
    #[serde(default)]
    parent: Option<String>,
    #[serde(default, rename = "type")]
    typ: Option<String>,
    #[serde(default)]
    props: Option<serde_json::Map<String, Value>>,
    #[serde(default)]
    index: Option<usize>,
    #[serde(default)]
    before: Option<String>,
    #[serde(default)]
    after: Option<String>,
    #[serde(default)]
    find: Option<String>,
    #[serde(default)]
    replace: Option<String>,
    #[serde(default)]
    depth: Option<usize>,
    #[serde(default)]
    mode: Option<String>,
}

fn value_to_string(v: &Value) -> String {
    match v {
        Value::String(s) => s.clone(),
        other => other.to_string(),
    }
}

fn item_props(item: &BatchItem) -> Props {
    let pairs = item
        .props
        .as_ref()
        .map(|m| {
            m.iter()
                .map(|(k, v)| (k.clone(), value_to_string(v)))
                .collect()
        })
        .unwrap_or_default();
    Props::from_pairs(pairs)
}

fn item_position(item: &BatchItem) -> Position {
    if let Some(i) = item.index {
        Position::Index(i)
    } else if let Some(b) = &item.before {
        Position::Before(b.clone())
    } else if let Some(a) = &item.after {
        Position::After(a.clone())
    } else {
        Position::Append
    }
}

/// Run all operations. Returns the report plus whether every item succeeded.
pub fn run_batch(
    handler: &mut dyn Handler,
    source: &str,
    stop_on_error: bool,
) -> Result<(Report, bool)> {
    let items: Vec<BatchItem> = serde_json::from_str(source)
        .context("batch input must be a JSON array of operation objects")?;
    if items.is_empty() {
        bail!("batch input contains no operations");
    }
    let mut results = Vec::new();
    let mut all_ok = true;
    for (i, item) in items.iter().enumerate() {
        let outcome = run_item(handler, item);
        match outcome {
            Ok(report) => {
                results.push(json!({
                    "index": i,
                    "command": item.command,
                    "ok": true,
                    "result": serde_json::from_str::<Value>(&report.render(true))
                        .ok()
                        .and_then(|v| v.get("data").cloned())
                        .unwrap_or(Value::Null),
                }));
            }
            Err(err) => {
                all_ok = false;
                results.push(json!({
                    "index": i,
                    "command": item.command,
                    "ok": false,
                    "error": format!("{err:#}"),
                }));
                if stop_on_error {
                    break;
                }
            }
        }
    }
    let succeeded = results.iter().filter(|r| r["ok"] == true).count();
    let text = results
        .iter()
        .map(|r| {
            if r["ok"] == true {
                format!("[{}] {} ok", r["index"], r["command"].as_str().unwrap_or("?"))
            } else {
                format!(
                    "[{}] {} FAILED: {}",
                    r["index"],
                    r["command"].as_str().unwrap_or("?"),
                    r["error"].as_str().unwrap_or("?")
                )
            }
        })
        .collect::<Vec<_>>()
        .join("\n");
    let report = Report::Data {
        text: format!("{text}\n{succeeded}/{} operations succeeded", items.len()),
        data: json!({
            "total": items.len(),
            "succeeded": succeeded,
            "results": results,
        }),
    };
    Ok((report, all_ok))
}

fn run_item(handler: &mut dyn Handler, item: &BatchItem) -> Result<Report> {
    let props = item_props(item);
    match item.command.as_str() {
        "get" => handler.get(
            item.path.as_deref().context("get needs 'path'")?,
            item.depth.unwrap_or(0),
        ),
        "view" => handler.view(item.mode.as_deref().unwrap_or("outline")),
        "add" => handler.add(
            item.parent
                .as_deref()
                .or(item.path.as_deref())
                .context("add needs 'parent' (or 'path')")?,
            item.typ.as_deref().context("add needs 'type'")?,
            &props,
            &item_position(item),
        ),
        "set" => {
            // find/replace may come as top-level fields or props.
            let find_owned = item
                .find
                .clone()
                .or_else(|| props.get("find").map(|s| s.to_string()));
            let replace_owned = item
                .replace
                .clone()
                .or_else(|| props.get("replace").map(|s| s.to_string()));
            handler.set(
                item.path.as_deref().context("set needs 'path'")?,
                &props,
                find_owned.as_deref(),
                replace_owned.as_deref(),
            )
        }
        "remove" => handler.remove(item.path.as_deref().context("remove needs 'path'")?),
        "validate" => handler.validate(),
        other => bail!("unsupported batch command '{other}' (get/view/add/set/remove/validate)"),
    }
}
