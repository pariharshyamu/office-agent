//! `officecli mcp` — a Model Context Protocol server over stdio.
//!
//! Mirrors upstream OfficeCLI's MCP design: a single `officecli` tool whose
//! one parameter is the CLI command string, passed through verbatim. Uses
//! newline-delimited JSON-RPC 2.0 (the MCP stdio transport).

use anyhow::{bail, Result};
use serde_json::{json, Value};
use std::io::{BufRead, Write};

const PROTOCOL_VERSION: &str = "2024-11-05";

pub fn serve() -> Result<()> {
    let stdin = std::io::stdin();
    let stdout = std::io::stdout();
    for line in stdin.lock().lines() {
        let line = line?;
        if line.trim().is_empty() {
            continue;
        }
        let msg: Value = match serde_json::from_str(&line) {
            Ok(v) => v,
            Err(e) => {
                let mut out = stdout.lock();
                writeln!(
                    out,
                    "{}",
                    json!({
                        "jsonrpc": "2.0", "id": null,
                        "error": { "code": -32700, "message": format!("parse error: {e}") },
                    })
                )?;
                continue;
            }
        };
        let id = msg.get("id").cloned();
        let method = msg.get("method").and_then(|m| m.as_str()).unwrap_or("");
        // Notifications (no id) get no response.
        if id.is_none() || id == Some(Value::Null) {
            continue;
        }
        let response = match handle(method, msg.get("params")) {
            Ok(result) => json!({ "jsonrpc": "2.0", "id": id, "result": result }),
            Err(err) => json!({
                "jsonrpc": "2.0", "id": id,
                "error": { "code": -32603, "message": format!("{err:#}") },
            }),
        };
        let mut out = stdout.lock();
        writeln!(out, "{response}")?;
        out.flush()?;
    }
    Ok(())
}

fn handle(method: &str, params: Option<&Value>) -> Result<Value> {
    match method {
        "initialize" => Ok(json!({
            "protocolVersion": PROTOCOL_VERSION,
            "capabilities": { "tools": {} },
            "serverInfo": {
                "name": "officecli",
                "version": env!("CARGO_PKG_VERSION"),
            },
        })),
        "ping" => Ok(json!({})),
        "tools/list" => Ok(json!({
            "tools": [{
                "name": "officecli",
                "description": "Run an officecli command to read, create, or edit Office documents (.docx/.xlsx/.pptx). Pass the full command line without the leading 'officecli', e.g. \"create deck.pptx\" or \"add deck.pptx / --type slide --prop title='Q4'\". Run \"help\" for the command guide.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "command": {
                            "type": "string",
                            "description": "The officecli command line (shell-style quoting supported)",
                        },
                    },
                    "required": ["command"],
                },
            }],
        })),
        "tools/call" => {
            let params = params.ok_or_else(|| anyhow::anyhow!("tools/call needs params"))?;
            let name = params.get("name").and_then(|n| n.as_str()).unwrap_or("");
            if name != "officecli" {
                bail!("unknown tool '{name}'");
            }
            let command = params
                .get("arguments")
                .and_then(|a| a.get("command"))
                .and_then(|c| c.as_str())
                .ok_or_else(|| anyhow::anyhow!("missing 'command' argument"))?;
            let (text, is_error) = run_command_line(command);
            Ok(json!({
                "content": [{ "type": "text", "text": text }],
                "isError": is_error,
            }))
        }
        other => bail!("method '{other}' not supported"),
    }
}

/// Execute one CLI command line in-process; never panics the server.
fn run_command_line(command: &str) -> (String, bool) {
    let words = match shell_split(command) {
        Ok(w) => w,
        Err(e) => return (format!("error: {e}"), true),
    };
    let mut args = vec!["officecli".to_string()];
    args.extend(words);
    match crate::execute_args(args, false) {
        Ok((report, all_ok)) => (report.render(false), !all_ok),
        Err(err) => (format!("error: {err:#}"), true),
    }
}

/// Minimal shell-style splitter: whitespace-separated, single/double quotes
/// group words, backslash escapes inside double quotes and bare words.
fn shell_split(input: &str) -> Result<Vec<String>> {
    let mut words = Vec::new();
    let mut cur = String::new();
    let mut in_word = false;
    let mut chars = input.chars().peekable();
    while let Some(ch) = chars.next() {
        match ch {
            c if c.is_whitespace() => {
                if in_word {
                    words.push(std::mem::take(&mut cur));
                    in_word = false;
                }
            }
            '\'' => {
                in_word = true;
                for c in chars.by_ref() {
                    if c == '\'' {
                        break;
                    }
                    cur.push(c);
                }
            }
            '"' => {
                in_word = true;
                while let Some(c) = chars.next() {
                    match c {
                        '"' => break,
                        '\\' => {
                            if let Some(&next) = chars.peek() {
                                if matches!(next, '"' | '\\') {
                                    cur.push(next);
                                    chars.next();
                                } else {
                                    cur.push('\\');
                                }
                            }
                        }
                        c => cur.push(c),
                    }
                }
            }
            '\\' => {
                in_word = true;
                if let Some(c) = chars.next() {
                    cur.push(c);
                }
            }
            c => {
                in_word = true;
                cur.push(c);
            }
        }
    }
    if in_word {
        words.push(cur);
    }
    if words.is_empty() {
        bail!("empty command");
    }
    Ok(words)
}

#[cfg(test)]
mod tests {
    use super::shell_split;

    #[test]
    fn splits_quoted_words() {
        assert_eq!(
            shell_split(r#"add deck.pptx '/slide[1]' --prop text="Hello world""#).unwrap(),
            vec!["add", "deck.pptx", "/slide[1]", "--prop", "text=Hello world"]
        );
        assert_eq!(
            shell_split(r#"set f.docx / --find "a \"b\"""#).unwrap(),
            vec!["set", "f.docx", "/", "--find", "a \"b\""]
        );
    }
}
