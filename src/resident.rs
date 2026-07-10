//! `officecli resident` — a long-lived command loop for agents that issue
//! many commands and want to skip per-command process startup.
//!
//! Protocol: each stdin line is one officecli command string (same form as
//! the MCP tool's `command` argument, shell-style quoting supported); each
//! response is exactly one compact JSON line on stdout:
//!   {"ok":true,"data":...}  or  {"ok":false,"error":"..."}
//! `exit` / `quit` (or EOF) ends the loop.

use anyhow::Result;
use serde_json::json;
use std::io::{BufRead, Write};

pub fn serve() -> Result<()> {
    let stdin = std::io::stdin();
    let stdout = std::io::stdout();
    for line in stdin.lock().lines() {
        let line = line?;
        let trimmed = line.trim();
        if trimmed.is_empty() {
            continue;
        }
        if matches!(trimmed, "exit" | "quit") {
            break;
        }
        let response = run_line(trimmed);
        let mut out = stdout.lock();
        writeln!(out, "{response}")?;
        out.flush()?;
    }
    Ok(())
}

fn run_line(command: &str) -> String {
    let envelope = match crate::mcp::shell_split(command) {
        Err(e) => json!({ "ok": false, "error": format!("{e:#}") }),
        Ok(words) => {
            let mut args = vec!["officecli".to_string()];
            args.extend(words);
            match crate::execute_args(args, false) {
                Ok((report, ok)) => json!({ "ok": ok, "data": report.data_value() }),
                Err(err) => json!({ "ok": false, "error": format!("{err:#}") }),
            }
        }
    };
    serde_json::to_string(&envelope).unwrap()
}
