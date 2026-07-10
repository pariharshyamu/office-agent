//! `officecli watch` — a live HTML preview: a tiny single-threaded HTTP
//! server that re-renders the document's `view html` on every request and
//! auto-reloads the browser when the file changes on disk (mtime polling).

use anyhow::{Context, Result};
use std::io::{BufRead, BufReader, Write};
use std::net::{TcpListener, TcpStream};
use std::path::Path;

const RELOAD_SCRIPT: &str = r#"<script>
(function () {
  let v = null;
  setInterval(async () => {
    try {
      const r = await fetch('/__v', { cache: 'no-store' });
      const t = await r.text();
      if (v === null) v = t;
      else if (t !== v) location.reload();
    } catch (e) { /* server restarting; keep polling */ }
  }, 800);
})();
</script>"#;

pub fn serve(file: &Path, port: u16) -> Result<()> {
    let listener = TcpListener::bind(("127.0.0.1", port))
        .with_context(|| format!("cannot listen on 127.0.0.1:{port}"))?;
    let port = listener.local_addr()?.port();
    eprintln!(
        "watching {} — open http://127.0.0.1:{}/ (Ctrl-C to stop)",
        file.display(),
        port
    );
    for stream in listener.incoming() {
        let Ok(stream) = stream else { continue };
        if let Err(e) = handle(stream, file) {
            eprintln!("watch: {e:#}");
        }
    }
    Ok(())
}

fn version_stamp(file: &Path) -> String {
    std::fs::metadata(file)
        .and_then(|m| m.modified())
        .ok()
        .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|d| format!("{}.{}", d.as_secs(), d.subsec_nanos()))
        .unwrap_or_else(|| "missing".to_string())
}

fn handle(mut stream: TcpStream, file: &Path) -> Result<()> {
    let mut reader = BufReader::new(stream.try_clone()?);
    let mut request_line = String::new();
    reader.read_line(&mut request_line)?;
    // Drain headers (Connection: close, so no keep-alive bookkeeping).
    let mut line = String::new();
    while reader.read_line(&mut line)? > 2 {
        line.clear();
    }
    let path = request_line.split_whitespace().nth(1).unwrap_or("/");

    let (status, content_type, body) = match path {
        "/__v" => ("200 OK", "text/plain", version_stamp(file)),
        "/" | "/index.html" => {
            let html = match render(file) {
                Ok(html) => html,
                Err(e) => crate::html::page(
                    "officecli watch",
                    &format!(
                        "<p><strong>Cannot render {}:</strong> {}</p><p>The page reloads automatically once the file is readable again.</p>",
                        crate::html::escape(&file.display().to_string()),
                        crate::html::escape(&format!("{e:#}"))
                    ),
                ),
            };
            let html = match html.rfind("</body>") {
                Some(i) => format!("{}{}{}", &html[..i], RELOAD_SCRIPT, &html[i..]),
                None => format!("{html}{RELOAD_SCRIPT}"),
            };
            ("200 OK", "text/html; charset=utf-8", html)
        }
        _ => ("404 Not Found", "text/plain", "not found".to_string()),
    };
    write!(
        stream,
        "HTTP/1.1 {status}\r\nContent-Type: {content_type}\r\nContent-Length: {}\r\nConnection: close\r\n\r\n{body}",
        body.len()
    )?;
    stream.flush()?;
    Ok(())
}

fn render(file: &Path) -> Result<String> {
    let mut handler = crate::open_handler(file)?;
    Ok(handler.view("html")?.render(false))
}
