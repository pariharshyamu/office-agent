//! Shared helpers for `view html`: escaping and the standalone page shell.

pub fn escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&#39;"),
            _ => out.push(ch),
        }
    }
    out
}

pub fn page(title: &str, body: &str) -> String {
    format!(
        r#"<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>{title}</title>
<style>
  body {{ font-family: Calibri, 'Segoe UI', Arial, sans-serif; margin: 2rem auto; max-width: 60rem; color: #1a1a1a; }}
  table {{ border-collapse: collapse; margin: 0.75rem 0; }}
  td, th {{ border: 1px solid #999; padding: 0.25rem 0.6rem; }}
  th {{ background: #f0f0f0; }}
  .sheet-name {{ margin: 1.5rem 0 0.25rem; }}
  .row-num {{ background: #f0f0f0; color: #666; font-weight: normal; }}
  .slide {{ position: relative; background: white; border: 1px solid #ccc; margin: 1.5rem 0; overflow: hidden; }}
  .slide .shape {{ position: absolute; overflow: hidden; white-space: pre-wrap; }}
  .slide-label {{ color: #888; font-size: 0.85rem; margin-top: 1.5rem; }}
</style>
</head>
<body>
{body}
</body>
</html>
"#,
        title = escape(title),
        body = body
    )
}
