//! Document path parser.
//!
//! Paths address elements with 1-based indices, XPath style:
//!   /body/p[3]           third paragraph
//!   /slide[1]/shape[2]   second shape on slide 1
//!   /slide[1]/shape[@name=Title 1]
//!   /Sheet1/A1           Excel cell (sheet segment + cell reference)
//!
//! A missing index means "first" when resolving a single element and
//! "all" when resolving a set (query-like contexts).

use anyhow::{bail, Result};

#[derive(Debug, Clone, PartialEq)]
pub enum Predicate {
    /// `[3]` — 1-based position among same-named siblings.
    Index(usize),
    /// `[@name=Title 1]` — attribute equality.
    Attr(String, String),
}

#[derive(Debug, Clone, PartialEq)]
pub struct Segment {
    pub name: String,
    pub preds: Vec<Predicate>,
}

impl Segment {
    pub fn index(&self) -> Option<usize> {
        self.preds.iter().find_map(|p| match p {
            Predicate::Index(i) => Some(*i),
            _ => None,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct DocPath {
    pub segments: Vec<Segment>,
}

impl DocPath {
    pub fn is_root(&self) -> bool {
        self.segments.is_empty()
    }
}

pub fn parse(path: &str) -> Result<DocPath> {
    let trimmed = path.trim();
    // Excel-style `$Sheet1:A1` addressing is normalized to `/Sheet1/A1`.
    let normalized = if let Some(rest) = trimmed.strip_prefix('$') {
        match rest.split_once(':') {
            Some((sheet, cell)) => format!("/{sheet}/{cell}"),
            None => format!("/{rest}"),
        }
    } else {
        trimmed.to_string()
    };

    let body = normalized.strip_prefix('/').unwrap_or(&normalized);
    let mut segments = Vec::new();
    if body.is_empty() {
        return Ok(DocPath { segments });
    }
    for raw in split_segments(body)? {
        segments.push(parse_segment(&raw)?);
    }
    Ok(DocPath { segments })
}

/// Split on `/` but not inside `[...]` (attribute values may contain `/`).
fn split_segments(body: &str) -> Result<Vec<String>> {
    let mut out = Vec::new();
    let mut cur = String::new();
    let mut depth = 0usize;
    for ch in body.chars() {
        match ch {
            '[' => {
                depth += 1;
                cur.push(ch);
            }
            ']' => {
                depth = depth.saturating_sub(1);
                cur.push(ch);
            }
            '/' if depth == 0 => {
                if cur.is_empty() {
                    bail!("empty path segment in '{body}'");
                }
                out.push(std::mem::take(&mut cur));
            }
            _ => cur.push(ch),
        }
    }
    if depth != 0 {
        bail!("unbalanced brackets in path '{body}'");
    }
    if !cur.is_empty() {
        out.push(cur);
    }
    Ok(out)
}

fn parse_segment(raw: &str) -> Result<Segment> {
    let (name, rest) = match raw.find('[') {
        Some(i) => (&raw[..i], &raw[i..]),
        None => (raw, ""),
    };
    if name.is_empty() {
        bail!("path segment '{raw}' has no element name");
    }
    let mut preds = Vec::new();
    let mut remainder = rest;
    while let Some(stripped) = remainder.strip_prefix('[') {
        let end = stripped
            .find(']')
            .ok_or_else(|| anyhow::anyhow!("unclosed '[' in path segment '{raw}'"))?;
        let inner = &stripped[..end];
        remainder = &stripped[end + 1..];
        if let Some(attr) = inner.strip_prefix('@') {
            match attr.split_once('=') {
                Some((k, v)) => preds.push(Predicate::Attr(
                    k.trim().to_string(),
                    v.trim().trim_matches('"').trim_matches('\'').to_string(),
                )),
                None => bail!("attribute predicate '[@{attr}]' needs '=value'"),
            }
        } else {
            let n: usize = inner.trim().parse().map_err(|_| {
                anyhow::anyhow!("index '[{inner}]' is not a number (1-based indices)")
            })?;
            if n == 0 {
                bail!("indices are 1-based; '[0]' is invalid");
            }
            preds.push(Predicate::Index(n));
        }
    }
    Ok(Segment {
        name: name.to_string(),
        preds,
    })
}

/// Parse an A1-style cell reference into (column 0-based, row 0-based).
pub fn parse_cell_ref(cell: &str) -> Option<(u32, u32)> {
    let mut col: u64 = 0;
    let mut chars = cell.chars().peekable();
    let mut letters = 0;
    while let Some(&c) = chars.peek() {
        if c.is_ascii_alphabetic() {
            col = col * 26 + (c.to_ascii_uppercase() as u64 - 'A' as u64 + 1);
            letters += 1;
            chars.next();
        } else {
            break;
        }
    }
    if letters == 0 || letters > 3 {
        return None;
    }
    let digits: String = chars.collect();
    if digits.is_empty() || !digits.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    let row: u32 = digits.parse().ok()?;
    if row == 0 || col == 0 || col > 16384 {
        return None;
    }
    Some((col as u32 - 1, row - 1))
}

/// Convert 0-based column index to letters: 0 -> A, 27 -> AB.
pub fn col_letters(mut col: u32) -> String {
    let mut out = String::new();
    loop {
        out.insert(0, (b'A' + (col % 26) as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    out
}

pub fn cell_name(col: u32, row: u32) -> String {
    format!("{}{}", col_letters(col), row + 1)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_indexed_path() {
        let p = parse("/slide[1]/shape[2]").unwrap();
        assert_eq!(p.segments.len(), 2);
        assert_eq!(p.segments[0].name, "slide");
        assert_eq!(p.segments[0].index(), Some(1));
        assert_eq!(p.segments[1].index(), Some(2));
    }

    #[test]
    fn parses_attr_predicate() {
        let p = parse("/slide[1]/shape[@name=Title 1]").unwrap();
        assert_eq!(
            p.segments[1].preds[0],
            Predicate::Attr("name".into(), "Title 1".into())
        );
    }

    #[test]
    fn parses_root_and_dollar_form() {
        assert!(parse("/").unwrap().is_root());
        let p = parse("$Sheet1:B2").unwrap();
        assert_eq!(p.segments[0].name, "Sheet1");
        assert_eq!(p.segments[1].name, "B2");
    }

    #[test]
    fn cell_refs() {
        assert_eq!(parse_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("AB10"), Some((27, 9)));
        assert_eq!(parse_cell_ref("1A"), None);
        assert_eq!(cell_name(27, 9), "AB10");
        assert_eq!(col_letters(0), "A");
        assert_eq!(col_letters(25), "Z");
        assert_eq!(col_letters(26), "AA");
    }
}
