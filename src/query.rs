//! CSS-like selectors for `officecli query`, matched against the NodeInfo
//! trees the handlers already produce.
//!
//! Grammar (a subset of upstream OfficeCLI's):
//!   selector  = compound ( '>' compound )*
//!   compound  = [type] predicate*
//!   predicate = '[' attr op value ']' | ':contains("text")' | ':empty'
//!   op        = '=' | '!=' | '~=' | '>=' | '<=' | '>' | '<'
//!
//! `~=` is substring match. `>=`/`<=`/`>`/`<` compare numerically when both
//! sides parse as numbers, else lexicographically. A chain `a > b` matches
//! `b` nodes whose direct parent matches `a`.

use anyhow::{bail, Result};

use crate::out::NodeInfo;

#[derive(Debug, Clone, PartialEq)]
pub enum Op {
    Eq,
    Ne,
    Contains,
    Gte,
    Lte,
    Gt,
    Lt,
}

#[derive(Debug, Clone)]
pub enum Pred {
    Attr(String, Op, String),
    TextContains(String),
    Empty,
}

#[derive(Debug, Clone)]
pub struct Compound {
    /// Element type ("paragraph", "cell", ...); empty = any.
    pub typ: String,
    pub preds: Vec<Pred>,
}

#[derive(Debug, Clone)]
pub struct Selector {
    pub parts: Vec<Compound>,
}

/// Friendly type aliases shared across formats.
fn normalize_type(t: &str) -> String {
    match t.to_ascii_lowercase().as_str() {
        "p" | "para" => "paragraph".into(),
        "r" => "run".into(),
        "tbl" => "table".into(),
        "tr" => "row".into(),
        "tc" | "td" => "cell".into(),
        other => other.to_string(),
    }
}

pub fn parse(selector: &str) -> Result<Selector> {
    let mut parts = Vec::new();
    for raw in split_top_level(selector, '>') {
        let raw = raw.trim();
        if raw.is_empty() {
            bail!("empty compound in selector '{selector}'");
        }
        parts.push(parse_compound(raw)?);
    }
    if parts.is_empty() {
        bail!("empty selector");
    }
    Ok(Selector { parts })
}

/// Split on `sep` outside of brackets and quotes.
fn split_top_level(s: &str, sep: char) -> Vec<String> {
    let mut out = Vec::new();
    let mut cur = String::new();
    let mut depth = 0usize;
    let mut quote: Option<char> = None;
    for ch in s.chars() {
        match quote {
            Some(q) => {
                cur.push(ch);
                if ch == q {
                    quote = None;
                }
            }
            None => match ch {
                '"' | '\'' => {
                    quote = Some(ch);
                    cur.push(ch);
                }
                '[' | '(' => {
                    depth += 1;
                    cur.push(ch);
                }
                ']' | ')' => {
                    depth = depth.saturating_sub(1);
                    cur.push(ch);
                }
                c if c == sep && depth == 0 => out.push(std::mem::take(&mut cur)),
                _ => cur.push(ch),
            },
        }
    }
    if !cur.trim().is_empty() {
        out.push(cur);
    }
    out
}

fn parse_compound(raw: &str) -> Result<Compound> {
    let mut typ = String::new();
    let mut preds = Vec::new();
    let mut rest = raw;

    // Leading type name up to '[' or ':'.
    let type_end = rest
        .find(['[', ':'])
        .unwrap_or(rest.len());
    if type_end > 0 {
        typ = normalize_type(rest[..type_end].trim());
    }
    rest = &rest[type_end..];

    while !rest.is_empty() {
        if let Some(inner_start) = rest.strip_prefix('[') {
            let end = find_matching(inner_start, ']')
                .ok_or_else(|| anyhow::anyhow!("unclosed '[' in selector '{raw}'"))?;
            let inner = &inner_start[..end];
            preds.push(parse_attr_pred(inner)?);
            rest = &inner_start[end + 1..];
        } else if let Some(after) = rest.strip_prefix(':') {
            if let Some(args) = after.strip_prefix("contains(") {
                let end = find_matching(args, ')')
                    .ok_or_else(|| anyhow::anyhow!("unclosed ':contains(' in '{raw}'"))?;
                let text = args[..end]
                    .trim()
                    .trim_matches('"')
                    .trim_matches('\'')
                    .to_string();
                preds.push(Pred::TextContains(text));
                rest = &args[end + 1..];
            } else if let Some(r) = after.strip_prefix("empty") {
                preds.push(Pred::Empty);
                rest = r;
            } else {
                bail!("unknown pseudo-class in selector '{raw}' (supported: :contains, :empty)");
            }
        } else {
            bail!("unexpected characters '{rest}' in selector '{raw}'");
        }
    }
    Ok(Compound { typ, preds })
}

/// Position of the closing delimiter, respecting quotes.
fn find_matching(s: &str, close: char) -> Option<usize> {
    let mut quote: Option<char> = None;
    for (i, ch) in s.char_indices() {
        match quote {
            Some(q) => {
                if ch == q {
                    quote = None;
                }
            }
            None => match ch {
                '"' | '\'' => quote = Some(ch),
                c if c == close => return Some(i),
                _ => {}
            },
        }
    }
    None
}

fn parse_attr_pred(inner: &str) -> Result<Pred> {
    for (token, op) in [
        ("!=", Op::Ne),
        ("~=", Op::Contains),
        (">=", Op::Gte),
        ("<=", Op::Lte),
        ("=", Op::Eq),
        (">", Op::Gt),
        ("<", Op::Lt),
    ] {
        if let Some(idx) = inner.find(token) {
            let attr = inner[..idx].trim().trim_start_matches('@').to_string();
            let value = inner[idx + token.len()..]
                .trim()
                .trim_matches('"')
                .trim_matches('\'')
                .to_string();
            if attr.is_empty() {
                bail!("predicate '[{inner}]' has no attribute name");
            }
            return Ok(Pred::Attr(attr, op, value));
        }
    }
    bail!("predicate '[{inner}]' has no operator (=, !=, ~=, >=, <=, >, <)")
}

fn compare(op: &Op, actual: &str, expected: &str) -> bool {
    match op {
        Op::Eq => actual == expected,
        Op::Ne => actual != expected,
        Op::Contains => actual.contains(expected),
        Op::Gte | Op::Lte | Op::Gt | Op::Lt => {
            let ordering = match (actual.parse::<f64>(), expected.parse::<f64>()) {
                (Ok(a), Ok(b)) => a.partial_cmp(&b),
                (Err(_), Err(_)) => Some(actual.cmp(expected)),
                // Mixed numeric/text never satisfies an ordered comparison.
                _ => None,
            };
            match (op, ordering) {
                (Op::Gte, Some(o)) => o.is_ge(),
                (Op::Lte, Some(o)) => o.is_le(),
                (Op::Gt, Some(o)) => o.is_gt(),
                (Op::Lt, Some(o)) => o.is_lt(),
                _ => false,
            }
        }
    }
}

fn matches_compound(node: &NodeInfo, compound: &Compound) -> bool {
    if !compound.typ.is_empty() && normalize_type(&node.kind) != compound.typ {
        return false;
    }
    compound.preds.iter().all(|pred| match pred {
        Pred::Attr(attr, op, expected) => {
            // "text" and "value" fall back to the node text.
            let owned;
            let actual: &str = match node
                .attributes
                .iter()
                .find(|(k, _)| k.eq_ignore_ascii_case(attr))
            {
                Some((_, v)) => v,
                None if attr.eq_ignore_ascii_case("text")
                    || attr.eq_ignore_ascii_case("value") =>
                {
                    owned = node.text.clone().unwrap_or_default();
                    &owned
                }
                None => return matches!(op, Op::Ne),
            };
            compare(op, actual, expected)
        }
        Pred::TextContains(text) => node
            .text
            .as_deref()
            .map(|t| t.contains(text.as_str()))
            .unwrap_or(false),
        Pred::Empty => node.text.as_deref().map(|t| t.trim().is_empty()).unwrap_or(true)
            && node.children.is_empty(),
    })
}

/// Run a selector against a NodeInfo forest. Returns matching nodes
/// (children stripped) in document order.
pub fn run(selector: &Selector, roots: &[NodeInfo]) -> Vec<NodeInfo> {
    let mut results = Vec::new();
    // ancestors: whether each compound prefix is satisfied along the path.
    fn walk(
        node: &NodeInfo,
        parent_matched: &[bool],
        selector: &Selector,
        results: &mut Vec<NodeInfo>,
    ) {
        // matched[i] = this node completes compounds[0..=i] with direct-child
        // chaining relative to the parent's matched prefix.
        let n = selector.parts.len();
        let mut matched = vec![false; n];
        for i in 0..n {
            let prefix_ok = if i == 0 { true } else { parent_matched[i - 1] };
            if prefix_ok && matches_compound(node, &selector.parts[i]) {
                matched[i] = true;
            }
        }
        if matched[n - 1] {
            let mut flat = node.clone();
            flat.children.clear();
            results.push(flat);
        }
        for child in &node.children {
            walk(child, &matched, selector, results);
        }
    }
    let no_parent = vec![false; selector.parts.len()];
    for root in roots {
        walk(root, &no_parent, selector, &mut results);
    }
    results
}

#[cfg(test)]
mod tests {
    use super::*;

    fn node(kind: &str, text: &str, attrs: &[(&str, &str)]) -> NodeInfo {
        let mut n = NodeInfo::new(format!("/{kind}"), kind);
        n.text = Some(text.to_string());
        for (k, v) in attrs {
            n.attr(*k, *v);
        }
        n
    }

    #[test]
    fn parses_and_matches() {
        let sel = parse("paragraph[style=Normal] > run[font!=Arial]").unwrap();
        assert_eq!(sel.parts.len(), 2);

        let mut p = node("paragraph", "hi", &[("style", "Normal")]);
        p.children.push(node("run", "hi", &[("font", "Times")]));
        p.children.push(node("run", "yo", &[("font", "Arial")]));
        let hits = run(&sel, &[p]);
        assert_eq!(hits.len(), 1);
        assert_eq!(hits[0].text.as_deref(), Some("hi"));
    }

    #[test]
    fn numeric_comparison_and_contains() {
        let sel = parse("cell[value>100]").unwrap();
        let cells = vec![
            node("cell", "150", &[]),
            node("cell", "50", &[]),
            node("cell", "abc", &[]),
        ];
        let hits = run(&sel, &cells);
        assert_eq!(hits.len(), 1);

        let sel = parse(r#":contains("needle")"#).unwrap();
        let hits = run(&sel, &[node("paragraph", "a needle here", &[])]);
        assert_eq!(hits.len(), 1);
    }

    #[test]
    fn missing_attr_matches_ne() {
        let sel = parse("run[font!=Arial]").unwrap();
        let hits = run(&sel, &[node("run", "x", &[])]);
        assert_eq!(hits.len(), 1);
    }
}
