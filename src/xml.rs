//! Minimal lossless XML DOM used to read and edit OOXML parts.
//!
//! Qualified names and attribute order are preserved verbatim so that
//! round-tripping a document part does not disturb content we don't touch.

use anyhow::{bail, Context, Result};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::Cursor;

#[derive(Debug, Clone, PartialEq)]
pub enum XmlNode {
    Element(XmlElement),
    Text(String),
}

impl XmlNode {
    pub fn as_element(&self) -> Option<&XmlElement> {
        match self {
            XmlNode::Element(e) => Some(e),
            _ => None,
        }
    }

    pub fn as_element_mut(&mut self) -> Option<&mut XmlElement> {
        match self {
            XmlNode::Element(e) => Some(e),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct XmlElement {
    /// Qualified name as it appears in the source, e.g. `w:p`.
    pub name: String,
    /// Attributes with qualified names, order preserved.
    pub attrs: Vec<(String, String)>,
    pub children: Vec<XmlNode>,
}

impl XmlElement {
    pub fn new(name: &str) -> Self {
        XmlElement {
            name: name.to_string(),
            attrs: Vec::new(),
            children: Vec::new(),
        }
    }

    /// Local name with any namespace prefix stripped.
    pub fn local_name(&self) -> &str {
        match self.name.rfind(':') {
            Some(i) => &self.name[i + 1..],
            None => &self.name,
        }
    }

    pub fn attr(&self, qname: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(k, _)| k == qname)
            .map(|(_, v)| v.as_str())
    }

    /// Attribute lookup by local name (prefix-insensitive).
    pub fn attr_local(&self, local: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(k, _)| k.rsplit(':').next() == Some(local))
            .map(|(_, v)| v.as_str())
    }

    pub fn set_attr(&mut self, qname: &str, value: &str) {
        if let Some(pair) = self.attrs.iter_mut().find(|(k, _)| k == qname) {
            pair.1 = value.to_string();
        } else {
            self.attrs.push((qname.to_string(), value.to_string()));
        }
    }

    pub fn remove_attr(&mut self, qname: &str) {
        self.attrs.retain(|(k, _)| k != qname);
    }

    pub fn child(&self, local: &str) -> Option<&XmlElement> {
        self.children.iter().find_map(|n| match n {
            XmlNode::Element(e) if e.local_name() == local => Some(e),
            _ => None,
        })
    }

    pub fn child_mut(&mut self, local: &str) -> Option<&mut XmlElement> {
        self.children.iter_mut().find_map(|n| match n {
            XmlNode::Element(e) if e.local_name() == local => Some(e),
            _ => None,
        })
    }

    pub fn children_named(&self, local: &str) -> Vec<&XmlElement> {
        self.children
            .iter()
            .filter_map(|n| n.as_element())
            .filter(|e| e.local_name() == local)
            .collect()
    }

    pub fn elements(&self) -> impl Iterator<Item = &XmlElement> {
        self.children.iter().filter_map(|n| n.as_element())
    }

    /// Index (in `children`) of the Nth (0-based) element child named `local`.
    pub fn nth_child_index(&self, local: &str, n: usize) -> Option<usize> {
        let mut seen = 0;
        for (i, node) in self.children.iter().enumerate() {
            if let XmlNode::Element(e) = node {
                if e.local_name() == local {
                    if seen == n {
                        return Some(i);
                    }
                    seen += 1;
                }
            }
        }
        None
    }

    /// Get or create the first child element with local name `local`,
    /// using `qname` when creating. Created child is inserted at `front`.
    pub fn ensure_child(&mut self, local: &str, qname: &str, front: bool) -> &mut XmlElement {
        let pos = self.children.iter().position(|n| {
            matches!(n, XmlNode::Element(e) if e.local_name() == local)
        });
        let idx = match pos {
            Some(i) => i,
            None => {
                let el = XmlNode::Element(XmlElement::new(qname));
                if front {
                    self.children.insert(0, el);
                    0
                } else {
                    self.children.push(el);
                    self.children.len() - 1
                }
            }
        };
        self.children[idx].as_element_mut().unwrap()
    }

    pub fn push(&mut self, el: XmlElement) {
        self.children.push(XmlNode::Element(el));
    }

    pub fn push_text(&mut self, text: &str) {
        self.children.push(XmlNode::Text(text.to_string()));
    }

    /// Concatenated text of all descendant text nodes.
    pub fn text_content(&self) -> String {
        let mut out = String::new();
        collect_text(self, &mut out);
        out
    }
}

fn collect_text(el: &XmlElement, out: &mut String) {
    for child in &el.children {
        match child {
            XmlNode::Text(t) => out.push_str(t),
            XmlNode::Element(e) => collect_text(e, out),
        }
    }
}

/// Build an element with attributes in one expression.
pub fn el(name: &str, attrs: &[(&str, &str)]) -> XmlElement {
    let mut e = XmlElement::new(name);
    for (k, v) in attrs {
        e.attrs.push((k.to_string(), v.to_string()));
    }
    e
}

pub fn parse(bytes: &[u8]) -> Result<XmlElement> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().expand_empty_elements = false;
    let mut stack: Vec<XmlElement> = Vec::new();
    let mut root: Option<XmlElement> = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                stack.push(start_to_element(&e)?);
            }
            Event::Empty(e) => {
                let element = start_to_element(&e)?;
                attach(&mut stack, &mut root, element)?;
            }
            Event::End(_) => {
                let element = stack.pop().context("unbalanced XML end tag")?;
                attach(&mut stack, &mut root, element)?;
            }
            Event::Text(t) => {
                let text = t.unescape()?.into_owned();
                if let Some(parent) = stack.last_mut() {
                    // Skip pure inter-element whitespace at parse time only when
                    // the parent has element children context; OOXML never uses
                    // significant mixed content outside text runs, and text runs
                    // (`w:t`, `a:t`, `t`) contain a single text node anyway.
                    if !text.trim().is_empty() || parent.local_name() == "t" {
                        parent.children.push(XmlNode::Text(text));
                    }
                }
            }
            Event::CData(t) => {
                let text = String::from_utf8_lossy(t.as_ref()).into_owned();
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(XmlNode::Text(text));
                }
            }
            Event::Eof => break,
            _ => {} // declarations, comments, processing instructions
        }
        buf.clear();
    }

    root.context("no root element found")
}

fn start_to_element(e: &BytesStart) -> Result<XmlElement> {
    let name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
    let mut element = XmlElement::new(&name);
    for attr in e.attributes() {
        let attr = attr?;
        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        let value = attr.unescape_value()?.into_owned();
        element.attrs.push((key, value));
    }
    Ok(element)
}

fn attach(
    stack: &mut [XmlElement],
    root: &mut Option<XmlElement>,
    element: XmlElement,
) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlNode::Element(element));
    } else if root.is_none() {
        *root = Some(element);
    } else {
        bail!("multiple root elements");
    }
    Ok(())
}

pub fn serialize(root: &XmlElement) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    writer.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;
    write_element(&mut writer, root)?;
    Ok(writer.into_inner().into_inner())
}

fn write_element(writer: &mut Writer<Cursor<Vec<u8>>>, el: &XmlElement) -> Result<()> {
    let mut start = BytesStart::new(&el.name);
    for (k, v) in &el.attrs {
        start.push_attribute((k.as_str(), v.as_str()));
    }
    if el.children.is_empty() {
        writer.write_event(Event::Empty(start))?;
        return Ok(());
    }
    writer.write_event(Event::Start(start))?;
    for child in &el.children {
        match child {
            XmlNode::Element(e) => write_element(writer, e)?,
            XmlNode::Text(t) => writer.write_event(Event::Text(BytesText::new(t)))?,
        }
    }
    writer.write_event(Event::End(BytesEnd::new(&el.name)))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn roundtrip_preserves_structure() {
        let src = br#"<?xml version="1.0"?><w:document xmlns:w="urn:x"><w:body><w:p w:rsidR="00A"><w:r><w:t xml:space="preserve">hello &amp; goodbye </w:t></w:r></w:p></w:body></w:document>"#;
        let doc = parse(src).unwrap();
        assert_eq!(doc.name, "w:document");
        let body = doc.child("body").unwrap();
        let p = body.child("p").unwrap();
        assert_eq!(p.attr("w:rsidR"), Some("00A"));
        assert_eq!(p.text_content(), "hello & goodbye ");
        let out = serialize(&doc).unwrap();
        let again = parse(&out).unwrap();
        assert_eq!(doc, again);
    }
}
