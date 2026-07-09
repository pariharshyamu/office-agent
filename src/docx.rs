//! Word (.docx) handler: paragraphs, runs, tables, formatting, find/replace.

use anyhow::{bail, Context, Result};
use regex::Regex;
use serde_json::json;
use std::path::Path;

use crate::handler::{Handler, Position};
use crate::out::{NodeInfo, Report};
use crate::path::{self, DocPath, Predicate, Segment};
use crate::props::{parse_align, parse_bool, parse_color, parse_pt, Props};
use crate::pkg::Package;
use crate::xml::{el, XmlElement, XmlNode};

const DOCUMENT_PART: &str = "word/document.xml";
const DOC_RELS_PART: &str = "word/_rels/document.xml.rels";

pub struct Docx {
    pkg: Package,
    doc: XmlElement,
    rels: XmlElement,
}

/// Map friendly element names to OOXML local names.
fn ooxml_name(name: &str) -> Option<&'static str> {
    match name.to_ascii_lowercase().as_str() {
        "body" => Some("body"),
        "p" | "paragraph" | "para" => Some("p"),
        "r" | "run" => Some("r"),
        "tbl" | "table" => Some("tbl"),
        "tr" | "row" => Some("tr"),
        "tc" | "td" | "cell" => Some("tc"),
        "hyperlink" => Some("hyperlink"),
        "sectpr" | "section" => Some("sectPr"),
        _ => None,
    }
}

/// Friendly display name for an OOXML local name.
fn friendly_name(local: &str) -> &str {
    match local {
        "p" => "paragraph",
        "r" => "run",
        "tbl" => "table",
        "tr" => "row",
        "tc" => "cell",
        other => other,
    }
}

impl Docx {
    pub fn new(pkg: Package) -> Result<Docx> {
        let doc = pkg.xml(DOCUMENT_PART)?;
        let rels = pkg.xml(DOC_RELS_PART)?;
        Ok(Docx { pkg, doc, rels })
    }

    fn build_image_paragraph(&mut self, props: &Props) -> Result<XmlElement> {
        let src = props
            .get("src")
            .context("image needs --prop src=path/to/file.png")?;
        let image = crate::media::load_image(src)?;
        let part = crate::media::store_image(&mut self.pkg, "word/media", &image)?;
        let target = part.strip_prefix("word/").unwrap_or(&part).to_string();
        let rid =
            crate::media::add_relationship(&mut self.rels, crate::media::IMAGE_REL_TYPE, &target);

        let w = props
            .get("w")
            .or_else(|| props.get("width"))
            .map(crate::props::parse_emu)
            .transpose()?
            .unwrap_or(image.width_emu);
        // Keep aspect ratio when only one dimension is given.
        let h = match props
            .get("h")
            .or_else(|| props.get("height"))
            .map(crate::props::parse_emu)
            .transpose()?
        {
            Some(h) => h,
            None if w != image.width_emu && image.width_emu > 0 => {
                image.height_emu * w / image.width_emu
            }
            None => image.height_emu,
        };
        let pic_id = self.next_drawing_id();
        let name = props
            .get("name")
            .map(|s| s.to_string())
            .unwrap_or_else(|| format!("Picture {pic_id}"));

        let mut inline = el(
            "wp:inline",
            &[
                ("distT", "0"),
                ("distB", "0"),
                ("distL", "0"),
                ("distR", "0"),
                (
                    "xmlns:wp",
                    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                ),
            ],
        );
        inline.push(el(
            "wp:extent",
            &[("cx", w.to_string().as_str()), ("cy", h.to_string().as_str())],
        ));
        inline.push(el(
            "wp:docPr",
            &[("id", pic_id.to_string().as_str()), ("name", name.as_str())],
        ));
        let mut graphic = el(
            "a:graphic",
            &[("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main")],
        );
        let mut gdata = el(
            "a:graphicData",
            &[("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")],
        );
        let mut pic = el(
            "pic:pic",
            &[(
                "xmlns:pic",
                "http://schemas.openxmlformats.org/drawingml/2006/picture",
            )],
        );
        let mut nv = XmlElement::new("pic:nvPicPr");
        nv.push(el(
            "pic:cNvPr",
            &[("id", pic_id.to_string().as_str()), ("name", name.as_str())],
        ));
        nv.push(XmlElement::new("pic:cNvPicPr"));
        pic.push(nv);
        let mut fill = XmlElement::new("pic:blipFill");
        fill.push(el("a:blip", &[("r:embed", rid.as_str())]));
        let mut stretch = XmlElement::new("a:stretch");
        stretch.push(XmlElement::new("a:fillRect"));
        fill.push(stretch);
        pic.push(fill);
        let mut sppr = XmlElement::new("pic:spPr");
        let mut xfrm = XmlElement::new("a:xfrm");
        xfrm.push(el("a:off", &[("x", "0"), ("y", "0")]));
        xfrm.push(el(
            "a:ext",
            &[("cx", w.to_string().as_str()), ("cy", h.to_string().as_str())],
        ));
        sppr.push(xfrm);
        let mut geom = el("a:prstGeom", &[("prst", "rect")]);
        geom.push(XmlElement::new("a:avLst"));
        sppr.push(geom);
        pic.push(sppr);
        gdata.push(pic);
        graphic.push(gdata);
        inline.push(graphic);

        let mut drawing = XmlElement::new("w:drawing");
        drawing.push(inline);
        let mut r = XmlElement::new("w:r");
        r.push(drawing);
        let mut p = XmlElement::new("w:p");
        if let Some(align) = props.get("align") {
            let (jc, _) = parse_align(align)?;
            let mut ppr = XmlElement::new("w:pPr");
            ppr.push(el("w:jc", &[("w:val", jc)]));
            p.push(ppr);
        }
        p.push(r);
        Ok(p)
    }

    fn next_drawing_id(&self) -> u64 {
        let mut max_id = 0;
        fn walk(e: &XmlElement, max_id: &mut u64) {
            if matches!(e.local_name(), "docPr" | "cNvPr") {
                if let Some(id) = e.attr_local("id").and_then(|v| v.parse::<u64>().ok()) {
                    *max_id = (*max_id).max(id);
                }
            }
            for c in e.elements() {
                walk(c, max_id);
            }
        }
        walk(&self.doc, &mut max_id);
        max_id + 1
    }

    fn body(&self) -> Result<&XmlElement> {
        self.doc.child("body").context("document has no <w:body>")
    }

    fn body_mut(&mut self) -> Result<&mut XmlElement> {
        self.doc
            .child_mut("body")
            .context("document has no <w:body>")
    }

    /// Resolve a path to child indices from the body element.
    /// Returns the chain of `children` indices to walk.
    fn resolve(&self, dpath: &DocPath) -> Result<Vec<usize>> {
        let body = self.body()?;
        let mut segments = dpath.segments.as_slice();
        // A leading /body segment is optional.
        if segments
            .first()
            .map(|s| s.name.eq_ignore_ascii_case("body"))
            .unwrap_or(false)
        {
            segments = &segments[1..];
        }
        let mut indices = Vec::new();
        let mut current = body;
        for seg in segments {
            let idx = find_child(current, seg)?;
            indices.push(idx);
            current = current.children[idx].as_element().unwrap();
        }
        Ok(indices)
    }

    fn node_at(&self, indices: &[usize]) -> Result<&XmlElement> {
        let mut current = self.body()?;
        for &i in indices {
            current = current.children[i]
                .as_element()
                .context("internal: resolved path hit a text node")?;
        }
        Ok(current)
    }

    fn node_at_mut(&mut self, indices: &[usize]) -> Result<&mut XmlElement> {
        let mut current = self.body_mut()?;
        for &i in indices {
            current = current.children[i]
                .as_element_mut()
                .context("internal: resolved path hit a text node")?;
        }
        Ok(current)
    }

    /// Canonical display path (e.g. /body/p[3]) for a resolved index chain.
    fn display_path(&self, indices: &[usize]) -> Result<String> {
        let mut out = String::from("/body");
        let mut current = self.body()?;
        for &i in indices {
            let element = current.children[i].as_element().unwrap();
            let local = element.local_name().to_string();
            let mut nth = 0;
            for node in &current.children[..i] {
                if let XmlNode::Element(e) = node {
                    if e.local_name() == local {
                        nth += 1;
                    }
                }
            }
            out.push_str(&format!("/{}[{}]", local, nth + 1));
            current = element;
        }
        Ok(out)
    }

    fn node_info(&self, element: &XmlElement, path: &str, depth: usize) -> NodeInfo {
        let local = element.local_name();
        let mut info = NodeInfo::new(path, friendly_name(local));
        match local {
            "p" => {
                info.text = Some(paragraph_text(element));
                if let Some(ppr) = element.child("pPr") {
                    if let Some(style) = ppr.child("pStyle").and_then(|s| s.attr_local("val")) {
                        info.attr("style", style);
                    }
                    if let Some(jc) = ppr.child("jc").and_then(|s| s.attr_local("val")) {
                        info.attr("align", jc);
                    }
                }
            }
            "r" => {
                info.text = Some(run_text(element));
                if let Some(rpr) = element.child("rPr") {
                    describe_rpr(rpr, &mut info);
                }
            }
            "tbl" => {
                let rows = element.children_named("tr").len();
                let cols = element
                    .children_named("tr")
                    .first()
                    .map(|tr| tr.children_named("tc").len())
                    .unwrap_or(0);
                info.attr("rows", rows.to_string());
                info.attr("cols", cols.to_string());
            }
            "tr" | "tc" => {
                info.text = Some(element.text_content());
            }
            _ => {}
        }
        if depth > 0 {
            let mut counts: std::collections::HashMap<String, usize> = Default::default();
            for child in element.elements() {
                let cl = child.local_name();
                if matches!(cl, "pPr" | "rPr" | "tblPr" | "tblGrid" | "trPr" | "tcPr" | "sectPr") {
                    continue;
                }
                let n = counts.entry(cl.to_string()).or_insert(0);
                *n += 1;
                let child_path = format!("{}/{}[{}]", path, cl, *n);
                info.children
                    .push(self.node_info(child, &child_path, depth - 1));
            }
        }
        info
    }

    // ------------------------------------------------------------- add ----

    fn build_paragraph(&self, props: &Props) -> Result<XmlElement> {
        let mut p = XmlElement::new("w:p");
        let mut ppr = XmlElement::new("w:pPr");
        if let Some(style) = props.get("style") {
            ppr.push(el("w:pStyle", &[("w:val", style)]));
        }
        if let Some(align) = props.get("align") {
            let (jc, _) = parse_align(align)?;
            ppr.push(el("w:jc", &[("w:val", jc)]));
        }
        if !ppr.children.is_empty() {
            p.push(ppr);
        }
        let text = props.get("text").unwrap_or("");
        if !text.is_empty() {
            p.push(build_run(text, props)?);
        }
        Ok(p)
    }

    fn build_table(&self, props: &Props) -> Result<XmlElement> {
        let rows: usize = props
            .get("rows")
            .map(|v| v.parse())
            .transpose()
            .context("rows must be a number")?
            .unwrap_or(2);
        let cols: usize = props
            .get("cols")
            .map(|v| v.parse())
            .transpose()
            .context("cols must be a number")?
            .unwrap_or(2);
        if rows == 0 || cols == 0 || rows > 10_000 || cols > 500 {
            bail!("table size {rows}x{cols} out of range");
        }
        let mut tbl = XmlElement::new("w:tbl");
        let mut tblpr = XmlElement::new("w:tblPr");
        tblpr.push(el("w:tblW", &[("w:w", "0"), ("w:type", "auto")]));
        let mut borders = XmlElement::new("w:tblBorders");
        for side in ["top", "left", "bottom", "right", "insideH", "insideV"] {
            borders.push(el(
                &format!("w:{side}"),
                &[("w:val", "single"), ("w:sz", "4"), ("w:color", "auto")],
            ));
        }
        tblpr.push(borders);
        tbl.push(tblpr);
        let mut grid = XmlElement::new("w:tblGrid");
        for _ in 0..cols {
            grid.push(el("w:gridCol", &[("w:w", "2400")]));
        }
        tbl.push(grid);
        for _ in 0..rows {
            let mut tr = XmlElement::new("w:tr");
            for _ in 0..cols {
                let mut tc = XmlElement::new("w:tc");
                let mut tcpr = XmlElement::new("w:tcPr");
                tcpr.push(el("w:tcW", &[("w:w", "2400"), ("w:type", "dxa")]));
                tc.push(tcpr);
                tc.push(XmlElement::new("w:p"));
                tr.push(tc);
            }
            tbl.push(tr);
        }
        Ok(tbl)
    }

    fn insert_into(
        &mut self,
        parent_indices: &[usize],
        element: XmlElement,
        pos: &Position,
    ) -> Result<Vec<usize>> {
        // Compute the child insertion index.
        let insert_at = {
            let parent = self.node_at(parent_indices)?;
            match pos {
                Position::Append => {
                    // Keep trailing sectPr last in the body.
                    let mut at = parent.children.len();
                    if parent.local_name() == "body" {
                        if let Some(i) = parent
                            .children
                            .iter()
                            .position(|n| matches!(n, XmlNode::Element(e) if e.local_name() == "sectPr"))
                        {
                            at = i;
                        }
                    }
                    at
                }
                Position::Index(n) => {
                    // 0-based among element children; map to raw child index.
                    let mut seen = 0;
                    let mut at = parent.children.len();
                    for (i, node) in parent.children.iter().enumerate() {
                        if matches!(node, XmlNode::Element(_)) {
                            if seen == *n {
                                at = i;
                                break;
                            }
                            seen += 1;
                        }
                    }
                    at
                }
                Position::Before(p) | Position::After(p) => {
                    let anchor = self.resolve(&path::parse(p)?)?;
                    if anchor.len() != parent_indices.len() + 1
                        || anchor[..parent_indices.len()] != *parent_indices
                    {
                        bail!("anchor '{p}' is not a direct child of the target parent");
                    }
                    let base = anchor[parent_indices.len()];
                    if matches!(pos, Position::After(_)) {
                        base + 1
                    } else {
                        base
                    }
                }
            }
        };
        let parent = self.node_at_mut(parent_indices)?;
        parent.children.insert(insert_at, XmlNode::Element(element));
        let mut result = parent_indices.to_vec();
        result.push(insert_at);
        Ok(result)
    }

    // ------------------------------------------------------------- set ----

    fn apply_paragraph_props(&mut self, indices: &[usize], props: &Props) -> Result<()> {
        let p = self.node_at_mut(indices)?;
        if let Some(text) = props.get("text") {
            set_paragraph_text(p, text);
        }
        if props.has("style") || props.has("align") {
            let style = props.get("style").map(|s| s.to_string());
            let align = props.get("align").map(|s| s.to_string());
            let ppr = p.ensure_child("pPr", "w:pPr", true);
            if let Some(style) = style {
                let e = ppr.ensure_child("pStyle", "w:pStyle", true);
                e.set_attr("w:val", &style);
            }
            if let Some(align) = align {
                let (jc, _) = parse_align(&align)?;
                let e = ppr.ensure_child("jc", "w:jc", false);
                e.set_attr("w:val", jc);
            }
        }
        if has_run_format_props(props) {
            let p = self.node_at_mut(indices)?;
            let mut run_idxs = Vec::new();
            for (i, node) in p.children.iter().enumerate() {
                if matches!(node, XmlNode::Element(e) if e.local_name() == "r") {
                    run_idxs.push(i);
                }
            }
            for i in run_idxs {
                let run = p.children[i].as_element_mut().unwrap();
                apply_run_props(run, props)?;
            }
        }
        Ok(())
    }
}

fn find_child(parent: &XmlElement, seg: &Segment) -> Result<usize> {
    let local = ooxml_name(&seg.name)
        .map(|s| s.to_string())
        .unwrap_or_else(|| seg.name.clone());
    let mut matches: Vec<usize> = Vec::new();
    for (i, node) in parent.children.iter().enumerate() {
        if let XmlNode::Element(e) = node {
            if e.local_name() == local {
                let attr_ok = seg.preds.iter().all(|p| match p {
                    Predicate::Attr(k, v) => e.attr_local(k).map(|av| av == v).unwrap_or(false),
                    Predicate::Index(_) => true,
                });
                if attr_ok {
                    matches.push(i);
                }
            }
        }
    }
    let nth = seg.index().unwrap_or(1);
    matches.get(nth - 1).copied().with_context(|| {
        format!(
            "no element matches '{}[{}]' under <{}> ({} candidate(s))",
            seg.name,
            nth,
            parent.local_name(),
            matches.len()
        )
    })
}

fn paragraph_text(p: &XmlElement) -> String {
    let mut out = String::new();
    collect_para_text(p, &mut out);
    out
}

fn collect_para_text(e: &XmlElement, out: &mut String) {
    for child in e.elements() {
        match child.local_name() {
            "t" => out.push_str(&child.text_content()),
            "br" => out.push('\n'),
            "tab" => out.push('\t'),
            "rPr" | "pPr" => {}
            _ => collect_para_text(child, out),
        }
    }
}

fn run_text(r: &XmlElement) -> String {
    let mut out = String::new();
    collect_para_text(r, &mut out);
    out
}

fn describe_rpr(rpr: &XmlElement, info: &mut NodeInfo) {
    if rpr.child("b").is_some() {
        info.attr("bold", "true");
    }
    if rpr.child("i").is_some() {
        info.attr("italic", "true");
    }
    if rpr.child("u").is_some() {
        info.attr("underline", "true");
    }
    if let Some(sz) = rpr.child("sz").and_then(|s| s.attr_local("val")) {
        if let Ok(half) = sz.parse::<f64>() {
            info.attr("size", format!("{}pt", half / 2.0));
        }
    }
    if let Some(color) = rpr.child("color").and_then(|s| s.attr_local("val")) {
        info.attr("color", color);
    }
    if let Some(fonts) = rpr.child("rFonts").and_then(|s| s.attr_local("ascii")) {
        info.attr("font", fonts);
    }
}

fn has_run_format_props(props: &Props) -> bool {
    ["bold", "italic", "underline", "size", "color", "font", "highlight"]
        .iter()
        .any(|k| props.has(k) || props.has(&format!("font.{k}")))
}

/// Build a `w:r` from text + formatting props. Literal `\n` sequences in the
/// text become line breaks, `\t` become tabs.
fn build_run(text: &str, props: &Props) -> Result<XmlElement> {
    let mut r = XmlElement::new("w:r");
    let rpr = build_rpr(props)?;
    if let Some(rpr) = rpr {
        r.push(rpr);
    }
    push_text_nodes(&mut r, text);
    Ok(r)
}

/// Append w:t / w:br / w:tab children for a text that may contain escapes.
fn push_text_nodes(r: &mut XmlElement, text: &str) {
    let unescaped = text.replace("\\n", "\n").replace("\\t", "\t");
    let mut buf = String::new();
    for ch in unescaped.chars() {
        match ch {
            '\n' => {
                flush_text(r, &mut buf);
                r.push(XmlElement::new("w:br"));
            }
            '\t' => {
                flush_text(r, &mut buf);
                r.push(XmlElement::new("w:tab"));
            }
            c => buf.push(c),
        }
    }
    flush_text(r, &mut buf);
}

fn flush_text(r: &mut XmlElement, buf: &mut String) {
    if buf.is_empty() {
        return;
    }
    let mut t = XmlElement::new("w:t");
    t.set_attr("xml:space", "preserve");
    t.push_text(buf);
    r.push(t);
    buf.clear();
}

fn build_rpr(props: &Props) -> Result<Option<XmlElement>> {
    let mut rpr = XmlElement::new("w:rPr");
    let get = |k: &str| {
        props
            .get(k)
            .or_else(|| props.get(&format!("font.{k}")))
            .map(|s| s.to_string())
    };
    if let Some(font) = get("font").or_else(|| get("name")) {
        rpr.push(el(
            "w:rFonts",
            &[
                ("w:ascii", font.as_str()),
                ("w:hAnsi", font.as_str()),
                ("w:eastAsia", font.as_str()),
                ("w:cs", font.as_str()),
            ],
        ));
    }
    if let Some(v) = get("bold") {
        if parse_bool(&v)? {
            rpr.push(XmlElement::new("w:b"));
        }
    }
    if let Some(v) = get("italic") {
        if parse_bool(&v)? {
            rpr.push(XmlElement::new("w:i"));
        }
    }
    if let Some(v) = get("underline") {
        if parse_bool(&v)? {
            rpr.push(el("w:u", &[("w:val", "single")]));
        }
    }
    if let Some(v) = get("color") {
        rpr.push(el("w:color", &[("w:val", parse_color(&v)?.as_str())]));
    }
    if let Some(v) = get("size") {
        let half = (parse_pt(&v)? * 2.0).round() as i64;
        rpr.push(el("w:sz", &[("w:val", half.to_string().as_str())]));
        rpr.push(el("w:szCs", &[("w:val", half.to_string().as_str())]));
    }
    if let Some(v) = get("highlight") {
        rpr.push(el("w:highlight", &[("w:val", v.as_str())]));
    }
    if rpr.children.is_empty() {
        Ok(None)
    } else {
        Ok(Some(rpr))
    }
}

/// Apply formatting props to an existing run's rPr in place.
fn apply_run_props(run: &mut XmlElement, props: &Props) -> Result<()> {
    if let Some(text) = props.get("text") {
        // Replace all text-ish children, keep rPr.
        run.children.retain(
            |n| matches!(n, XmlNode::Element(e) if e.local_name() == "rPr"),
        );
        push_text_nodes(run, text);
    }
    let get = |k: &str| {
        props
            .get(k)
            .or_else(|| props.get(&format!("font.{k}")))
            .map(|s| s.to_string())
    };
    let needs_rpr = has_run_format_props(props);
    if !needs_rpr {
        return Ok(());
    }
    let rpr = run.ensure_child("rPr", "w:rPr", true);
    let toggle = |rpr: &mut XmlElement, local: &str, qname: &str, on: bool| {
        if on {
            rpr.ensure_child(local, qname, false);
        } else {
            rpr.children
                .retain(|n| !matches!(n, XmlNode::Element(e) if e.local_name() == local));
        }
    };
    if let Some(v) = get("bold") {
        toggle(rpr, "b", "w:b", parse_bool(&v)?);
    }
    if let Some(v) = get("italic") {
        toggle(rpr, "i", "w:i", parse_bool(&v)?);
    }
    if let Some(v) = get("underline") {
        if parse_bool(&v)? {
            rpr.ensure_child("u", "w:u", false).set_attr("w:val", "single");
        } else {
            rpr.children
                .retain(|n| !matches!(n, XmlNode::Element(e) if e.local_name() == "u"));
        }
    }
    if let Some(v) = get("color") {
        let color = parse_color(&v)?;
        rpr.ensure_child("color", "w:color", false).set_attr("w:val", &color);
    }
    if let Some(v) = get("size") {
        let half = (parse_pt(&v)? * 2.0).round() as i64;
        rpr.ensure_child("sz", "w:sz", false)
            .set_attr("w:val", &half.to_string());
        rpr.ensure_child("szCs", "w:szCs", false)
            .set_attr("w:val", &half.to_string());
    }
    if let Some(v) = get("font").or_else(|| get("name")) {
        let fonts = rpr.ensure_child("rFonts", "w:rFonts", false);
        for attr in ["w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"] {
            fonts.set_attr(attr, &v);
        }
    }
    if let Some(v) = get("highlight") {
        rpr.ensure_child("highlight", "w:highlight", false)
            .set_attr("w:val", &v);
    }
    Ok(())
}

fn set_paragraph_text(p: &mut XmlElement, text: &str) {
    // Preserve the first run's formatting if there is one.
    let first_rpr = p
        .elements()
        .find(|e| e.local_name() == "r")
        .and_then(|r| r.child("rPr"))
        .cloned();
    p.children.retain(|n| {
        matches!(n, XmlNode::Element(e) if e.local_name() == "pPr")
    });
    let mut run = XmlElement::new("w:r");
    if let Some(rpr) = first_rpr {
        run.push(rpr);
    }
    push_text_nodes(&mut run, text);
    p.push(run);
}

// ------------------------------------------------------- find / replace ----

/// Depth-first visit of every paragraph in a subtree (including paragraphs
/// nested in table cells).
fn for_each_paragraph<F>(e: &mut XmlElement, f: &mut F) -> Result<()>
where
    F: FnMut(&mut XmlElement) -> Result<()>,
{
    if e.local_name() == "p" {
        return f(e);
    }
    for child in e.children.iter_mut() {
        if let XmlNode::Element(ce) = child {
            for_each_paragraph(ce, f)?;
        }
    }
    Ok(())
}

/// A contiguous piece of paragraph text belonging to one `w:t` node.
struct TextPiece {
    /// Index path (within the paragraph) to the run element.
    run_idx: usize,
    /// Child index of the `w:t` inside the run.
    t_idx: usize,
    start: usize,
    len: usize,
}

pub(crate) fn build_find_regex(find: &str, props: &Props) -> Result<Regex> {
    let is_regex = props.get_bool("regex")?.unwrap_or(false);
    let pattern = if is_regex {
        find.to_string()
    } else {
        regex::escape(find)
    };
    Regex::new(&pattern).with_context(|| format!("invalid regex '{find}'"))
}

/// Collect the searchable text of a paragraph along with piece offsets.
/// Only top-level runs' `w:t` nodes participate; breaks count as '\n'.
fn paragraph_pieces(p: &XmlElement) -> (String, Vec<TextPiece>) {
    let mut joined = String::new();
    let mut pieces = Vec::new();
    for (run_idx, node) in p.children.iter().enumerate() {
        let XmlNode::Element(run) = node else { continue };
        if run.local_name() != "r" {
            continue;
        }
        for (t_idx, child) in run.children.iter().enumerate() {
            let XmlNode::Element(e) = child else { continue };
            match e.local_name() {
                "t" => {
                    let text = e.text_content();
                    pieces.push(TextPiece {
                        run_idx,
                        t_idx,
                        start: joined.len(),
                        len: text.len(),
                    });
                    joined.push_str(&text);
                }
                "br" => joined.push('\n'),
                "tab" => joined.push('\t'),
                _ => {}
            }
        }
    }
    (joined, pieces)
}

fn set_t_text(p: &mut XmlElement, run_idx: usize, t_idx: usize, text: &str) {
    if let Some(run) = p.children[run_idx].as_element_mut() {
        if let Some(t) = run.children[t_idx].as_element_mut() {
            t.children.clear();
            if !text.is_empty() {
                t.push_text(text);
            }
            // DrawingML's a:t allows no attributes; only WordprocessingML
            // text nodes take xml:space.
            if t.name.starts_with("w:") {
                t.set_attr("xml:space", "preserve");
            }
        }
    }
}

/// Replace all regex matches inside one paragraph. Returns match count.
pub(crate) fn replace_in_paragraph(p: &mut XmlElement, re: &Regex, replacement: &str) -> usize {
    let (joined, pieces) = paragraph_pieces(p);
    let matches: Vec<(usize, usize)> = re
        .find_iter(&joined)
        .map(|m| (m.start(), m.end()))
        .collect();
    if matches.is_empty() {
        return 0;
    }
    // Current text of each piece, edited as we go (right-to-left).
    let mut texts: Vec<String> = pieces
        .iter()
        .map(|pc| joined[pc.start..pc.start + pc.len].to_string())
        .collect();
    for &(s, e) in matches.iter().rev() {
        let mut first_hit: Option<usize> = None;
        for (i, pc) in pieces.iter().enumerate() {
            let ps = pc.start;
            let pe = pc.start + pc.len;
            if pe <= s || ps >= e {
                continue; // no overlap
            }
            let ls = s.saturating_sub(ps).min(pc.len);
            let le = (e - ps).min(pc.len);
            if first_hit.is_none() {
                first_hit = Some(i);
                let mut t = String::new();
                t.push_str(&texts[i][..ls]);
                t.push_str(replacement);
                t.push_str(&texts[i][le..]);
                texts[i] = t;
            } else {
                let mut t = String::new();
                t.push_str(&texts[i][..ls]);
                t.push_str(&texts[i][le..]);
                texts[i] = t;
            }
        }
    }
    for (i, pc) in pieces.iter().enumerate() {
        set_t_text(p, pc.run_idx, pc.t_idx, &texts[i]);
    }
    matches.len()
}

/// Apply formatting props to every match in a paragraph, splitting runs at
/// match boundaries. Returns match count.
fn format_in_paragraph(p: &mut XmlElement, re: &Regex, props: &Props) -> Result<usize> {
    let mut total = 0;
    // Re-scan after every applied match: splitting invalidates offsets.
    // `progress` guards against zero-width regex loops.
    let mut search_from = 0usize;
    loop {
        let (joined, pieces) = paragraph_pieces(p);
        let m = match re.find_at(&joined, search_from) {
            Some(m) if m.end() > m.start() => m,
            _ => break,
        };
        let (s, e) = (m.start(), m.end());
        // Collect (run_idx, t_idx, local_start, local_end) for affected pieces.
        let affected: Vec<(usize, usize, usize, usize)> = pieces
            .iter()
            .filter(|pc| pc.start + pc.len > s && pc.start < e)
            .map(|pc| {
                let ls = s.saturating_sub(pc.start).min(pc.len);
                let le = (e - pc.start).min(pc.len);
                (pc.run_idx, pc.t_idx, ls, le)
            })
            .collect();
        // Split runs right-to-left so earlier indices stay valid.
        for &(run_idx, t_idx, ls, le) in affected.iter().rev() {
            split_and_format_run(p, run_idx, t_idx, ls, le, props)?;
        }
        total += 1;
        search_from = e;
        if search_from > joined.len() {
            break;
        }
    }
    Ok(total)
}

/// Split the run at `run_idx` around `t_idx`'s [ls, le) range and apply
/// formatting to the matched middle part.
fn split_and_format_run(
    p: &mut XmlElement,
    run_idx: usize,
    t_idx: usize,
    ls: usize,
    le: usize,
    props: &Props,
) -> Result<()> {
    let run = p.children[run_idx]
        .as_element()
        .context("run vanished during split")?
        .clone();
    let t_text = run.children[t_idx]
        .as_element()
        .map(|t| t.text_content())
        .unwrap_or_default();
    let (pre, rest) = t_text.split_at(ls.min(t_text.len()));
    let (mid, post) = rest.split_at((le - ls).min(rest.len()));
    let rpr = run.child("rPr").cloned();

    let make_run = |text: &str| -> XmlElement {
        let mut r = XmlElement::new("w:r");
        if let Some(rpr) = &rpr {
            r.push(rpr.clone());
        }
        let mut t = XmlElement::new("w:t");
        t.set_attr("xml:space", "preserve");
        t.push_text(text);
        r.push(t);
        r
    };

    // Children of the original run before/after the split t node keep their
    // position: those before go to the "pre" run, those after to "post".
    let mut pre_run = XmlElement::new("w:r");
    let mut post_run = XmlElement::new("w:r");
    if let Some(rpr) = &rpr {
        pre_run.push(rpr.clone());
        post_run.push(rpr.clone());
    }
    for (i, child) in run.children.iter().enumerate() {
        if matches!(child, XmlNode::Element(e) if e.local_name() == "rPr") {
            continue;
        }
        if i < t_idx {
            pre_run.children.push(child.clone());
        } else if i > t_idx {
            post_run.children.push(child.clone());
        }
    }
    if !pre.is_empty() {
        let mut t = XmlElement::new("w:t");
        t.set_attr("xml:space", "preserve");
        t.push_text(pre);
        pre_run.push(t);
    }
    if !post.is_empty() {
        let mut t = XmlElement::new("w:t");
        t.set_attr("xml:space", "preserve");
        t.push_text(post);
        // post text goes before any trailing children? No: it comes first.
        let insert_at = if post_run.child("rPr").is_some() { 1 } else { 0 };
        post_run.children.insert(insert_at, XmlNode::Element(t));
    }
    let mut mid_run = make_run(mid);
    apply_run_props(&mut mid_run, props)?;

    let mut replacement: Vec<XmlNode> = Vec::new();
    let has_content = |r: &XmlElement| {
        r.elements().any(|e| e.local_name() != "rPr")
    };
    if has_content(&pre_run) {
        replacement.push(XmlNode::Element(pre_run));
    }
    replacement.push(XmlNode::Element(mid_run));
    if has_content(&post_run) {
        replacement.push(XmlNode::Element(post_run));
    }
    p.children.splice(run_idx..run_idx + 1, replacement);
    Ok(())
}

// ----------------------------------------------------------- Handler ----

impl Handler for Docx {
    fn view(&mut self, mode: &str) -> Result<Report> {
        let body = self.body()?;
        match mode {
            "text" => {
                let mut out = String::new();
                for element in body.elements() {
                    match element.local_name() {
                        "p" => {
                            out.push_str(&paragraph_text(element));
                            out.push('\n');
                        }
                        "tbl" => {
                            for tr in element.children_named("tr") {
                                let cells: Vec<String> = tr
                                    .children_named("tc")
                                    .iter()
                                    .map(|tc| tc.text_content())
                                    .collect();
                                out.push_str(&cells.join("\t"));
                                out.push('\n');
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Report::Text(out))
            }
            "outline" => {
                let mut out = String::new();
                let mut plain = 0usize;
                let flush_plain = |out: &mut String, plain: &mut usize| {
                    if *plain > 0 {
                        out.push_str(&format!("  ({} paragraph(s))\n", plain));
                        *plain = 0;
                    }
                };
                for element in body.elements() {
                    match element.local_name() {
                        "p" => {
                            let style = element
                                .child("pPr")
                                .and_then(|ppr| ppr.child("pStyle"))
                                .and_then(|s| s.attr_local("val"))
                                .unwrap_or("Normal");
                            if style.starts_with("Heading") || style == "Title" {
                                flush_plain(&mut out, &mut plain);
                                let level: usize = style
                                    .strip_prefix("Heading")
                                    .and_then(|n| n.parse().ok())
                                    .unwrap_or(1);
                                out.push_str(&"  ".repeat(level.saturating_sub(1)));
                                out.push_str(&format!("{}: {}\n", style, paragraph_text(element)));
                            } else if !paragraph_text(element).trim().is_empty() {
                                plain += 1;
                            }
                        }
                        "tbl" => {
                            flush_plain(&mut out, &mut plain);
                            let rows = element.children_named("tr").len();
                            let cols = element
                                .children_named("tr")
                                .first()
                                .map(|tr| tr.children_named("tc").len())
                                .unwrap_or(0);
                            out.push_str(&format!("  [table {rows}x{cols}]\n"));
                        }
                        _ => {}
                    }
                }
                flush_plain(&mut out, &mut plain);
                if out.is_empty() {
                    out.push_str("(empty document)\n");
                }
                Ok(Report::Text(out))
            }
            "stats" => {
                let mut paragraphs = 0usize;
                let mut tables = 0usize;
                let mut text = String::new();
                for element in body.elements() {
                    match element.local_name() {
                        "p" => {
                            paragraphs += 1;
                            text.push_str(&paragraph_text(element));
                            text.push('\n');
                        }
                        "tbl" => {
                            tables += 1;
                            text.push_str(&element.text_content());
                        }
                        _ => {}
                    }
                }
                let words = text.split_whitespace().count();
                let chars = text.chars().filter(|c| !c.is_whitespace()).count();
                Ok(Report::Data {
                    text: format!(
                        "paragraphs: {paragraphs}\ntables: {tables}\nwords: {words}\ncharacters: {chars}"
                    ),
                    data: json!({
                        "paragraphs": paragraphs,
                        "tables": tables,
                        "words": words,
                        "characters": chars,
                    }),
                })
            }
            "html" => {
                let mut out = String::new();
                for element in body.elements() {
                    render_html_block(element, &mut out);
                }
                Ok(Report::Text(crate::html::page("Document", &out)))
            }
            other => bail!("unknown view mode '{other}' for docx (text/outline/stats/html)"),
        }
    }

    fn get(&mut self, path_str: &str, depth: usize) -> Result<Report> {
        let dpath = path::parse(path_str)?;
        if dpath.is_root()
            || (dpath.segments.len() == 1 && dpath.segments[0].name.eq_ignore_ascii_case("body"))
        {
            let body = self.body()?;
            let info = self.node_info(body, "/body", depth.max(1));
            return Ok(Report::Nodes(vec![info]));
        }
        let indices = self.resolve(&dpath)?;
        let display = self.display_path(&indices)?;
        let node = self.node_at(&indices)?;
        let info = self.node_info(node, &display, depth);
        Ok(Report::Nodes(vec![info]))
    }

    fn add(&mut self, parent: &str, typ: &str, props: &Props, pos: &Position) -> Result<Report> {
        let dpath = path::parse(parent)?;
        let parent_indices = self.resolve(&dpath)?;
        let parent_local = self.node_at(&parent_indices)?.local_name().to_string();

        let element = match typ.to_ascii_lowercase().as_str() {
            "paragraph" | "p" | "para" => {
                if !matches!(parent_local.as_str(), "body" | "tc") {
                    bail!("paragraphs can be added to /body or a table cell, not <{parent_local}>");
                }
                self.build_paragraph(props)?
            }
            "run" | "r" => {
                if parent_local != "p" {
                    bail!("runs can only be added to a paragraph");
                }
                build_run(props.get("text").unwrap_or(""), props)?
            }
            "table" | "tbl" => {
                if !matches!(parent_local.as_str(), "body" | "tc") {
                    bail!("tables can be added to /body or a table cell");
                }
                self.build_table(props)?
            }
            "row" | "tr" => {
                if parent_local != "tbl" {
                    bail!("rows can only be added to a table");
                }
                let cols = {
                    let tbl = self.node_at(&parent_indices)?;
                    tbl.children_named("tr")
                        .first()
                        .map(|tr| tr.children_named("tc").len())
                        .unwrap_or(1)
                };
                let mut tr = XmlElement::new("w:tr");
                for _ in 0..cols {
                    let mut tc = XmlElement::new("w:tc");
                    let mut tcpr = XmlElement::new("w:tcPr");
                    tcpr.push(el("w:tcW", &[("w:w", "2400"), ("w:type", "dxa")]));
                    tc.push(tcpr);
                    tc.push(XmlElement::new("w:p"));
                    tr.push(tc);
                }
                tr
            }
            "break" | "pagebreak" => {
                let mut p = XmlElement::new("w:p");
                let mut r = XmlElement::new("w:r");
                r.push(el("w:br", &[("w:type", "page")]));
                p.push(r);
                p
            }
            "image" | "picture" => {
                if !matches!(parent_local.as_str(), "body" | "tc") {
                    bail!("images can be added to /body or a table cell");
                }
                self.build_image_paragraph(props)?
            }
            other => bail!(
                "unsupported docx element type '{other}' (paragraph/run/table/row/break/image)"
            ),
        };

        let new_indices = self.insert_into(&parent_indices, element, pos)?;
        let display = self.display_path(&new_indices)?;
        let node = self.node_at(&new_indices)?;
        let info = self.node_info(node, &display, 0);
        Ok(Report::Nodes(vec![info]))
    }

    fn set(
        &mut self,
        path_str: &str,
        props: &Props,
        find: Option<&str>,
        replace: Option<&str>,
    ) -> Result<Report> {
        let dpath = path::parse(path_str)?;

        if let Some(find) = find {
            let re = build_find_regex(find, props)?;
            // Determine scope: root = whole body, otherwise the element.
            let scope_indices = if dpath.is_root() {
                Vec::new()
            } else {
                self.resolve(&dpath)?
            };
            let mut matched = 0usize;
            {
                let scope = if scope_indices.is_empty() {
                    self.body_mut()?
                } else {
                    self.node_at_mut(&scope_indices)?
                };
                for_each_paragraph(scope, &mut |p| {
                    matched += match replace {
                        Some(rep) => replace_in_paragraph(p, &re, rep),
                        None => format_in_paragraph(p, &re, props)?,
                    };
                    Ok(())
                })?;
            }
            return Ok(Report::Data {
                text: format!("matched: {matched}"),
                data: json!({ "matched": matched }),
            });
        }

        if props.is_empty() {
            bail!("set requires --prop key=value (or --find/--replace)");
        }
        if dpath.is_root() {
            bail!("set on '/' requires --find/--replace; document-level props are not supported yet");
        }
        let indices = self.resolve(&dpath)?;
        let local = self.node_at(&indices)?.local_name().to_string();
        match local.as_str() {
            "p" => self.apply_paragraph_props(&indices, props)?,
            "r" => {
                let run = self.node_at_mut(&indices)?;
                apply_run_props(run, props)?;
            }
            "tc" => {
                if let Some(text) = props.get("text") {
                    let text = text.to_string();
                    let tc = self.node_at_mut(&indices)?;
                    tc.children.retain(|n| {
                        matches!(n, XmlNode::Element(e) if e.local_name() == "tcPr")
                    });
                    let mut p = XmlElement::new("w:p");
                    let mut r = XmlElement::new("w:r");
                    push_text_nodes(&mut r, &text);
                    p.push(r);
                    tc.push(p);
                }
                if has_run_format_props(props) {
                    let tc = self.node_at_mut(&indices)?;
                    let mut stack: Vec<&mut XmlElement> = vec![tc];
                    while let Some(e) = stack.pop() {
                        if e.local_name() == "r" {
                            apply_run_props(e, props)?;
                            continue;
                        }
                        for child in e.children.iter_mut() {
                            if let XmlNode::Element(ce) = child {
                                stack.push(ce);
                            }
                        }
                    }
                }
            }
            other => bail!("set is not supported on <{other}> (try paragraph, run, or cell paths)"),
        }
        let display = self.display_path(&indices)?;
        let node = self.node_at(&indices)?;
        let info = self.node_info(node, &display, 0);
        Ok(Report::Nodes(vec![info]))
    }

    fn remove(&mut self, path_str: &str) -> Result<Report> {
        let dpath = path::parse(path_str)?;
        if dpath.is_root() {
            bail!("cannot remove the document root");
        }
        let indices = self.resolve(&dpath)?;
        let display = self.display_path(&indices)?;
        let (parent_indices, last) = indices.split_at(indices.len() - 1);
        let parent = self.node_at_mut(parent_indices)?;
        parent.children.remove(last[0]);
        Ok(Report::Data {
            text: format!("removed {display}"),
            data: json!({ "removed": display }),
        })
    }

    fn validate(&mut self) -> Result<Report> {
        let mut problems = Vec::new();
        if self.body().is_err() {
            problems.push("document.xml has no <w:body>".to_string());
        }
        for part in ["[Content_Types].xml", "_rels/.rels", "word/styles.xml"] {
            if !self.pkg.has_part(part) {
                problems.push(format!("missing package part {part}"));
            }
        }
        let ok = problems.is_empty();
        Ok(Report::Data {
            text: if ok {
                "valid".to_string()
            } else {
                problems.join("\n")
            },
            data: json!({ "valid": ok, "problems": problems }),
        })
    }

    fn save(&mut self, path: &Path) -> Result<()> {
        self.pkg.put_xml(DOCUMENT_PART, &self.doc)?;
        self.pkg.put_xml(DOC_RELS_PART, &self.rels)?;
        self.pkg.save(path)
    }

    fn tree(&mut self) -> Result<Vec<NodeInfo>> {
        let body = self.body()?;
        Ok(vec![self.node_info(body, "/body", 8)])
    }

    fn move_el(&mut self, path_str: &str, to: Option<&str>, pos: &Position) -> Result<Report> {
        let indices = self.resolve(&path::parse(path_str)?)?;
        if indices.is_empty() {
            bail!("cannot move the document body");
        }
        let element = self.node_at(&indices)?.clone();
        let (src_parent, last) = indices.split_at(indices.len() - 1);
        self.node_at_mut(src_parent)?.children.remove(last[0]);
        // Anchors in `pos` are resolved after removal.
        let dest_parent = match to {
            Some(p) => self.resolve(&path::parse(p)?)?,
            None => src_parent.to_vec(),
        };
        let new_indices = self.insert_into(&dest_parent, element, pos)?;
        let display = self.display_path(&new_indices)?;
        let node = self.node_at(&new_indices)?;
        let info = self.node_info(node, &display, 0);
        Ok(Report::Nodes(vec![info]))
    }

    fn swap(&mut self, path1: &str, path2: &str) -> Result<Report> {
        let i1 = self.resolve(&path::parse(path1)?)?;
        let i2 = self.resolve(&path::parse(path2)?)?;
        if i1.is_empty() || i2.is_empty() {
            bail!("cannot swap the document body");
        }
        if i1.starts_with(&i2) || i2.starts_with(&i1) {
            bail!("cannot swap an element with its ancestor");
        }
        let n1 = self.node_at(&i1)?.clone();
        let n2 = self.node_at(&i2)?.clone();
        *self.node_at_mut(&i1)? = n2;
        *self.node_at_mut(&i2)? = n1;
        Ok(Report::Data {
            text: format!("swapped {path1} and {path2}"),
            data: json!({ "swapped": [path1, path2] }),
        })
    }

    fn dump(&mut self) -> Result<serde_json::Value> {
        let body = self.body()?;
        let mut ops = Vec::new();
        let mut tbl_count = 0usize;
        for element in body.elements() {
            match element.local_name() {
                "p" => {
                    let mut props = serde_json::Map::new();
                    props.insert("text".into(), json!(paragraph_text(element)));
                    if let Some(ppr) = element.child("pPr") {
                        if let Some(style) =
                            ppr.child("pStyle").and_then(|s| s.attr_local("val"))
                        {
                            props.insert("style".into(), json!(style));
                        }
                        if let Some(jc) = ppr.child("jc").and_then(|s| s.attr_local("val")) {
                            props.insert("align".into(), json!(jc));
                        }
                    }
                    // First run's formatting as a paragraph-level approximation.
                    if let Some(rpr) = element
                        .elements()
                        .find(|e| e.local_name() == "r")
                        .and_then(|r| r.child("rPr"))
                    {
                        let mut info = NodeInfo::new("", "run");
                        describe_rpr(rpr, &mut info);
                        for (k, v) in info.attributes {
                            props.insert(k, json!(v));
                        }
                    }
                    ops.push(json!({
                        "command": "add", "parent": "/body",
                        "type": "paragraph", "props": props,
                    }));
                }
                "tbl" => {
                    tbl_count += 1;
                    let rows = element.children_named("tr");
                    let cols = rows
                        .first()
                        .map(|tr| tr.children_named("tc").len())
                        .unwrap_or(0);
                    ops.push(json!({
                        "command": "add", "parent": "/body", "type": "table",
                        "props": { "rows": rows.len().to_string(), "cols": cols.to_string() },
                    }));
                    for (ri, tr) in rows.iter().enumerate() {
                        for (ci, tc) in tr.children_named("tc").iter().enumerate() {
                            let text = tc.text_content();
                            if text.is_empty() {
                                continue;
                            }
                            ops.push(json!({
                                "command": "set",
                                "path": format!("/body/tbl[{}]/tr[{}]/tc[{}]", tbl_count, ri + 1, ci + 1),
                                "props": { "text": text },
                            }));
                        }
                    }
                }
                _ => {}
            }
        }
        Ok(serde_json::Value::Array(ops))
    }
}

// -------------------------------------------------------------- html ----

fn render_html_block(element: &XmlElement, out: &mut String) {
    match element.local_name() {
        "p" => {
            let style = element
                .child("pPr")
                .and_then(|ppr| ppr.child("pStyle"))
                .and_then(|s| s.attr_local("val"))
                .unwrap_or("Normal");
            let (open, close) = match style {
                "Title" => ("<h1 class=\"title\">", "</h1>"),
                "Heading1" => ("<h1>", "</h1>"),
                "Heading2" => ("<h2>", "</h2>"),
                "Heading3" => ("<h3>", "</h3>"),
                _ => ("<p>", "</p>"),
            };
            let align = element
                .child("pPr")
                .and_then(|ppr| ppr.child("jc"))
                .and_then(|s| s.attr_local("val"));
            let open = match align {
                Some(a) if a != "left" => {
                    let css = if a == "both" { "justify" } else { a };
                    open.replace('>', &format!(" style=\"text-align:{css}\">"))
                }
                _ => open.to_string(),
            };
            out.push_str(&open);
            for run in element.elements().filter(|e| e.local_name() == "r") {
                render_html_run(run, out);
            }
            out.push_str(close);
            out.push('\n');
        }
        "tbl" => {
            out.push_str("<table>\n");
            for tr in element.children_named("tr") {
                out.push_str("<tr>");
                for tc in tr.children_named("tc") {
                    out.push_str("<td>");
                    for child in tc.elements() {
                        render_html_block(child, out);
                    }
                    out.push_str("</td>");
                }
                out.push_str("</tr>\n");
            }
            out.push_str("</table>\n");
        }
        _ => {}
    }
}

fn render_html_run(run: &XmlElement, out: &mut String) {
    let mut css = String::new();
    if let Some(rpr) = run.child("rPr") {
        if rpr.child("b").is_some() {
            css.push_str("font-weight:bold;");
        }
        if rpr.child("i").is_some() {
            css.push_str("font-style:italic;");
        }
        if rpr.child("u").is_some() {
            css.push_str("text-decoration:underline;");
        }
        if let Some(color) = rpr.child("color").and_then(|c| c.attr_local("val")) {
            if color != "auto" {
                css.push_str(&format!("color:#{color};"));
            }
        }
        if let Some(sz) = rpr.child("sz").and_then(|s| s.attr_local("val")) {
            if let Ok(half) = sz.parse::<f64>() {
                css.push_str(&format!("font-size:{}pt;", half / 2.0));
            }
        }
        if let Some(font) = rpr.child("rFonts").and_then(|f| f.attr_local("ascii")) {
            css.push_str(&format!("font-family:'{font}';"));
        }
    }
    let text = crate::html::escape(&run_text(run)).replace('\n', "<br>");
    if css.is_empty() {
        out.push_str(&text);
    } else {
        out.push_str(&format!("<span style=\"{css}\">{text}</span>"));
    }
}
