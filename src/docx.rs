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
const COMMENTS_PART: &str = "word/comments.xml";
const NUMBERING_PART: &str = "word/numbering.xml";
/// numId of the shared bullet / decimal list definitions we create.
const BULLET_NUM_ID: u32 = 1;
const DECIMAL_NUM_ID: u32 = 2;
const FOOTNOTES_PART: &str = "word/footnotes.xml";
const SETTINGS_PART: &str = "word/settings.xml";
const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

pub struct Docx {
    pkg: Package,
    doc: XmlElement,
    rels: XmlElement,
    /// word/numbering.xml, when the document has one (lists).
    numbering: Option<XmlElement>,
}

/// Map a friendly list kind to our shared numId (None = remove numbering).
fn list_num_id(kind: &str) -> Result<Option<u32>> {
    match kind.to_ascii_lowercase().as_str() {
        "bullet" | "bullets" | "ul" => Ok(Some(BULLET_NUM_ID)),
        "number" | "numbered" | "decimal" | "ol" => Ok(Some(DECIMAL_NUM_ID)),
        "none" => Ok(None),
        other => bail!("unknown list kind '{other}' (bullet/number/none)"),
    }
}

/// Build a `w:numPr` for a list level.
fn build_numpr(num_id: u32, level: u32) -> XmlElement {
    let mut numpr = XmlElement::new("w:numPr");
    numpr.push(el("w:ilvl", &[("w:val", level.to_string().as_str())]));
    numpr.push(el("w:numId", &[("w:val", num_id.to_string().as_str())]));
    numpr
}

/// (numId, level) of a paragraph's numbering, if any.
fn paragraph_numpr(p: &XmlElement) -> Option<(String, u32)> {
    let numpr = p.child("pPr")?.child("numPr")?;
    let num_id = numpr.child("numId")?.attr_local("val")?.to_string();
    let level = numpr
        .child("ilvl")
        .and_then(|i| i.attr_local("val"))
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    Some((num_id, level))
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
        let numbering = if pkg.has_part(NUMBERING_PART) {
            Some(pkg.xml(NUMBERING_PART)?)
        } else {
            None
        };
        Ok(Docx { pkg, doc, rels, numbering })
    }

    /// "bullet" or "number" for a numId, resolved through numbering.xml.
    fn list_kind_for_num(&self, num_id: &str) -> Option<&'static str> {
        let numbering = self.numbering.as_ref()?;
        let abs_id = numbering
            .children_named("num")
            .into_iter()
            .find(|n| n.attr_local("numId") == Some(num_id))?
            .child("abstractNumId")?
            .attr_local("val")?
            .to_string();
        let abs = numbering
            .children_named("abstractNum")
            .into_iter()
            .find(|a| a.attr_local("abstractNumId") == Some(abs_id.as_str()))?;
        let fmt = abs
            .children_named("lvl")
            .first()?
            .child("numFmt")?
            .attr_local("val")?;
        Some(if fmt == "bullet" { "bullet" } else { "number" })
    }

    fn build_image_paragraph(&mut self, props: &Props) -> Result<XmlElement> {
        let image = crate::media::image_from_props(props)?;
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
            let chain = find_child(current, seg)?;
            for &idx in &chain {
                current = current.children[idx].as_element().unwrap();
            }
            indices.extend(chain);
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
    /// sdt wrappers are invisible: elements inside them are numbered among
    /// their logical siblings.
    fn display_path(&self, indices: &[usize]) -> Result<String> {
        let mut out = String::from("/body");
        let mut parent = self.body()?;
        let mut rest: &[usize] = indices;
        while !rest.is_empty() {
            let entries = logical_children(parent);
            let hit = entries
                .iter()
                .position(|(chain, _)| rest.len() >= chain.len() && rest[..chain.len()] == chain[..])
                .context("internal: index chain does not match the document")?;
            let (chain, element) = &entries[hit];
            let local = element.local_name();
            let nth = entries[..hit]
                .iter()
                .filter(|(_, e)| e.local_name() == local)
                .count();
            out.push_str(&format!("/{}[{}]", local, nth + 1));
            rest = &rest[chain.len()..];
            parent = element;
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
                if let Some((num_id, level)) = paragraph_numpr(element) {
                    let kind = self.list_kind_for_num(&num_id).unwrap_or("list");
                    info.attr("list", kind);
                    if level > 0 {
                        info.attr("level", level.to_string());
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
            "hyperlink" => {
                info.text = Some(element.text_content());
                if let Some(url) = element
                    .attr("r:id")
                    .or_else(|| element.attr_local("id"))
                    .and_then(|rid| {
                        self.rels
                            .children_named("Relationship")
                            .into_iter()
                            .find(|r| r.attr_local("Id") == Some(rid))
                            .and_then(|r| r.attr_local("Target"))
                    })
                {
                    info.attr("url", url);
                }
            }
            _ => {}
        }
        if depth > 0 {
            let mut counts: std::collections::HashMap<String, usize> = Default::default();
            for (_, child) in logical_children(element) {
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
        if let Some(kind) = props.get("list") {
            if let Some(num_id) = list_num_id(kind)? {
                let level: u32 = props
                    .get("level")
                    .map(|v| v.parse())
                    .transpose()
                    .context("level must be a number (0-8)")?
                    .unwrap_or(0)
                    .min(8);
                ppr.push(build_numpr(num_id, level));
            }
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
                    // Anchors inside sdt wrappers resolve to longer chains;
                    // insertion happens beside the outermost wrapper.
                    let anchor = self.resolve(&path::parse(p)?)?;
                    if anchor.len() < parent_indices.len() + 1
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

    // ------------------------------------- comments / footnotes / fields ----

    /// Load a satellite part (comments/footnotes), creating it with `blank`
    /// plus its content-type override and document relationship when absent.
    fn ensure_satellite(
        &mut self,
        part: &'static str,
        blank: XmlElement,
        content_type: &str,
        rel_type: &str,
    ) -> Result<XmlElement> {
        if self.pkg.has_part(part) {
            return self.pkg.xml(part);
        }
        self.pkg.add_override(part, content_type)?;
        let target = part.strip_prefix("word/").unwrap_or(part).to_string();
        let has_rel = self
            .rels
            .children_named("Relationship")
            .into_iter()
            .any(|r| r.attr_local("Type") == Some(rel_type));
        if !has_rel {
            crate::media::add_relationship(&mut self.rels, rel_type, &target);
        }
        Ok(blank)
    }

    fn add_comment(&mut self, target: &[usize], props: &Props) -> Result<Report> {
        let text = props.get("text").context("comment needs --prop text=...")?.to_string();
        let author = props.get("author").unwrap_or("officecli").to_string();
        let initials: String = author
            .split_whitespace()
            .filter_map(|w| w.chars().next())
            .take(3)
            .collect::<String>()
            .to_uppercase();

        if self.node_at(target)?.local_name() != "p" {
            bail!("comments attach to a paragraph path like /body/p[2]");
        }

        let mut comments = self.ensure_satellite(
            COMMENTS_PART,
            el("w:comments", &[("xmlns:w", W_NS)]),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
            &format!("{REL_NS}/comments"),
        )?;
        let id = comments
            .children_named("comment")
            .into_iter()
            .filter_map(|c| c.attr_local("id").and_then(|v| v.parse::<u32>().ok()))
            .max()
            .map(|m| m + 1)
            .unwrap_or(1);
        let id_str = id.to_string();

        let mut comment = el(
            "w:comment",
            &[
                ("w:id", id_str.as_str()),
                ("w:author", author.as_str()),
                ("w:initials", if initials.is_empty() { "OC" } else { &initials }),
            ],
        );
        let mut p = XmlElement::new("w:p");
        let mut r = XmlElement::new("w:r");
        push_text_nodes(&mut r, &text);
        p.push(r);
        comment.push(p);
        comments.push(comment);
        self.pkg.put_xml(COMMENTS_PART, &comments)?;

        // Range markers + reference run on the target paragraph.
        let para = self.node_at_mut(target)?;
        let after_ppr = para
            .children
            .iter()
            .position(|n| !matches!(n, XmlNode::Element(e) if e.local_name() == "pPr"))
            .unwrap_or(para.children.len());
        para.children.insert(
            after_ppr,
            XmlNode::Element(el("w:commentRangeStart", &[("w:id", id_str.as_str())])),
        );
        para.push(el("w:commentRangeEnd", &[("w:id", id_str.as_str())]));
        let mut ref_run = XmlElement::new("w:r");
        ref_run.push(el("w:commentReference", &[("w:id", id_str.as_str())]));
        para.push(ref_run);

        let mut info = NodeInfo::new(format!("/comment[{id}]"), "comment");
        info.text = Some(text);
        info.attr("id", id_str);
        info.attr("author", author);
        Ok(Report::Nodes(vec![info]))
    }

    fn remove_comment(&mut self, id: u32) -> Result<Report> {
        if !self.pkg.has_part(COMMENTS_PART) {
            bail!("document has no comments");
        }
        let mut comments = self.pkg.xml(COMMENTS_PART)?;
        let id_str = id.to_string();
        let before = comments.children.len();
        comments.children.retain(|n| {
            !matches!(n, XmlNode::Element(e)
                if e.local_name() == "comment" && e.attr_local("id") == Some(id_str.as_str()))
        });
        if comments.children.len() == before {
            bail!("no comment with id {id}");
        }
        self.pkg.put_xml(COMMENTS_PART, &comments)?;
        // Strip markers and reference runs from the body.
        fn strip(e: &mut XmlElement, id: &str) {
            e.children.retain(|n| {
                let XmlNode::Element(c) = n else { return true };
                let is_marker = matches!(c.local_name(), "commentRangeStart" | "commentRangeEnd")
                    && c.attr_local("id") == Some(id);
                let is_ref_run = c.local_name() == "r"
                    && c.elements().any(|g| {
                        g.local_name() == "commentReference" && g.attr_local("id") == Some(id)
                    });
                !(is_marker || is_ref_run)
            });
            for child in e.children.iter_mut() {
                if let XmlNode::Element(c) = child {
                    strip(c, id);
                }
            }
        }
        strip(self.body_mut()?, &id_str);
        Ok(Report::Data {
            text: format!("removed comment {id}"),
            data: json!({ "removed": format!("/comment[{id}]") }),
        })
    }

    /// List of (id, author, text) comments.
    fn comments(&self) -> Result<Vec<(String, String, String)>> {
        if !self.pkg.has_part(COMMENTS_PART) {
            return Ok(Vec::new());
        }
        let comments = self.pkg.xml(COMMENTS_PART)?;
        Ok(comments
            .children_named("comment")
            .into_iter()
            .map(|c| {
                (
                    c.attr_local("id").unwrap_or("?").to_string(),
                    c.attr_local("author").unwrap_or("").to_string(),
                    c.text_content(),
                )
            })
            .collect())
    }

    fn add_footnote(&mut self, target: &[usize], props: &Props) -> Result<Report> {
        let text = props.get("text").context("footnote needs --prop text=...")?.to_string();
        if self.node_at(target)?.local_name() != "p" {
            bail!("footnotes attach to a paragraph path like /body/p[2]");
        }
        let mut blank = el("w:footnotes", &[("xmlns:w", W_NS)]);
        for (typ, id, marker) in [
            ("separator", "-1", "w:separator"),
            ("continuationSeparator", "0", "w:continuationSeparator"),
        ] {
            let mut fnote = el("w:footnote", &[("w:type", typ), ("w:id", id)]);
            let mut p = XmlElement::new("w:p");
            let mut r = XmlElement::new("w:r");
            r.push(XmlElement::new(marker));
            p.push(r);
            fnote.push(p);
            blank.push(fnote);
        }
        let mut footnotes = self.ensure_satellite(
            FOOTNOTES_PART,
            blank,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            &format!("{REL_NS}/footnotes"),
        )?;
        let id = footnotes
            .children_named("footnote")
            .into_iter()
            .filter_map(|c| c.attr_local("id").and_then(|v| v.parse::<i32>().ok()))
            .max()
            .map(|m| m.max(0) + 1)
            .unwrap_or(1);
        let id_str = id.to_string();

        let mut fnote = el("w:footnote", &[("w:id", id_str.as_str())]);
        let mut p = XmlElement::new("w:p");
        let mut ref_run = XmlElement::new("w:r");
        let mut rpr = XmlElement::new("w:rPr");
        rpr.push(el("w:vertAlign", &[("w:val", "superscript")]));
        ref_run.push(rpr);
        ref_run.push(XmlElement::new("w:footnoteRef"));
        p.push(ref_run);
        let mut text_run = XmlElement::new("w:r");
        push_text_nodes(&mut text_run, &format!(" {text}"));
        p.push(text_run);
        fnote.push(p);
        footnotes.push(fnote);
        self.pkg.put_xml(FOOTNOTES_PART, &footnotes)?;

        // Superscript reference in the body paragraph.
        let para = self.node_at_mut(target)?;
        let mut r = XmlElement::new("w:r");
        let mut rpr = XmlElement::new("w:rPr");
        rpr.push(el("w:vertAlign", &[("w:val", "superscript")]));
        r.push(rpr);
        r.push(el("w:footnoteReference", &[("w:id", id_str.as_str())]));
        para.push(r);

        let mut info = NodeInfo::new(format!("/footnote[{id}]"), "footnote");
        info.text = Some(text);
        info.attr("id", id_str);
        Ok(Report::Nodes(vec![info]))
    }

    /// Complex field runs: begin(dirty) + instrText + separate + placeholder + end.
    fn build_field_runs(instr: &str, placeholder: &str) -> Vec<XmlElement> {
        let mut runs = Vec::new();
        let mut begin = XmlElement::new("w:r");
        begin.push(el("w:fldChar", &[("w:fldCharType", "begin"), ("w:dirty", "true")]));
        runs.push(begin);
        let mut instr_r = XmlElement::new("w:r");
        let mut it = el("w:instrText", &[("xml:space", "preserve")]);
        it.push_text(&format!(" {instr} "));
        instr_r.push(it);
        runs.push(instr_r);
        let mut sep = XmlElement::new("w:r");
        sep.push(el("w:fldChar", &[("w:fldCharType", "separate")]));
        runs.push(sep);
        if !placeholder.is_empty() {
            let mut text_r = XmlElement::new("w:r");
            let mut t = XmlElement::new("w:t");
            t.set_attr("xml:space", "preserve");
            t.push_text(placeholder);
            text_r.push(t);
            runs.push(text_r);
        }
        let mut end = XmlElement::new("w:r");
        end.push(el("w:fldChar", &[("w:fldCharType", "end")]));
        runs.push(end);
        runs
    }

    /// Field instruction from a friendly name or raw code.
    fn field_instr(props: &Props) -> Result<String> {
        if let Some(code) = props.get("code").or_else(|| props.get("instr")) {
            return Ok(code.to_string());
        }
        let kind = props
            .get("kind")
            .context("field needs --prop kind=page|numpages|date|time|filename|author or --prop code=\"...\"")?;
        Ok(match kind.to_ascii_lowercase().as_str() {
            "page" => "PAGE".to_string(),
            "numpages" | "pages" => "NUMPAGES".to_string(),
            "date" => "DATE".to_string(),
            "time" => "TIME".to_string(),
            "filename" => "FILENAME".to_string(),
            "author" => "AUTHOR".to_string(),
            other => bail!("unknown field kind '{other}' (page/numpages/date/time/filename/author, or use code=)"),
        })
    }

    fn add_field(&mut self, target: &[usize], props: &Props) -> Result<Report> {
        if self.node_at(target)?.local_name() != "p" {
            bail!("fields are added to a paragraph path like /body/p[2]");
        }
        let instr = Self::field_instr(props)?;
        let para = self.node_at_mut(target)?;
        for run in Self::build_field_runs(&instr, "") {
            para.push(run);
        }
        let display = self.display_path(target)?;
        let mut info = NodeInfo::new(&display, "field");
        info.attr("code", instr);
        Ok(Report::Nodes(vec![info]))
    }

    /// Make sure word/numbering.xml exists with the shared bullet (numId 1)
    /// and decimal (numId 2) definitions.
    fn ensure_numbering(&mut self) -> Result<()> {
        if self.pkg.has_part(NUMBERING_PART) {
            return Ok(());
        }
        let mut numbering = el("w:numbering", &[("xmlns:w", W_NS)]);
        for (abstract_id, bullet) in [(0u32, true), (1u32, false)] {
            let mut abs = el(
                "w:abstractNum",
                &[("w:abstractNumId", abstract_id.to_string().as_str())],
            );
            abs.push(el("w:multiLevelType", &[("w:val", "hybridMultilevel")]));
            for ilvl in 0..9u32 {
                let mut lvl = el("w:lvl", &[("w:ilvl", ilvl.to_string().as_str())]);
                lvl.push(el("w:start", &[("w:val", "1")]));
                if bullet {
                    lvl.push(el("w:numFmt", &[("w:val", "bullet")]));
                    lvl.push(el("w:lvlText", &[("w:val", "•")]));
                } else {
                    lvl.push(el("w:numFmt", &[("w:val", "decimal")]));
                    lvl.push(el(
                        "w:lvlText",
                        &[("w:val", format!("%{}.", ilvl + 1).as_str())],
                    ));
                }
                lvl.push(el("w:lvlJc", &[("w:val", "left")]));
                let mut ppr = XmlElement::new("w:pPr");
                let left = 720 * (ilvl + 1);
                ppr.push(el(
                    "w:ind",
                    &[("w:left", left.to_string().as_str()), ("w:hanging", "360")],
                ));
                lvl.push(ppr);
                abs.push(lvl);
            }
            numbering.push(abs);
        }
        for (num_id, abstract_id) in [(BULLET_NUM_ID, 0u32), (DECIMAL_NUM_ID, 1u32)] {
            let mut num = el("w:num", &[("w:numId", num_id.to_string().as_str())]);
            num.push(el(
                "w:abstractNumId",
                &[("w:val", abstract_id.to_string().as_str())],
            ));
            numbering.push(num);
        }
        self.pkg.add_override(
            NUMBERING_PART,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        )?;
        let rel_type = format!("{REL_NS}/numbering");
        let has_rel = self
            .rels
            .children_named("Relationship")
            .into_iter()
            .any(|r| r.attr_local("Type") == Some(rel_type.as_str()));
        if !has_rel {
            crate::media::add_relationship(&mut self.rels, &rel_type, "numbering.xml");
        }
        self.pkg.put_xml(NUMBERING_PART, &numbering)
    }

    /// "• " / "3. " prefix (with level indent) for a list paragraph.
    fn list_marker(
        &self,
        p: &XmlElement,
        counters: &mut std::collections::HashMap<(String, u32), u32>,
    ) -> Option<String> {
        let (num_id, level) = paragraph_numpr(p)?;
        let indent = "  ".repeat(level as usize);
        let marker = match self.list_kind_for_num(&num_id) {
            Some("number") => {
                let c = counters.entry((num_id, level)).or_insert(0);
                *c += 1;
                format!("{c}.")
            }
            _ => "•".to_string(),
        };
        Some(format!("{indent}{marker} "))
    }

    /// Add a whole list: one paragraph per `\n`-separated item, leading
    /// tabs selecting deeper levels.
    fn add_list(&mut self, parent: &[usize], props: &Props, pos: &Position) -> Result<Report> {
        let parent_local = self.node_at(parent)?.local_name().to_string();
        if !matches!(parent_local.as_str(), "body" | "tc") {
            bail!("lists are added to /body or a table cell");
        }
        let items_raw = props
            .get("items")
            .context(r#"list needs --prop items="First\nSecond\n\tNested" (\n separates items, leading \t = deeper level)"#)?
            .replace("\\n", "\n")
            .replace("\\t", "\t");
        let kind = props.get("kind").unwrap_or("bullet").to_string();
        if list_num_id(&kind)?.is_none() {
            bail!("list kind cannot be 'none'");
        }
        self.ensure_numbering()?;
        let items: Vec<(usize, String)> = items_raw
            .split('\n')
            .filter(|l| !l.trim().is_empty())
            .map(|l| {
                let level = l.chars().take_while(|c| *c == '\t').count().min(8);
                (level, l.trim_start_matches('\t').trim_end().to_string())
            })
            .collect();
        if items.is_empty() {
            bail!("list has no items");
        }
        // Per-item props: text/list/level plus pass-through formatting.
        let passthrough: Vec<(String, String)> = props
            .iter()
            .filter(|(k, _)| {
                !matches!(
                    k.to_ascii_lowercase().as_str(),
                    "items" | "kind" | "list" | "level" | "text"
                )
            })
            .map(|(k, v)| (k.to_string(), v.to_string()))
            .collect();
        let build = |me: &Self, level: usize, text: &str| -> Result<XmlElement> {
            let mut pairs = vec![
                ("text".to_string(), text.to_string()),
                ("list".to_string(), kind.clone()),
                ("level".to_string(), level.to_string()),
            ];
            pairs.extend(passthrough.iter().cloned());
            me.build_paragraph(&Props::from_pairs(pairs))
        };
        match pos {
            Position::Append => {
                for (level, text) in &items {
                    let p = build(self, *level, text)?;
                    self.insert_into(parent, p, pos)?;
                }
            }
            // Inserting at a fixed anchor: go in reverse so the final
            // document order matches the input order.
            _ => {
                for (level, text) in items.iter().rev() {
                    let p = build(self, *level, text)?;
                    self.insert_into(parent, p, pos)?;
                }
            }
        }
        Ok(Report::Data {
            text: format!("added {} {kind} list item(s)", items.len()),
            data: json!({ "items": items.len(), "kind": kind }),
        })
    }

    /// Get (or create at the end of the body) the document-level sectPr.
    fn sectpr_mut(&mut self) -> Result<&mut XmlElement> {
        let body = self.body_mut()?;
        Ok(body.ensure_child("sectPr", "w:sectPr", false))
    }

    /// Document-level page setup: orientation, page-size, margins.
    fn apply_page_setup(&mut self, props: &Props) -> Result<Vec<String>> {
        let mut changed = Vec::new();
        // Sizes in twips (1/20 pt): EMU / 635.
        let to_twips = |v: &str| -> Result<i64> { Ok(crate::props::parse_emu(v)? / 635) };

        if let Some(size) = props.get("page-size").or_else(|| props.get("pagesize")) {
            let (w, h) = match size.to_ascii_lowercase().as_str() {
                "letter" => (12_240, 15_840),
                "a4" => (11_906, 16_838),
                "a3" => (16_838, 23_811),
                "legal" => (12_240, 20_160),
                custom => match custom.split_once('x') {
                    Some((w, h)) => (to_twips(w)?, to_twips(h)?),
                    None => bail!("page-size is letter/a4/a3/legal or WxH (e.g. 8.5inx11in)"),
                },
            };
            let sectpr = self.sectpr_mut()?;
            let pgsz = sectpr.ensure_child("pgSz", "w:pgSz", true);
            pgsz.set_attr("w:w", &w.to_string());
            pgsz.set_attr("w:h", &h.to_string());
            changed.push(format!("page-size={size}"));
        }
        if let Some(orientation) = props.get("orientation") {
            let landscape = match orientation.to_ascii_lowercase().as_str() {
                "landscape" => true,
                "portrait" => false,
                other => bail!("orientation is portrait or landscape, got '{other}'"),
            };
            let sectpr = self.sectpr_mut()?;
            let pgsz = sectpr.ensure_child("pgSz", "w:pgSz", true);
            let w: i64 = pgsz.attr("w:w").and_then(|v| v.parse().ok()).unwrap_or(12_240);
            let h: i64 = pgsz.attr("w:h").and_then(|v| v.parse().ok()).unwrap_or(15_840);
            let currently_landscape = pgsz.attr("w:orient") == Some("landscape");
            if landscape != currently_landscape {
                pgsz.set_attr("w:w", &h.to_string());
                pgsz.set_attr("w:h", &w.to_string());
            }
            if landscape {
                pgsz.set_attr("w:orient", "landscape");
            } else {
                pgsz.remove_attr("w:orient");
            }
            changed.push(format!("orientation={orientation}"));
        }
        let uniform = props.get("margins").or_else(|| props.get("margin"));
        let sides = [
            ("margin-top", "w:top"),
            ("margin-right", "w:right"),
            ("margin-bottom", "w:bottom"),
            ("margin-left", "w:left"),
        ];
        if uniform.is_some() || sides.iter().any(|(k, _)| props.has(k)) {
            let uniform_twips = uniform.map(to_twips).transpose()?;
            let mut values = Vec::new();
            for (key, attr) in sides {
                let v = match props.get(key) {
                    Some(v) => Some(to_twips(v)?),
                    None => uniform_twips,
                };
                values.push((attr, v));
            }
            let sectpr = self.sectpr_mut()?;
            let pgmar = sectpr.ensure_child("pgMar", "w:pgMar", false);
            for (attr, v) in values {
                if let Some(v) = v {
                    pgmar.set_attr(attr, &v.to_string());
                }
            }
            changed.push("margins".to_string());
        }
        Ok(changed)
    }

    /// Add (or replace) the default header or footer.
    fn add_header_footer(&mut self, kind: &str, props: &Props) -> Result<Report> {
        let is_header = kind == "header";
        let (root_name, rel_kind, ct) = if is_header {
            (
                "w:hdr",
                "header",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            )
        } else {
            (
                "w:ftr",
                "footer",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
            )
        };
        let rel_type = format!("{REL_NS}/{rel_kind}");
        let ref_name = if is_header { "headerReference" } else { "footerReference" };
        let ref_qname = format!("w:{ref_name}");

        // Replace any existing default reference (and its part).
        self.remove_header_footer(kind).ok();

        let n = self
            .pkg
            .part_names()
            .filter_map(|p| {
                p.strip_prefix(&format!("word/{rel_kind}"))
                    .and_then(|s| s.strip_suffix(".xml"))
                    .and_then(|x| x.parse::<u32>().ok())
            })
            .max()
            .unwrap_or(0)
            + 1;
        let part = format!("word/{rel_kind}{n}.xml");

        let mut root = el(root_name, &[("xmlns:w", W_NS), ("xmlns:r", REL_NS)]);
        let mut p = XmlElement::new("w:p");
        let mut ppr = XmlElement::new("w:pPr");
        let align = props.get("align").unwrap_or(if is_header { "left" } else { "center" });
        let (jc, _) = parse_align(align)?;
        ppr.push(el("w:jc", &[("w:val", jc)]));
        p.push(ppr);
        if let Some(text) = props.get("text") {
            p.push(build_run(text, props)?);
        }
        if props.get_bool("page-numbers")?.unwrap_or(false)
            || props.get_bool("pagenumbers")?.unwrap_or(false)
        {
            let push_literal = |p: &mut XmlElement, text: &str| {
                let mut r = XmlElement::new("w:r");
                let mut t = XmlElement::new("w:t");
                t.set_attr("xml:space", "preserve");
                t.push_text(text);
                r.push(t);
                p.push(r);
            };
            if props.has("text") {
                push_literal(&mut p, " — ");
            }
            push_literal(&mut p, "Page ");
            for run in Self::build_field_runs("PAGE", "1") {
                p.push(run);
            }
            push_literal(&mut p, " of ");
            for run in Self::build_field_runs("NUMPAGES", "1") {
                p.push(run);
            }
        }
        root.push(p);
        self.pkg.put_xml(&part, &root)?;
        self.pkg.add_override(&part, ct)?;
        let target = part.strip_prefix("word/").unwrap_or(&part).to_string();
        let rid = crate::media::add_relationship(&mut self.rels, &rel_type, &target);

        // References go first inside sectPr.
        let sectpr = self.sectpr_mut()?;
        let mut reference = el(&ref_qname, &[("w:type", "default")]);
        reference.set_attr("r:id", &rid);
        sectpr.children.insert(0, XmlNode::Element(reference));

        let mut info = NodeInfo::new(format!("/{kind}"), kind);
        info.text = props.get("text").map(|s| s.to_string());
        Ok(Report::Nodes(vec![info]))
    }

    /// Remove the default header or footer (reference, rel, and part).
    fn remove_header_footer(&mut self, kind: &str) -> Result<Report> {
        let is_header = kind == "header";
        let ref_name = if is_header { "headerReference" } else { "footerReference" };
        let rid = self
            .body()?
            .child("sectPr")
            .and_then(|s| {
                s.children_named(ref_name)
                    .into_iter()
                    .find(|r| r.attr_local("type") == Some("default"))
            })
            .and_then(|r| r.attr("r:id").or_else(|| r.attr_local("id")))
            .map(|s| s.to_string())
            .with_context(|| format!("document has no default {kind}"))?;
        let target = self
            .rels
            .children_named("Relationship")
            .into_iter()
            .find(|r| r.attr_local("Id") == Some(rid.as_str()))
            .and_then(|r| r.attr_local("Target"))
            .map(|t| format!("word/{}", t.trim_start_matches("./")));
        if let Some(sectpr) = self.body_mut()?.child_mut("sectPr") {
            sectpr.children.retain(|n| {
                !matches!(n, XmlNode::Element(e)
                    if e.local_name() == ref_name
                        && e.attr_local("type") == Some("default"))
            });
        }
        self.rels.children.retain(|n| {
            !matches!(n, XmlNode::Element(e)
                if e.local_name() == "Relationship" && e.attr_local("Id") == Some(rid.as_str()))
        });
        if let Some(part) = target {
            let mut ct = self.pkg.xml("[Content_Types].xml")?;
            let part_name = format!("/{part}");
            ct.children.retain(|n| {
                !matches!(n, XmlNode::Element(e)
                    if e.local_name() == "Override"
                        && e.attr_local("PartName") == Some(&part_name))
            });
            self.pkg.put_xml("[Content_Types].xml", &ct)?;
            self.pkg.remove_part(&part);
        }
        Ok(Report::Data {
            text: format!("removed the default {kind}"),
            data: json!({ "removed": format!("/{kind}") }),
        })
    }

    /// Make sure word/settings.xml exists and asks Word to update fields on
    /// open (used by TOC so it populates itself).
    fn ensure_update_fields(&mut self) -> Result<()> {
        let mut settings = self.ensure_satellite(
            SETTINGS_PART,
            el("w:settings", &[("xmlns:w", W_NS)]),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
            &format!("{REL_NS}/settings"),
        )?;
        settings
            .ensure_child("updateFields", "w:updateFields", true)
            .set_attr("w:val", "true");
        self.pkg.put_xml(SETTINGS_PART, &settings)
    }

    fn build_toc_paragraph(&self, props: &Props) -> XmlElement {
        let levels = props.get("levels").unwrap_or("1-3");
        let instr = format!("TOC \\o \"{levels}\" \\h \\z \\u");
        let mut p = XmlElement::new("w:p");
        for run in Self::build_field_runs(
            &instr,
            "Table of contents (open in Word and update the field, or print/export, to populate).",
        ) {
            p.push(run);
        }
        p
    }

    // ------------------------------------------------------------ dump ----

    /// If the paragraph is an embedded picture, return `add image` props
    /// with the bytes base64-encoded as `srcdata`.
    fn dump_image_paragraph(
        &self,
        p: &XmlElement,
    ) -> Result<Option<serde_json::Map<String, serde_json::Value>>> {
        let Some(blip) = find_descendant(p, "blip") else {
            return Ok(None);
        };
        let Some(rid) = blip.attr("r:embed").or_else(|| blip.attr_local("embed")) else {
            return Ok(None);
        };
        let Some(target) = self
            .rels
            .children_named("Relationship")
            .into_iter()
            .find(|r| r.attr_local("Id") == Some(rid))
            .and_then(|r| r.attr_local("Target"))
        else {
            return Ok(None);
        };
        let part = format!("word/{}", target.trim_start_matches("./"));
        let Ok(bytes) = self.pkg.raw(&part) else {
            return Ok(None);
        };
        use base64::Engine;
        let mut props = serde_json::Map::new();
        props.insert(
            "srcdata".into(),
            json!(base64::engine::general_purpose::STANDARD.encode(bytes)),
        );
        if let Some(extent) = find_descendant(p, "extent") {
            if let (Some(cx), Some(cy)) = (extent.attr_local("cx"), extent.attr_local("cy")) {
                props.insert("w".into(), json!(cx));
                props.insert("h".into(), json!(cy));
            }
        }
        Ok(Some(props))
    }

    // ------------------------------------------------------------- set ----

    fn apply_paragraph_props(&mut self, indices: &[usize], props: &Props) -> Result<()> {
        let p = self.node_at_mut(indices)?;
        if let Some(text) = props.get("text") {
            set_paragraph_text(p, text);
        }
        if props.has("style") || props.has("align") || props.has("list") {
            let style = props.get("style").map(|s| s.to_string());
            let align = props.get("align").map(|s| s.to_string());
            let list = props
                .get("list")
                .map(|kind| {
                    let level: u32 = props
                        .get("level")
                        .map(|v| v.parse())
                        .transpose()
                        .context("level must be a number (0-8)")?
                        .unwrap_or(0)
                        .min(8);
                    Ok::<_, anyhow::Error>((list_num_id(kind)?, level))
                })
                .transpose()?;
            let p = self.node_at_mut(indices)?;
            let ppr = p.ensure_child("pPr", "w:pPr", true);
            if let Some(style) = style {
                let e = ppr.ensure_child("pStyle", "w:pStyle", true);
                e.set_attr("w:val", &style);
            }
            if let Some((num_id, level)) = list {
                ppr.children.retain(|n| {
                    !matches!(n, XmlNode::Element(e) if e.local_name() == "numPr")
                });
                if let Some(num_id) = num_id {
                    // numPr follows pStyle in the pPr schema order.
                    let at = ppr
                        .children
                        .iter()
                        .position(|n| {
                            !matches!(n, XmlNode::Element(e) if e.local_name() == "pStyle")
                        })
                        .unwrap_or(ppr.children.len());
                    ppr.children
                        .insert(at, XmlNode::Element(build_numpr(num_id, level)));
                }
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

/// Logical children of `parent`, looking straight through `w:sdt` content
/// controls (Word wraps arbitrary blocks in them). Each entry is the raw
/// `children`-index chain from `parent` to the element.
fn logical_children(parent: &XmlElement) -> Vec<(Vec<usize>, &XmlElement)> {
    fn walk<'a>(
        e: &'a XmlElement,
        prefix: &mut Vec<usize>,
        out: &mut Vec<(Vec<usize>, &'a XmlElement)>,
    ) {
        for (i, node) in e.children.iter().enumerate() {
            let XmlNode::Element(c) = node else { continue };
            if c.local_name() == "sdt" {
                if let Some(ci) = c.children.iter().position(|n| {
                    matches!(n, XmlNode::Element(g) if g.local_name() == "sdtContent")
                }) {
                    prefix.push(i);
                    prefix.push(ci);
                    walk(c.children[ci].as_element().unwrap(), prefix, out);
                    prefix.pop();
                    prefix.pop();
                }
                continue;
            }
            let mut chain = prefix.clone();
            chain.push(i);
            out.push((chain, c));
        }
    }
    let mut out = Vec::new();
    walk(parent, &mut Vec::new(), &mut out);
    out
}

/// Resolve one path segment to the raw index chain of the matching logical
/// child (which may sit inside sdt wrappers).
fn find_child(parent: &XmlElement, seg: &Segment) -> Result<Vec<usize>> {
    let local = ooxml_name(&seg.name)
        .map(|s| s.to_string())
        .unwrap_or_else(|| seg.name.clone());
    let mut matches: Vec<Vec<usize>> = Vec::new();
    for (chain, e) in logical_children(parent) {
        if e.local_name() == local {
            let attr_ok = seg.preds.iter().all(|p| match p {
                Predicate::Attr(k, v) => e.attr_local(k).map(|av| av == v).unwrap_or(false),
                Predicate::Index(_) => true,
            });
            if attr_ok {
                matches.push(chain);
            }
        }
    }
    let nth = seg.index().unwrap_or(1);
    let count = matches.len();
    matches.into_iter().nth(nth - 1).with_context(|| {
        format!(
            "no element matches '{}[{}]' under <{}> ({} candidate(s))",
            seg.name,
            nth,
            parent.local_name(),
            count
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

/// First descendant element with the given local name (depth-first).
fn find_descendant<'a>(e: &'a XmlElement, local: &str) -> Option<&'a XmlElement> {
    for child in e.elements() {
        if child.local_name() == local {
            return Some(child);
        }
        if let Some(found) = find_descendant(child, local) {
            return Some(found);
        }
    }
    None
}

/// Run formatting as replayable prop pairs.
fn rpr_props(rpr: &XmlElement) -> Vec<(String, String)> {
    let mut info = NodeInfo::new("", "run");
    describe_rpr(rpr, &mut info);
    info.attributes
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
                let mut counters: std::collections::HashMap<(String, u32), u32> = Default::default();
                for (_, element) in logical_children(&body) {
                    match element.local_name() {
                        "p" => {
                            if let Some(marker) = self.list_marker(element, &mut counters) {
                                out.push_str(&marker);
                            }
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
                for (_, element) in logical_children(&body) {
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
                for (_, element) in logical_children(&body) {
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
                for (_, element) in logical_children(&body) {
                    render_html_block(element, &mut out);
                }
                Ok(Report::Text(crate::html::page("Document", &out)))
            }
            "comments" => {
                let comments = self.comments()?;
                let mut nodes = Vec::new();
                for (id, author, text) in comments {
                    let mut info = NodeInfo::new(format!("/comment[{id}]"), "comment");
                    info.text = Some(text);
                    info.attr("id", id);
                    info.attr("author", author);
                    nodes.push(info);
                }
                Ok(Report::Nodes(nodes))
            }
            other => bail!("unknown view mode '{other}' for docx (text/outline/stats/html/comments)"),
        }
    }

    fn get(&mut self, path_str: &str, depth: usize, _computed: bool) -> Result<Report> {
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

        // Types that attach to an existing element rather than inserting a
        // new sibling.
        match typ.to_ascii_lowercase().as_str() {
            "comment" => return self.add_comment(&parent_indices, props),
            "footnote" => return self.add_footnote(&parent_indices, props),
            "field" => return self.add_field(&parent_indices, props),
            "list" => return self.add_list(&parent_indices, props, pos),
            kind @ ("header" | "footer") => return self.add_header_footer(kind, props),
            _ => {}
        }
        if props.has("list") {
            self.ensure_numbering()?;
        }

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
            "toc" => {
                if parent_local != "body" {
                    bail!("a TOC is added to /body");
                }
                self.ensure_update_fields()?;
                self.build_toc_paragraph(props)
            }
            "hyperlink" | "link" => {
                if parent_local != "p" {
                    bail!("hyperlinks are added to a paragraph, e.g. add doc.docx '/body/p[1]' --type hyperlink --prop url=https://...");
                }
                let url = props
                    .get("url")
                    .or_else(|| props.get("href"))
                    .context("hyperlink needs --prop url=https://...")?
                    .to_string();
                let rid = crate::media::add_external_relationship(
                    &mut self.rels,
                    crate::media::HYPERLINK_REL_TYPE,
                    &url,
                );
                let text = props.get("text").unwrap_or(&url).to_string();
                let mut link = el("w:hyperlink", &[("r:id", rid.as_str()), ("w:history", "1")]);
                let mut run = build_run(&text, props)?;
                let rpr = run.ensure_child("rPr", "w:rPr", true);
                if rpr.child("color").is_none() {
                    rpr.push(el("w:color", &[("w:val", "0563C1")]));
                }
                if rpr.child("u").is_none() {
                    rpr.push(el("w:u", &[("w:val", "single")]));
                }
                link.push(run);
                link
            }
            other => bail!(
                "unsupported docx element type '{other}' (paragraph/run/table/row/break/image/toc/field/comment/footnote)"
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
            let changed = self.apply_page_setup(props)?;
            if changed.is_empty() {
                bail!(
                    "document-level set supports page setup props: page-size=letter|a4|a3|legal|WxH, orientation=portrait|landscape, margins=1in (or margin-top/right/bottom/left)"
                );
            }
            return Ok(Report::Data {
                text: format!("set {}", changed.join(", ")),
                data: json!({ "set": changed }),
            });
        }
        let indices = self.resolve(&dpath)?;
        let local = self.node_at(&indices)?.local_name().to_string();
        match local.as_str() {
            "p" => {
                if props.has("list") {
                    self.ensure_numbering()?;
                }
                self.apply_paragraph_props(&indices, props)?
            }
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
        // /comment[N] removes comment N plus its body markers.
        if dpath.segments.len() == 1 && dpath.segments[0].name.eq_ignore_ascii_case("comment") {
            let id = dpath.segments[0]
                .index()
                .context("comment removal needs an id, e.g. /comment[1]")? as u32;
            return self.remove_comment(id);
        }
        if dpath.segments.len() == 1 {
            let name = dpath.segments[0].name.to_ascii_lowercase();
            if name == "header" || name == "footer" {
                return self.remove_header_footer(&name);
            }
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

    fn copy_el(&mut self, path_str: &str, pos: &Position) -> Result<Report> {
        let indices = self.resolve(&path::parse(path_str)?)?;
        if indices.is_empty() {
            bail!("cannot copy the document body");
        }
        let element = self.node_at(&indices)?.clone();
        if element.local_name() == "sectPr" {
            bail!("section properties cannot be copied");
        }
        let (parent_indices, last) = indices.split_at(indices.len() - 1);
        // Default lands the copy right after the original.
        let pos = match pos {
            Position::Append => {
                let parent = self.node_at(parent_indices)?;
                let ordinal = parent.children[..last[0]]
                    .iter()
                    .filter(|n| matches!(n, XmlNode::Element(_)))
                    .count();
                Position::Index(ordinal + 1)
            }
            p => p.clone(),
        };
        let new_indices = self.insert_into(parent_indices, element, &pos)?;
        let display = self.display_path(&new_indices)?;
        let node = self.node_at(&new_indices)?;
        let mut info = self.node_info(node, &display, 0);
        info.attr("copied-from", path_str);
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

    fn screenshot(&mut self) -> Result<Vec<Vec<u8>>> {
        use crate::render::{Canvas, Color, Span, BLACK, GRID, WHITE};
        // US Letter at 96 dpi with 1in margins.
        let (page_w, page_h, margin) = (816.0f32, 1056.0f32, 96.0f32);
        let content_w = page_w - 2.0 * margin;
        let px_per_pt = 96.0 / 72.0;

        let body = self.body()?.clone();
        let mut pages: Vec<Canvas> = vec![Canvas::new(page_w as u32, page_h as u32, WHITE)?];
        let mut y = margin;
        let mut list_counters: std::collections::HashMap<(String, u32), u32> = Default::default();
        macro_rules! new_page {
            () => {{
                pages.push(Canvas::new(page_w as u32, page_h as u32, WHITE)?);
                y = margin;
            }};
        }
        macro_rules! need {
            ($h:expr) => {
                if y + $h > page_h - margin && y > margin {
                    new_page!();
                }
            };
        }

        for (_, element) in logical_children(&body) {
            match element.local_name() {
                "p" => {
                    // Embedded image?
                    if let Some(props) = self.dump_image_paragraph(element)? {
                        use base64::Engine;
                        let bytes = props
                            .get("srcdata")
                            .and_then(|v| v.as_str())
                            .and_then(|b64| {
                                base64::engine::general_purpose::STANDARD.decode(b64).ok()
                            })
                            .unwrap_or_default();
                        let emu_px = |v: Option<&serde_json::Value>| {
                            v.and_then(|v| v.as_str())
                                .and_then(|s| s.parse::<f64>().ok())
                                .map(|emu| (emu / 9525.0) as f32)
                        };
                        let iw = emu_px(props.get("w")).unwrap_or(200.0).min(content_w);
                        let ih = emu_px(props.get("h")).unwrap_or(150.0).min(page_h - 2.0 * margin);
                        need!(ih);
                        let align = element
                            .child("pPr")
                            .and_then(|ppr| ppr.child("jc"))
                            .and_then(|s| s.attr_local("val"));
                        let x = match align {
                            Some("center") => margin + (content_w - iw) / 2.0,
                            Some("right") => margin + content_w - iw,
                            _ => margin,
                        };
                        pages.last_mut().unwrap().draw_image(&bytes, x, y, iw, ih);
                        y += ih + 8.0;
                        continue;
                    }
                    let style = element
                        .child("pPr")
                        .and_then(|ppr| ppr.child("pStyle"))
                        .and_then(|s| s.attr_local("val"))
                        .unwrap_or("Normal");
                    let (base_pt, base_bold) = match style {
                        "Title" => (28.0, true),
                        "Heading1" => (20.0, true),
                        "Heading2" => (16.0, true),
                        "Heading3" => (14.0, true),
                        _ => (11.0, false),
                    };
                    let mut spans: Vec<Span> = Vec::new();
                    let mut forced_break = false;
                    for run in element.elements().filter(|e| e.local_name() == "r") {
                        if run
                            .elements()
                            .any(|c| c.local_name() == "br" && c.attr_local("type") == Some("page"))
                        {
                            forced_break = true;
                        }
                        let rpr = run.child("rPr");
                        let size_pt = rpr
                            .and_then(|rp| rp.child("sz"))
                            .and_then(|s| s.attr_local("val"))
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|half| half / 2.0)
                            .unwrap_or(base_pt);
                        let bold = base_bold
                            || rpr.map(|rp| rp.child("b").is_some()).unwrap_or(false);
                        let color = rpr
                            .and_then(|rp| rp.child("color"))
                            .and_then(|c| c.attr_local("val"))
                            .and_then(Color::from_hex)
                            .unwrap_or(BLACK);
                        let text = run_text(run);
                        if !text.is_empty() {
                            spans.push(Span {
                                text,
                                size: size_pt * px_per_pt,
                                color,
                                bold,
                            });
                        }
                    }
                    // List paragraphs get a marker span and an indent.
                    let mut left_indent = 0.0f32;
                    if let Some(marker) = self.list_marker(element, &mut list_counters) {
                        let level = paragraph_numpr(element).map(|(_, l)| l).unwrap_or(0);
                        left_indent = 24.0 * (level + 1) as f32;
                        spans.insert(
                            0,
                            Span {
                                text: marker.trim_start().to_string(),
                                size: base_pt * px_per_pt,
                                color: BLACK,
                                bold: false,
                            },
                        );
                    }
                    let align = element
                        .child("pPr")
                        .and_then(|ppr| ppr.child("jc"))
                        .and_then(|s| s.attr_local("val"));
                    let max_size = spans
                        .iter()
                        .map(|s| s.size)
                        .fold(base_pt * px_per_pt, f32::max);
                    let line_h = max_size * 1.35;
                    if spans.is_empty() {
                        y += line_h * 0.6;
                    } else {
                        let canvas_probe = pages.last().unwrap();
                        let usable = content_w - left_indent;
                        let lines = canvas_probe.layout_spans(&spans, usable);
                        for line in lines {
                            need!(line_h);
                            let canvas = pages.last_mut().unwrap();
                            let lw = canvas.spans_width(&line);
                            let lx = match align {
                                Some("center") => margin + (usable - lw) / 2.0,
                                Some("right") => margin + usable - lw,
                                _ => margin,
                            } + left_indent;
                            let asc = canvas.ascent(max_size, base_bold);
                            canvas.draw_spans_line(&line, lx, y + asc);
                            y += line_h;
                        }
                        y += line_h * 0.3;
                    }
                    if forced_break {
                        new_page!();
                    }
                }
                "tbl" => {
                    let rows = element.children_named("tr");
                    let cols = rows
                        .first()
                        .map(|tr| tr.children_named("tc").len())
                        .unwrap_or(0)
                        .max(1);
                    let col_w = content_w / cols as f32;
                    let row_h = 30.0f32;
                    let font = 11.0 * px_per_pt;
                    for tr in rows {
                        need!(row_h);
                        let canvas = pages.last_mut().unwrap();
                        for (ci, tc) in tr.children_named("tc").iter().enumerate() {
                            let cx = margin + col_w * ci as f32;
                            canvas.stroke_rect(cx, y, col_w, row_h, GRID);
                            let text = tc.text_content();
                            canvas.draw_text(&text, cx + 6.0, y + row_h - 10.0, font, BLACK, false);
                        }
                        y += row_h;
                    }
                    y += 10.0;
                }
                _ => {}
            }
        }
        pages.iter().map(|c| c.png()).collect()
    }

    fn dump(&mut self) -> Result<serde_json::Value> {
        let body = self.body()?.clone();
        let mut ops = Vec::new();
        let mut tbl_count = 0usize;
        let mut p_count = 0usize;
        for (_, element) in logical_children(&body) {
            match element.local_name() {
                "p" => {
                    p_count += 1;
                    // Image paragraphs replay as `add image` with the bytes
                    // embedded (srcdata).
                    if let Some(mut props) = self.dump_image_paragraph(element)? {
                        if let Some(jc) = element
                            .child("pPr")
                            .and_then(|ppr| ppr.child("jc"))
                            .and_then(|s| s.attr_local("val"))
                        {
                            props.insert("align".into(), json!(jc));
                        }
                        ops.push(json!({
                            "command": "add", "parent": "/body",
                            "type": "image", "props": props,
                        }));
                        continue;
                    }
                    let mut props = serde_json::Map::new();
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
                    if let Some((num_id, level)) = paragraph_numpr(element) {
                        let kind = self.list_kind_for_num(&num_id).unwrap_or("bullet");
                        props.insert("list".into(), json!(kind));
                        if level > 0 {
                            props.insert("level".into(), json!(level.to_string()));
                        }
                    }
                    let runs: Vec<&XmlElement> =
                        element.elements().filter(|e| e.local_name() == "r").collect();
                    let uniform = runs.len() <= 1
                        || runs.windows(2).all(|w| w[0].child("rPr") == w[1].child("rPr"));
                    if uniform {
                        props.insert("text".into(), json!(paragraph_text(element)));
                        if let Some(rpr) = runs.first().and_then(|r| r.child("rPr")) {
                            for (k, v) in rpr_props(rpr) {
                                props.insert(k, json!(v));
                            }
                        }
                        ops.push(json!({
                            "command": "add", "parent": "/body",
                            "type": "paragraph", "props": props,
                        }));
                    } else {
                        // Mixed formatting: empty paragraph + one run op each.
                        ops.push(json!({
                            "command": "add", "parent": "/body",
                            "type": "paragraph", "props": props,
                        }));
                        for run in runs {
                            let mut rprops = serde_json::Map::new();
                            rprops.insert("text".into(), json!(run_text(run)));
                            if let Some(rpr) = run.child("rPr") {
                                for (k, v) in rpr_props(rpr) {
                                    rprops.insert(k, json!(v));
                                }
                            }
                            ops.push(json!({
                                "command": "add",
                                "parent": format!("/body/p[{p_count}]"),
                                "type": "run", "props": rprops,
                            }));
                        }
                    }
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
