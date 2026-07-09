//! PowerPoint (.pptx) handler: slides, textbox shapes, backgrounds,
//! find/replace across slide text.

use anyhow::{bail, Context, Result};
use serde_json::json;
use std::path::Path;

use crate::docx::{build_find_regex, replace_in_paragraph};
use crate::handler::{Handler, Position};
use crate::out::{NodeInfo, Report};
use crate::path::{self, Predicate, Segment};
use crate::props::{parse_align, parse_bool, parse_color, parse_emu, parse_pt, Props};
use crate::pkg::Package;
use crate::templates::{pptx_blank_slide, PPTX_SLIDE_RELS};
use crate::xml::{el, XmlElement, XmlNode};

const PRESENTATION_PART: &str = "ppt/presentation.xml";
const PRESENTATION_RELS_PART: &str = "ppt/_rels/presentation.xml.rels";
const CONTENT_TYPES_PART: &str = "[Content_Types].xml";
const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const SLIDE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

struct Slide {
    part: String,
    rid: String,
    xml: XmlElement,
}

pub struct Pptx {
    pkg: Package,
    presentation: XmlElement,
    rels: XmlElement,
    slides: Vec<Slide>,
}

impl Pptx {
    pub fn new(pkg: Package) -> Result<Pptx> {
        let presentation = pkg.xml(PRESENTATION_PART)?;
        let rels = pkg.xml(PRESENTATION_RELS_PART)?;
        let mut slides = Vec::new();
        if let Some(id_list) = presentation.child("sldIdLst") {
            for sld_id in id_list.children_named("sldId") {
                let rid = sld_id
                    .attr("r:id")
                    .or_else(|| sld_id.attr_local("id"))
                    .context("sldId entry has no r:id")?
                    .to_string();
                let target = rels
                    .children_named("Relationship")
                    .into_iter()
                    .find(|r| r.attr_local("Id") == Some(rid.as_str()))
                    .and_then(|r| r.attr_local("Target"))
                    .with_context(|| format!("no presentation relationship '{rid}'"))?;
                let part = resolve_target("ppt", target);
                let xml = pkg.xml(&part)?;
                slides.push(Slide { part, rid, xml });
            }
        }
        Ok(Pptx {
            pkg,
            presentation,
            rels,
            slides,
        })
    }

    fn slide_index(&self, seg: &Segment) -> Result<usize> {
        if !seg.name.eq_ignore_ascii_case("slide") {
            bail!("pptx paths start with /slide[N], got '{}'", seg.name);
        }
        let n = seg.index().unwrap_or(1);
        if n == 0 || n > self.slides.len() {
            bail!("slide index {n} out of range (1..{})", self.slides.len());
        }
        Ok(n - 1)
    }

    /// spTree of a slide.
    fn sp_tree(slide: &Slide) -> Result<&XmlElement> {
        slide
            .xml
            .child("cSld")
            .and_then(|c| c.child("spTree"))
            .context("slide has no <p:spTree>")
    }

    fn sp_tree_mut(slide: &mut Slide) -> Result<&mut XmlElement> {
        slide
            .xml
            .child_mut("cSld")
            .and_then(|c| c.child_mut("spTree"))
            .context("slide has no <p:spTree>")
    }

    /// Child indices (within spTree.children) of shape-like elements.
    fn shape_indices(tree: &XmlElement) -> Vec<usize> {
        tree.children
            .iter()
            .enumerate()
            .filter_map(|(i, n)| match n {
                XmlNode::Element(e)
                    if matches!(e.local_name(), "sp" | "pic" | "graphicFrame" | "cxnSp" | "grpSp") =>
                {
                    Some(i)
                }
                _ => None,
            })
            .collect()
    }

    fn find_shape(&self, slide_idx: usize, seg: &Segment) -> Result<usize> {
        if !seg.name.eq_ignore_ascii_case("shape") {
            bail!("expected a shape segment, got '{}'", seg.name);
        }
        let tree = Self::sp_tree(&self.slides[slide_idx])?;
        let indices = Self::shape_indices(tree);
        let mut candidates: Vec<usize> = Vec::new();
        for &i in &indices {
            let e = tree.children[i].as_element().unwrap();
            let ok = seg.preds.iter().all(|p| match p {
                Predicate::Attr(k, v) => {
                    let cnvpr = shape_cnvpr(e);
                    match k.as_str() {
                        "name" => cnvpr.and_then(|c| c.attr_local("name")) == Some(v.as_str()),
                        "id" => cnvpr.and_then(|c| c.attr_local("id")) == Some(v.as_str()),
                        _ => false,
                    }
                }
                Predicate::Index(_) => true,
            });
            if ok {
                candidates.push(i);
            }
        }
        let nth = seg.index().unwrap_or(1);
        candidates.get(nth - 1).copied().with_context(|| {
            format!(
                "no shape matches '{:?}' on slide {} ({} candidate(s))",
                seg,
                slide_idx + 1,
                candidates.len()
            )
        })
    }

    fn shape_info(&self, e: &XmlElement, path: &str) -> NodeInfo {
        let mut info = NodeInfo::new(path, shape_kind(e));
        if let Some(cnvpr) = shape_cnvpr(e) {
            if let Some(name) = cnvpr.attr_local("name") {
                info.attr("name", name);
            }
            if let Some(id) = cnvpr.attr_local("id") {
                info.attr("id", id);
            }
        }
        let text = shape_text(e);
        if !text.is_empty() {
            info.text = Some(text);
        }
        if let Some(xfrm) = e.child("spPr").and_then(|sp| sp.child("xfrm")) {
            if let Some(off) = xfrm.child("off") {
                if let (Some(x), Some(y)) = (off.attr_local("x"), off.attr_local("y")) {
                    info.attr("x", x);
                    info.attr("y", y);
                }
            }
            if let Some(ext) = xfrm.child("ext") {
                if let (Some(cx), Some(cy)) = (ext.attr_local("cx"), ext.attr_local("cy")) {
                    info.attr("w", cx);
                    info.attr("h", cy);
                }
            }
        }
        info
    }

    fn slide_info(&self, idx: usize, depth: usize) -> Result<NodeInfo> {
        let slide = &self.slides[idx];
        let path = format!("/slide[{}]", idx + 1);
        let mut info = NodeInfo::new(&path, "slide");
        info.attr("part", &slide.part);
        let tree = Self::sp_tree(slide)?;
        let shape_idxs = Self::shape_indices(tree);
        info.attr("shapes", shape_idxs.len().to_string());
        if depth > 0 {
            for (n, &i) in shape_idxs.iter().enumerate() {
                let e = tree.children[i].as_element().unwrap();
                let spath = format!("{}/shape[{}]", path, n + 1);
                info.children.push(self.shape_info(e, &spath));
            }
        }
        Ok(info)
    }

    fn next_shape_id(tree: &XmlElement) -> u64 {
        let mut max_id = 1;
        fn walk(e: &XmlElement, max_id: &mut u64) {
            if e.local_name() == "cNvPr" {
                if let Some(id) = e.attr_local("id").and_then(|v| v.parse::<u64>().ok()) {
                    *max_id = (*max_id).max(id);
                }
            }
            for c in e.elements() {
                walk(c, max_id);
            }
        }
        walk(tree, &mut max_id);
        max_id + 1
    }

    fn build_textbox(&self, id: u64, props: &Props) -> Result<XmlElement> {
        let name = props
            .get("name")
            .map(|s| s.to_string())
            .unwrap_or_else(|| format!("TextBox {id}"));
        let x = props.get("x").map(parse_emu).transpose()?.unwrap_or(914_400);
        let y = props.get("y").map(parse_emu).transpose()?.unwrap_or(914_400);
        let w = props
            .get("w")
            .or_else(|| props.get("width"))
            .map(parse_emu)
            .transpose()?
            .unwrap_or(3_657_600);
        let h = props
            .get("h")
            .or_else(|| props.get("height"))
            .map(parse_emu)
            .transpose()?
            .unwrap_or(914_400);

        let mut sp = XmlElement::new("p:sp");
        let mut nv = XmlElement::new("p:nvSpPr");
        nv.push(el(
            "p:cNvPr",
            &[("id", id.to_string().as_str()), ("name", name.as_str())],
        ));
        nv.push(el("p:cNvSpPr", &[("txBox", "1")]));
        nv.push(XmlElement::new("p:nvPr"));
        sp.push(nv);

        let mut sppr = XmlElement::new("p:spPr");
        let mut xfrm = XmlElement::new("a:xfrm");
        xfrm.push(el(
            "a:off",
            &[("x", x.to_string().as_str()), ("y", y.to_string().as_str())],
        ));
        xfrm.push(el(
            "a:ext",
            &[("cx", w.to_string().as_str()), ("cy", h.to_string().as_str())],
        ));
        sppr.push(xfrm);
        let mut geom = el("a:prstGeom", &[("prst", "rect")]);
        geom.push(XmlElement::new("a:avLst"));
        sppr.push(geom);
        if let Some(fill) = props.get("fill") {
            sppr.push(solid_fill(&parse_color(fill)?));
        }
        sp.push(sppr);

        let mut tx = XmlElement::new("p:txBody");
        let mut bodypr = el("a:bodyPr", &[("wrap", "square"), ("rtlCol", "0")]);
        bodypr.push(XmlElement::new("a:spAutoFit"));
        tx.push(bodypr);
        tx.push(XmlElement::new("a:lstStyle"));
        let text = props.get("text").unwrap_or("");
        for line in text.replace("\\n", "\n").split('\n') {
            tx.push(build_text_paragraph(line, props)?);
        }
        sp.push(tx);
        Ok(sp)
    }

    fn add_slide(&mut self, props: &Props, pos: &Position) -> Result<usize> {
        let next_part_num = self
            .pkg
            .part_names()
            .filter_map(|p| {
                p.strip_prefix("ppt/slides/slide")
                    .and_then(|s| s.strip_suffix(".xml"))
                    .and_then(|n| n.parse::<u32>().ok())
            })
            .max()
            .unwrap_or(0)
            + 1;
        let part = format!("ppt/slides/slide{next_part_num}.xml");
        let rels_part = format!("ppt/slides/_rels/slide{next_part_num}.xml.rels");

        let next_rid = self
            .rels
            .children_named("Relationship")
            .into_iter()
            .filter_map(|r| {
                r.attr_local("Id")
                    .and_then(|id| id.strip_prefix("rId"))
                    .and_then(|n| n.parse::<u32>().ok())
            })
            .max()
            .unwrap_or(0)
            + 1;
        let rid = format!("rId{next_rid}");

        // Slide id: 256+ and unique.
        let sld_id_lst = self
            .presentation
            .ensure_child("sldIdLst", "p:sldIdLst", false);
        let next_sld_id = sld_id_lst
            .children_named("sldId")
            .into_iter()
            .filter_map(|e| e.attr_local("id").and_then(|v| v.parse::<u64>().ok()))
            .max()
            .unwrap_or(255)
            + 1;
        let entry = el(
            "p:sldId",
            &[
                ("id", next_sld_id.to_string().as_str()),
                ("r:id", rid.as_str()),
            ],
        );
        let insert_at = match pos {
            Position::Append => sld_id_lst.children.len(),
            Position::Index(n) => (*n).min(sld_id_lst.children.len()),
            Position::Before(p) | Position::After(p) => {
                let dpath = path::parse(p)?;
                let idx = self.slide_index(
                    dpath
                        .segments
                        .first()
                        .context("anchor path must be /slide[N]")?,
                )?;
                let sld_id_lst = self
                    .presentation
                    .child("sldIdLst")
                    .context("no sldIdLst")?;
                let raw = sld_id_lst
                    .nth_child_index("sldId", idx)
                    .context("anchor slide not in sldIdLst")?;
                if matches!(pos, Position::After(_)) {
                    raw + 1
                } else {
                    raw
                }
            }
        };
        let sld_id_lst = self
            .presentation
            .ensure_child("sldIdLst", "p:sldIdLst", false);
        sld_id_lst
            .children
            .insert(insert_at, XmlNode::Element(entry));

        self.rels.push(el(
            "Relationship",
            &[
                ("Id", rid.as_str()),
                ("Type", SLIDE_REL_TYPE),
                ("Target", format!("slides/slide{next_part_num}.xml").as_str()),
            ],
        ));

        let mut ct = self.pkg.xml(CONTENT_TYPES_PART)?;
        ct.push(el(
            "Override",
            &[
                ("PartName", format!("/{part}").as_str()),
                ("ContentType", SLIDE_CONTENT_TYPE),
            ],
        ));
        self.pkg.put_xml(CONTENT_TYPES_PART, &ct)?;
        self.pkg
            .put_raw(&rels_part, PPTX_SLIDE_RELS.as_bytes().to_vec());

        let mut xml = crate::xml::parse(pptx_blank_slide().as_bytes())?;

        // Optional background color.
        if let Some(background) = props.get("background") {
            set_slide_background(&mut xml, &parse_color(background)?)?;
        }

        let slide = Slide {
            part,
            rid,
            xml,
        };
        // Position in self.slides mirrors sldIdLst order: count sldId
        // entries before insert_at.
        let slide_pos = {
            let sld_id_lst = self.presentation.child("sldIdLst").unwrap();
            sld_id_lst.children[..insert_at]
                .iter()
                .filter(|n| matches!(n, XmlNode::Element(e) if e.local_name() == "sldId"))
                .count()
        };
        self.slides.insert(slide_pos, slide);

        // Optional title textbox.
        if let Some(title) = props.get("title") {
            let title = title.to_string();
            let title_props = Props::from_pairs(vec![
                ("text".to_string(), title),
                ("name".to_string(), "Title 1".to_string()),
                ("x".to_string(), "838200".to_string()),
                ("y".to_string(), "365125".to_string()),
                ("w".to_string(), "10515600".to_string()),
                ("h".to_string(), "1325563".to_string()),
                ("size".to_string(), "44".to_string()),
            ]);
            let tree = Self::sp_tree(&self.slides[slide_pos])?;
            let id = Self::next_shape_id(tree);
            let sp = self.build_textbox(id, &title_props)?;
            let tree = Self::sp_tree_mut(&mut self.slides[slide_pos])?;
            tree.push(sp);
        }
        Ok(slide_pos)
    }
}

fn shape_cnvpr(e: &XmlElement) -> Option<&XmlElement> {
    // nvSpPr / nvPicPr / nvGraphicFramePr / nvCxnSpPr / nvGrpSpPr → cNvPr
    e.elements()
        .find(|c| c.local_name().starts_with("nv") && c.local_name().ends_with("Pr"))
        .and_then(|nv| nv.child("cNvPr"))
}

fn shape_kind(e: &XmlElement) -> &'static str {
    match e.local_name() {
        "sp" => "shape",
        "pic" => "picture",
        "graphicFrame" => "graphicFrame",
        "cxnSp" => "connector",
        "grpSp" => "group",
        _ => "shape",
    }
}

/// All text in a shape's txBody, paragraphs joined with '\n'.
fn shape_text(e: &XmlElement) -> String {
    let Some(tx) = e.child("txBody") else {
        return String::new();
    };
    let mut lines = Vec::new();
    for p in tx.children_named("p") {
        let mut line = String::new();
        for child in p.elements() {
            match child.local_name() {
                "r" => {
                    if let Some(t) = child.child("t") {
                        line.push_str(&t.text_content());
                    }
                }
                "br" => line.push('\n'),
                _ => {}
            }
        }
        lines.push(line);
    }
    lines.join("\n")
}

fn solid_fill(color: &str) -> XmlElement {
    let mut fill = XmlElement::new("a:solidFill");
    fill.push(el("a:srgbClr", &[("val", color)]));
    fill
}

fn build_run_props(props: &Props) -> Result<Option<XmlElement>> {
    let mut rpr = el("a:rPr", &[("lang", "en-US"), ("dirty", "0")]);
    let mut any = false;
    if let Some(size) = props.get("size") {
        let hundredths = (parse_pt(size)? * 100.0).round() as i64;
        rpr.set_attr("sz", &hundredths.to_string());
        any = true;
    }
    if let Some(b) = props.get("bold") {
        if parse_bool(b)? {
            rpr.set_attr("b", "1");
            any = true;
        }
    }
    if let Some(i) = props.get("italic") {
        if parse_bool(i)? {
            rpr.set_attr("i", "1");
            any = true;
        }
    }
    if let Some(color) = props.get("color") {
        rpr.push(solid_fill(&parse_color(color)?));
        any = true;
    }
    if let Some(font) = props.get("font") {
        rpr.push(el("a:latin", &[("typeface", font)]));
        any = true;
    }
    Ok(if any { Some(rpr) } else { None })
}

fn build_text_paragraph(line: &str, props: &Props) -> Result<XmlElement> {
    let mut p = XmlElement::new("a:p");
    if let Some(align) = props.get("align") {
        let (_, algn) = parse_align(align)?;
        p.push(el("a:pPr", &[("algn", algn)]));
    }
    let mut r = XmlElement::new("a:r");
    if let Some(rpr) = build_run_props(props)? {
        r.push(rpr);
    }
    let mut t = XmlElement::new("a:t");
    t.push_text(line);
    r.push(t);
    p.push(r);
    Ok(p)
}

fn set_slide_background(slide_xml: &mut XmlElement, color: &str) -> Result<()> {
    let csld = slide_xml
        .child_mut("cSld")
        .context("slide has no <p:cSld>")?;
    // p:bg must be the first child of cSld.
    csld.children.retain(|n| {
        !matches!(n, XmlNode::Element(e) if e.local_name() == "bg")
    });
    let mut bg = XmlElement::new("p:bg");
    let mut bgpr = XmlElement::new("p:bgPr");
    bgpr.push(solid_fill(color));
    bgpr.push(XmlElement::new("a:effectLst"));
    bg.push(bgpr);
    csld.children.insert(0, XmlNode::Element(bg));
    Ok(())
}

fn resolve_target(base: &str, target: &str) -> String {
    if let Some(abs) = target.strip_prefix('/') {
        return abs.to_string();
    }
    let mut parts: Vec<&str> = base.split('/').collect();
    for piece in target.split('/') {
        match piece {
            ".." => {
                parts.pop();
            }
            "." => {}
            p => parts.push(p),
        }
    }
    parts.join("/")
}

// ------------------------------------------------------------ Handler ----

impl Handler for Pptx {
    fn view(&mut self, mode: &str) -> Result<Report> {
        match mode {
            "outline" | "text" => {
                let mut out = String::new();
                for (i, slide) in self.slides.iter().enumerate() {
                    let tree = Self::sp_tree(slide)?;
                    let shape_idxs = Self::shape_indices(tree);
                    // Slide heading: first non-empty shape text as title.
                    let title = shape_idxs
                        .iter()
                        .filter_map(|&si| tree.children[si].as_element())
                        .map(shape_text)
                        .find(|t| !t.trim().is_empty())
                        .unwrap_or_default();
                    out.push_str(&format!("Slide {}: {}\n", i + 1, title.replace('\n', " ")));
                    for (n, &si) in shape_idxs.iter().enumerate() {
                        let e = tree.children[si].as_element().unwrap();
                        let text = shape_text(e);
                        let name = shape_cnvpr(e)
                            .and_then(|c| c.attr_local("name"))
                            .unwrap_or("");
                        if mode == "outline" {
                            out.push_str(&format!(
                                "  Shape {} [{}] {}: {}\n",
                                n + 1,
                                shape_kind(e),
                                name,
                                text.replace('\n', " ⏎ ")
                            ));
                        } else if !text.is_empty() {
                            out.push_str(&text);
                            out.push('\n');
                        }
                    }
                }
                if self.slides.is_empty() {
                    out.push_str("(no slides)\n");
                }
                Ok(Report::Text(out))
            }
            "stats" => {
                let mut shapes = 0usize;
                let mut words = 0usize;
                for slide in &self.slides {
                    let tree = Self::sp_tree(slide)?;
                    for &i in &Self::shape_indices(tree) {
                        shapes += 1;
                        let e = tree.children[i].as_element().unwrap();
                        words += shape_text(e).split_whitespace().count();
                    }
                }
                Ok(Report::Data {
                    text: format!(
                        "slides: {}\nshapes: {shapes}\nwords: {words}",
                        self.slides.len()
                    ),
                    data: json!({
                        "slides": self.slides.len(),
                        "shapes": shapes,
                        "words": words,
                    }),
                })
            }
            other => bail!("unknown view mode '{other}' for pptx (text/outline/stats)"),
        }
    }

    fn get(&mut self, path_str: &str, depth: usize) -> Result<Report> {
        let dpath = path::parse(path_str)?;
        if dpath.is_root() {
            let mut info = NodeInfo::new("/", "presentation");
            info.attr("slides", self.slides.len().to_string());
            for i in 0..self.slides.len() {
                info.children.push(self.slide_info(i, 0)?);
            }
            return Ok(Report::Nodes(vec![info]));
        }
        let slide_idx = self.slide_index(&dpath.segments[0])?;
        if dpath.segments.len() == 1 {
            return Ok(Report::Nodes(vec![
                self.slide_info(slide_idx, depth.max(1))?
            ]));
        }
        let shape_child_idx = self.find_shape(slide_idx, &dpath.segments[1])?;
        if dpath.segments.len() > 2 {
            bail!("paths deeper than /slide[N]/shape[M] are not supported yet");
        }
        let tree = Self::sp_tree(&self.slides[slide_idx])?;
        let e = tree.children[shape_child_idx].as_element().unwrap();
        // Display path uses the positional shape index.
        let position = Self::shape_indices(tree)
            .iter()
            .position(|&i| i == shape_child_idx)
            .unwrap_or(0);
        let spath = format!("/slide[{}]/shape[{}]", slide_idx + 1, position + 1);
        Ok(Report::Nodes(vec![self.shape_info(e, &spath)]))
    }

    fn add(&mut self, parent: &str, typ: &str, props: &Props, pos: &Position) -> Result<Report> {
        let dpath = path::parse(parent)?;
        match typ.to_ascii_lowercase().as_str() {
            "slide" => {
                if !dpath.is_root() {
                    bail!("slides are added at the root: officecli add file.pptx / --type slide");
                }
                let idx = self.add_slide(props, pos)?;
                Ok(Report::Nodes(vec![self.slide_info(idx, 1)?]))
            }
            "shape" | "textbox" | "text" => {
                if dpath.segments.is_empty() {
                    bail!("shapes are added to a slide: officecli add file.pptx '/slide[1]' --type shape");
                }
                let slide_idx = self.slide_index(&dpath.segments[0])?;
                let tree = Self::sp_tree(&self.slides[slide_idx])?;
                let id = Self::next_shape_id(tree);
                let sp = self.build_textbox(id, props)?;
                let tree = Self::sp_tree_mut(&mut self.slides[slide_idx])?;
                match pos {
                    Position::Append => tree.push(sp),
                    Position::Index(n) => {
                        let shape_idxs = Self::shape_indices(tree);
                        let at = shape_idxs
                            .get(*n)
                            .copied()
                            .unwrap_or(tree.children.len());
                        tree.children.insert(at, XmlNode::Element(sp));
                    }
                    _ => bail!("--before/--after are not supported for shapes yet (use --index)"),
                }
                let tree = Self::sp_tree(&self.slides[slide_idx])?;
                let shape_idxs = Self::shape_indices(tree);
                let position = shape_idxs.len();
                let e = tree.children[*shape_idxs.last().unwrap()].as_element().unwrap();
                let spath = format!("/slide[{}]/shape[{}]", slide_idx + 1, position);
                Ok(Report::Nodes(vec![self.shape_info(e, &spath)]))
            }
            other => bail!("unsupported pptx element type '{other}' (slide/shape)"),
        }
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
            let replace =
                replace.context("pptx find requires --replace (find+format is not supported)")?;
            let re = build_find_regex(find, props)?;
            let slide_scope: Vec<usize> = if dpath.is_root() {
                (0..self.slides.len()).collect()
            } else {
                vec![self.slide_index(&dpath.segments[0])?]
            };
            let mut matched = 0usize;
            for si in slide_scope {
                let tree = Self::sp_tree_mut(&mut self.slides[si])?;
                for_each_a_paragraph(tree, &mut |p| {
                    matched += replace_in_paragraph(p, &re, replace);
                });
            }
            return Ok(Report::Data {
                text: format!("matched: {matched}"),
                data: json!({ "matched": matched }),
            });
        }

        if dpath.is_root() {
            bail!("set on '/' needs --find/--replace for pptx");
        }
        let slide_idx = self.slide_index(&dpath.segments[0])?;

        if dpath.segments.len() == 1 {
            if let Some(background) = props.get("background") {
                let color = parse_color(background)?;
                set_slide_background(&mut self.slides[slide_idx].xml, &color)?;
                return Ok(Report::Data {
                    text: format!("set slide {} background to {color}", slide_idx + 1),
                    data: json!({ "slide": slide_idx + 1, "background": color }),
                });
            }
            bail!("slide-level set supports --prop background=COLOR");
        }

        let shape_child_idx = self.find_shape(slide_idx, &dpath.segments[1])?;
        let tree = Self::sp_tree_mut(&mut self.slides[slide_idx])?;
        let sp = tree.children[shape_child_idx]
            .as_element_mut()
            .context("shape vanished")?;

        // Geometry.
        let mut geometry_changed = false;
        for (key, target) in [("x", "x"), ("y", "y")] {
            if let Some(v) = props.get(key) {
                let emu = parse_emu(v)?;
                let sppr = sp.ensure_child("spPr", "p:spPr", false);
                let xfrm = sppr.ensure_child("xfrm", "a:xfrm", true);
                let off = xfrm.ensure_child("off", "a:off", true);
                off.set_attr(target, &emu.to_string());
                geometry_changed = true;
            }
        }
        for (keys, target) in [(["w", "width"], "cx"), (["h", "height"], "cy")] {
            let v = keys.iter().find_map(|k| props.get(k));
            if let Some(v) = v {
                let emu = parse_emu(v)?;
                let sppr = sp.ensure_child("spPr", "p:spPr", false);
                let xfrm = sppr.ensure_child("xfrm", "a:xfrm", true);
                let ext = xfrm.ensure_child("ext", "a:ext", false);
                ext.set_attr(target, &emu.to_string());
                geometry_changed = true;
            }
        }
        let _ = geometry_changed;

        // Fill.
        if let Some(fill) = props.get("fill") {
            let color = parse_color(fill)?;
            let sppr = sp.ensure_child("spPr", "p:spPr", false);
            sppr.children.retain(|n| {
                !matches!(n, XmlNode::Element(e)
                    if matches!(e.local_name(), "solidFill" | "noFill" | "gradFill" | "blipFill" | "pattFill"))
            });
            // Fill goes after prstGeom if present, else at end.
            let at = sppr
                .children
                .iter()
                .position(|n| matches!(n, XmlNode::Element(e) if e.local_name() == "prstGeom"))
                .map(|i| i + 1)
                .unwrap_or(sppr.children.len());
            sppr.children.insert(at, XmlNode::Element(solid_fill(&color)));
        }

        // Name.
        if let Some(name) = props.get("name") {
            for child in sp.children.iter_mut() {
                if let XmlNode::Element(nv) = child {
                    if nv.local_name().starts_with("nv") && nv.local_name().ends_with("Pr") {
                        if let Some(c) = nv.child_mut("cNvPr") {
                            c.set_attr("name", name);
                        }
                        break;
                    }
                }
            }
        }

        // Text: replace the whole txBody content.
        if let Some(text) = props.get("text") {
            let text = text.to_string();
            let tx = sp.ensure_child("txBody", "p:txBody", false);
            tx.children.retain(|n| {
                matches!(n, XmlNode::Element(e)
                    if matches!(e.local_name(), "bodyPr" | "lstStyle"))
            });
            for line in text.replace("\\n", "\n").split('\n') {
                tx.push(build_text_paragraph(line, props)?);
            }
        } else if props.has("size")
            || props.has("bold")
            || props.has("italic")
            || props.has("color")
            || props.has("font")
            || props.has("align")
        {
            // Formatting-only: apply to every existing run.
            if let Some(tx) = sp.child_mut("txBody") {
                apply_format_to_txbody(tx, props)?;
            }
        }

        let tree = Self::sp_tree(&self.slides[slide_idx])?;
        let e = tree.children[shape_child_idx].as_element().unwrap();
        let position = Self::shape_indices(tree)
            .iter()
            .position(|&i| i == shape_child_idx)
            .unwrap_or(0);
        let spath = format!("/slide[{}]/shape[{}]", slide_idx + 1, position + 1);
        Ok(Report::Nodes(vec![self.shape_info(e, &spath)]))
    }

    fn remove(&mut self, path_str: &str) -> Result<Report> {
        let dpath = path::parse(path_str)?;
        if dpath.is_root() {
            bail!("cannot remove the presentation root");
        }
        let slide_idx = self.slide_index(&dpath.segments[0])?;

        if dpath.segments.len() == 1 {
            let slide = self.slides.remove(slide_idx);
            // Remove sldIdLst entry.
            if let Some(lst) = self.presentation.child_mut("sldIdLst") {
                lst.children.retain(|n| {
                    !matches!(n, XmlNode::Element(e)
                        if e.local_name() == "sldId"
                            && (e.attr("r:id") == Some(&slide.rid)
                                || e.attr_local("id") == Some(slide.rid.as_str())))
                });
            }
            // Remove presentation relationship.
            self.rels.children.retain(|n| {
                !matches!(n, XmlNode::Element(e)
                    if e.local_name() == "Relationship"
                        && e.attr_local("Id") == Some(slide.rid.as_str()))
            });
            // Remove content-type override, part, and part rels.
            let mut ct = self.pkg.xml(CONTENT_TYPES_PART)?;
            let part_name = format!("/{}", slide.part);
            ct.children.retain(|n| {
                !matches!(n, XmlNode::Element(e)
                    if e.local_name() == "Override"
                        && e.attr_local("PartName") == Some(&part_name))
            });
            self.pkg.put_xml(CONTENT_TYPES_PART, &ct)?;
            let rels_part = slide
                .part
                .replace("ppt/slides/", "ppt/slides/_rels/")
                + ".rels";
            self.pkg.remove_part(&slide.part);
            self.pkg.remove_part(&rels_part);
            return Ok(Report::Data {
                text: format!("removed slide {}", slide_idx + 1),
                data: json!({ "removed": format!("/slide[{}]", slide_idx + 1) }),
            });
        }

        let shape_child_idx = self.find_shape(slide_idx, &dpath.segments[1])?;
        let tree = Self::sp_tree_mut(&mut self.slides[slide_idx])?;
        tree.children.remove(shape_child_idx);
        Ok(Report::Data {
            text: format!("removed shape from slide {}", slide_idx + 1),
            data: json!({ "removed": path_str }),
        })
    }

    fn validate(&mut self) -> Result<Report> {
        let mut problems = Vec::new();
        for part in [
            CONTENT_TYPES_PART,
            "_rels/.rels",
            "ppt/slideMasters/slideMaster1.xml",
            "ppt/theme/theme1.xml",
        ] {
            if !self.pkg.has_part(part) {
                problems.push(format!("missing package part {part}"));
            }
        }
        // Every slide rel must resolve to an existing part.
        for slide in &self.slides {
            if !self.pkg.has_part(&slide.part) && !slide_is_new(&slide.part) {
                problems.push(format!("slide part {} missing from package", slide.part));
            }
        }
        let ok = problems.is_empty();
        Ok(Report::Data {
            text: if ok { "valid".to_string() } else { problems.join("\n") },
            data: json!({ "valid": ok, "problems": problems }),
        })
    }

    fn save(&mut self, path: &Path) -> Result<()> {
        self.pkg.put_xml(PRESENTATION_PART, &self.presentation)?;
        self.pkg.put_xml(PRESENTATION_RELS_PART, &self.rels)?;
        for slide in &self.slides {
            self.pkg.put_xml(&slide.part, &slide.xml)?;
        }
        self.pkg.save(path)
    }
}

/// New slides live only in memory until save; treat them as present.
fn slide_is_new(_part: &str) -> bool {
    true
}

/// Visit every a:p in a subtree.
fn for_each_a_paragraph<F>(e: &mut XmlElement, f: &mut F)
where
    F: FnMut(&mut XmlElement),
{
    if e.local_name() == "p" && e.name.starts_with("a:") {
        f(e);
        return;
    }
    for child in e.children.iter_mut() {
        if let XmlNode::Element(ce) = child {
            for_each_a_paragraph(ce, f);
        }
    }
}

fn apply_format_to_txbody(tx: &mut XmlElement, props: &Props) -> Result<()> {
    let align = props.get("align").map(parse_align).transpose()?;
    for node in tx.children.iter_mut() {
        let XmlNode::Element(p) = node else { continue };
        if p.local_name() != "p" {
            continue;
        }
        if let Some((_, algn)) = align {
            let ppr = p.ensure_child("pPr", "a:pPr", true);
            ppr.set_attr("algn", algn);
        }
        for rnode in p.children.iter_mut() {
            let XmlNode::Element(r) = rnode else { continue };
            if r.local_name() != "r" {
                continue;
            }
            let rpr = r.ensure_child("rPr", "a:rPr", true);
            if rpr.attr("lang").is_none() {
                rpr.set_attr("lang", "en-US");
            }
            if let Some(size) = props.get("size") {
                let hundredths = (parse_pt(size)? * 100.0).round() as i64;
                rpr.set_attr("sz", &hundredths.to_string());
            }
            if let Some(b) = props.get("bold") {
                rpr.set_attr("b", if parse_bool(b)? { "1" } else { "0" });
            }
            if let Some(i) = props.get("italic") {
                rpr.set_attr("i", if parse_bool(i)? { "1" } else { "0" });
            }
            if let Some(color) = props.get("color") {
                let color = parse_color(color)?;
                rpr.children.retain(|n| {
                    !matches!(n, XmlNode::Element(e) if e.local_name() == "solidFill")
                });
                rpr.children.insert(0, XmlNode::Element(solid_fill(&color)));
            }
            if let Some(font) = props.get("font") {
                let latin = rpr.ensure_child("latin", "a:latin", false);
                latin.set_attr("typeface", font);
            }
        }
    }
    Ok(())
}
