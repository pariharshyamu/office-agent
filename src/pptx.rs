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
const NOTES_SLIDE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const NOTES_MASTER_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const NOTES_MASTER_PART: &str = "ppt/notesMasters/notesMaster1.xml";

struct Slide {
    part: String,
    rid: String,
    xml: XmlElement,
    /// The slide's own relationship part (layout, images, ...).
    rels: XmlElement,
}

fn slide_rels_part(slide_part: &str) -> String {
    slide_part.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
}

const EMPTY_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

pub struct Pptx {
    pkg: Package,
    presentation: XmlElement,
    rels: XmlElement,
    slides: Vec<Slide>,
    /// Theme palette: scheme color name (dk1, accent1, ...) → RRGGBB.
    theme: std::collections::HashMap<String, String>,
}

/// Load the clrScheme of the first theme part into name → RRGGBB.
fn load_theme_colors(pkg: &Package) -> std::collections::HashMap<String, String> {
    let mut map = std::collections::HashMap::new();
    let part = if pkg.has_part("ppt/theme/theme1.xml") {
        Some("ppt/theme/theme1.xml".to_string())
    } else {
        pkg.part_names()
            .find(|p| p.starts_with("ppt/theme/theme"))
            .map(|p| p.to_string())
    };
    let Some(part) = part else { return map };
    let Ok(theme) = pkg.xml(&part) else { return map };
    let Some(scheme) = theme
        .child("themeElements")
        .and_then(|te| te.child("clrScheme"))
    else {
        return map;
    };
    for entry in scheme.elements() {
        let color = entry
            .child("srgbClr")
            .and_then(|c| c.attr_local("val"))
            .or_else(|| entry.child("sysClr").and_then(|c| c.attr_local("lastClr")));
        if let Some(color) = color {
            map.insert(entry.local_name().to_string(), color.to_uppercase());
        }
    }
    map
}

/// Apply DrawingML color transforms (lumMod/lumOff/shade/tint) approximately.
/// RGB (0..255) to HSL (h in 0..360, s/l in 0..1).
fn rgb_to_hsl(r: f64, g: f64, b: f64) -> (f64, f64, f64) {
    let (r, g, b) = (r / 255.0, g / 255.0, b / 255.0);
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    if (max - min).abs() < 1e-9 {
        return (0.0, 0.0, l);
    }
    let d = max - min;
    let s = if l > 0.5 { d / (2.0 - max - min) } else { d / (max + min) };
    let h = if max == r {
        ((g - b) / d).rem_euclid(6.0)
    } else if max == g {
        (b - r) / d + 2.0
    } else {
        (r - g) / d + 4.0
    } * 60.0;
    (h, s, l)
}

fn hsl_to_rgb(h: f64, s: f64, l: f64) -> (f64, f64, f64) {
    if s <= 0.0 {
        let v = l * 255.0;
        return (v, v, v);
    }
    let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
    let p = 2.0 * l - q;
    let hk = h / 360.0;
    let channel = |mut t: f64| -> f64 {
        t = t.rem_euclid(1.0);
        let v = if t < 1.0 / 6.0 {
            p + (q - p) * 6.0 * t
        } else if t < 0.5 {
            q
        } else if t < 2.0 / 3.0 {
            p + (q - p) * (2.0 / 3.0 - t) * 6.0
        } else {
            p
        };
        v * 255.0
    };
    (channel(hk + 1.0 / 3.0), channel(hk), channel(hk - 1.0 / 3.0))
}

/// Apply DrawingML color transforms. lumMod/lumOff operate on HSL
/// luminance (matching Office's "Lighter/Darker N%" palette variants);
/// shade/tint are per-channel blends toward black/white.
fn apply_color_mods(hex: &str, clr: &XmlElement) -> String {
    let Ok(n) = u32::from_str_radix(hex, 16) else {
        return hex.to_string();
    };
    let mut rgb = [
        ((n >> 16) & 0xFF) as f64,
        ((n >> 8) & 0xFF) as f64,
        (n & 0xFF) as f64,
    ];
    for m in clr.elements() {
        let val = m
            .attr_local("val")
            .and_then(|v| v.parse::<f64>().ok())
            .map(|v| v / 100_000.0);
        let Some(val) = val else { continue };
        match m.local_name() {
            "lumMod" | "lumOff" => {
                let (h, s, mut l) = rgb_to_hsl(rgb[0], rgb[1], rgb[2]);
                if m.local_name() == "lumMod" {
                    l *= val;
                } else {
                    l += val;
                }
                let (r, g, b) = hsl_to_rgb(h, s, l.clamp(0.0, 1.0));
                rgb = [r, g, b];
            }
            "shade" => {
                for c in rgb.iter_mut() {
                    *c *= val;
                }
            }
            "tint" => {
                for c in rgb.iter_mut() {
                    *c = *c * val + 255.0 * (1.0 - val);
                }
            }
            _ => {}
        }
    }
    format!(
        "{:02X}{:02X}{:02X}",
        rgb[0].round().clamp(0.0, 255.0) as u8,
        rgb[1].round().clamp(0.0, 255.0) as u8,
        rgb[2].round().clamp(0.0, 255.0) as u8
    )
}

/// Resolve a fill-holding element (spPr, bgPr, rPr) to a hex color:
/// literal srgbClr or theme schemeClr (with tx/bg aliases).
fn resolve_fill_color(
    container: &XmlElement,
    theme: &std::collections::HashMap<String, String>,
) -> Option<String> {
    let fill = container.child("solidFill")?;
    resolve_color_choice(fill, theme)
}

/// Resolve the color child of any DrawingML color container.
fn resolve_color_choice(
    parent: &XmlElement,
    theme: &std::collections::HashMap<String, String>,
) -> Option<String> {
    if let Some(c) = parent.child("srgbClr") {
        let hex = c.attr_local("val")?.to_uppercase();
        return Some(apply_color_mods(&hex, c));
    }
    if let Some(c) = parent.child("schemeClr") {
        let name = c.attr_local("val")?;
        // clrMap aliases used on slides.
        let mapped = match name {
            "tx1" => "dk1",
            "bg1" => "lt1",
            "tx2" => "dk2",
            "bg2" => "lt2",
            other => other,
        };
        let base = theme.get(mapped)?;
        return Some(apply_color_mods(base, c));
    }
    None
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
                let rels_part = slide_rels_part(&part);
                let slide_rels = if pkg.has_part(&rels_part) {
                    pkg.xml(&rels_part)?
                } else {
                    crate::xml::parse(EMPTY_RELS.as_bytes())?
                };
                slides.push(Slide {
                    part,
                    rid,
                    xml,
                    rels: slide_rels,
                });
            }
        }
        let theme = load_theme_colors(&pkg);
        Ok(Pptx {
            pkg,
            presentation,
            rels,
            slides,
            theme,
        })
    }

    /// Slide background color, resolving theme references (bgPr and bgRef).
    fn slide_bg_color(&self, slide: &Slide) -> Option<String> {
        let bg = slide.xml.child("cSld")?.child("bg")?;
        if let Some(bgpr) = bg.child("bgPr") {
            return resolve_fill_color(bgpr, &self.theme);
        }
        if let Some(bgref) = bg.child("bgRef") {
            return resolve_color_choice(bgref, &self.theme);
        }
        None
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
        if let Some(xfrm) = shape_xfrm(e) {
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
        let slide_rels = crate::xml::parse(PPTX_SLIDE_RELS.as_bytes())?;

        let mut xml = crate::xml::parse(pptx_blank_slide().as_bytes())?;

        // Optional background color.
        if let Some(background) = props.get("background") {
            set_slide_background(&mut xml, &parse_color(background)?)?;
        }

        let slide = Slide {
            part,
            rid,
            xml,
            rels: slide_rels,
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

    /// chartSpace XML behind a chart graphicFrame, resolved through the
    /// slide rels (used by the screenshot renderer).
    fn chart_space_for_frame(&self, slide: &Slide, frame: &XmlElement) -> Option<XmlElement> {
        let rid = frame
            .child("graphic")?
            .child("graphicData")?
            .child("chart")?
            .attr_local("id")?;
        let target = slide
            .rels
            .children_named("Relationship")
            .into_iter()
            .find(|r| r.attr_local("Id") == Some(rid))?
            .attr_local("Target")?
            .to_string();
        self.pkg.xml(&resolve_target("ppt/slides", &target)).ok()
    }

    /// Notes slide part linked from a slide, if any.
    fn notes_part_for_slide(slide: &Slide) -> Option<String> {
        slide
            .rels
            .children_named("Relationship")
            .into_iter()
            .find(|r| r.attr_local("Type") == Some(NOTES_SLIDE_REL_TYPE))
            .and_then(|r| r.attr_local("Target"))
            .map(|t| resolve_target("ppt/slides", t))
    }

    /// Speaker notes text of a slide, if any.
    fn notes_text(&self, slide: &Slide) -> Option<String> {
        let part = Self::notes_part_for_slide(slide)?;
        let xml = self.pkg.xml(&part).ok()?;
        let tree = xml.child("cSld")?.child("spTree")?;
        for sp in tree.children_named("sp") {
            let is_body = sp
                .child("nvSpPr")
                .and_then(|nv| nv.child("nvPr"))
                .and_then(|n| n.child("ph"))
                .map(|ph| ph.attr_local("type") == Some("body"))
                .unwrap_or(false);
            if is_body {
                let text = shape_text(sp);
                if !text.trim().is_empty() {
                    return Some(text);
                }
            }
        }
        None
    }

    /// The notes master (and its theme) exist once per package.
    fn ensure_notes_master(&mut self) -> Result<()> {
        if self.pkg.has_part(NOTES_MASTER_PART) {
            return Ok(());
        }
        let theme_num = self
            .pkg
            .part_names()
            .filter_map(|p| {
                p.strip_prefix("ppt/theme/theme")
                    .and_then(|s| s.strip_suffix(".xml"))
                    .and_then(|n| n.parse::<u32>().ok())
            })
            .max()
            .unwrap_or(0)
            + 1;
        let theme_part = format!("ppt/theme/theme{theme_num}.xml");
        self.pkg
            .put_raw(&theme_part, crate::templates::pptx_theme().as_bytes().to_vec());
        self.pkg
            .add_override(&theme_part, "application/vnd.openxmlformats-officedocument.theme+xml")?;
        self.pkg.put_raw(
            NOTES_MASTER_PART,
            crate::templates::pptx_notes_master().into_bytes(),
        );
        self.pkg.add_override(
            NOTES_MASTER_PART,
            "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
        )?;
        let master_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme{theme_num}.xml"/></Relationships>"#
        );
        self.pkg.put_raw(
            "ppt/notesMasters/_rels/notesMaster1.xml.rels",
            master_rels.into_bytes(),
        );

        let rid = crate::media::add_relationship(
            &mut self.rels,
            NOTES_MASTER_REL_TYPE,
            "notesMasters/notesMaster1.xml",
        );
        if self.presentation.child("notesMasterIdLst").is_none() {
            let mut lst = XmlElement::new("p:notesMasterIdLst");
            lst.push(el("p:notesMasterId", &[("r:id", rid.as_str())]));
            // Schema order: right after sldMasterIdLst.
            let at = self
                .presentation
                .children
                .iter()
                .position(|n| {
                    matches!(n, XmlNode::Element(e) if e.local_name() == "sldMasterIdLst")
                })
                .map(|i| i + 1)
                .unwrap_or(0);
            self.presentation.children.insert(at, XmlNode::Element(lst));
        }
        Ok(())
    }

    /// Create or replace a slide's speaker notes.
    fn set_notes(&mut self, slide_idx: usize, text: &str) -> Result<()> {
        self.ensure_notes_master()?;
        let part = match Self::notes_part_for_slide(&self.slides[slide_idx]) {
            Some(part) => part,
            None => {
                let n = self
                    .pkg
                    .part_names()
                    .filter_map(|p| {
                        p.strip_prefix("ppt/notesSlides/notesSlide")
                            .and_then(|s| s.strip_suffix(".xml"))
                            .and_then(|n| n.parse::<u32>().ok())
                    })
                    .max()
                    .unwrap_or(0)
                    + 1;
                let part = format!("ppt/notesSlides/notesSlide{n}.xml");
                self.pkg.add_override(
                    &part,
                    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
                )?;
                let slide_file = self.slides[slide_idx]
                    .part
                    .rsplit('/')
                    .next()
                    .unwrap()
                    .to_string();
                let rels = format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{NOTES_MASTER_REL_TYPE}" Target="../notesMasters/notesMaster1.xml"/><Relationship Id="rId2" Type="{SLIDE_REL_TYPE}" Target="../slides/{slide_file}"/></Relationships>"#
                );
                self.pkg.put_raw(
                    &format!("ppt/notesSlides/_rels/notesSlide{n}.xml.rels"),
                    rels.into_bytes(),
                );
                crate::media::add_relationship(
                    &mut self.slides[slide_idx].rels,
                    NOTES_SLIDE_REL_TYPE,
                    &format!("../notesSlides/notesSlide{n}.xml"),
                );
                part
            }
        };
        self.pkg.put_xml(&part, &build_notes_xml(text))?;
        Ok(())
    }

    /// Insert a table as a graphicFrame with an a:tbl.
    fn add_table(&mut self, slide_idx: usize, props: &Props) -> Result<Report> {
        let data: Vec<Vec<String>> = props
            .get("data")
            .map(|d| crate::xlsx::parse_csv(&d.replace("\\n", "\n")))
            .unwrap_or_default();
        let rows: usize = props
            .get("rows")
            .map(|v| v.parse())
            .transpose()
            .context("rows must be a number")?
            .unwrap_or_else(|| data.len().max(2));
        let cols: usize = props
            .get("cols")
            .map(|v| v.parse())
            .transpose()
            .context("cols must be a number")?
            .unwrap_or_else(|| data.iter().map(|r| r.len()).max().unwrap_or(2));
        if rows == 0 || cols == 0 || rows > 500 || cols > 50 {
            bail!("table size {rows}x{cols} out of range");
        }
        let header = props.get_bool("header")?.unwrap_or(true);
        let x = props.get("x").map(parse_emu).transpose()?.unwrap_or(914_400);
        let y = props.get("y").map(parse_emu).transpose()?.unwrap_or(1_600_200);
        let w = props
            .get("w")
            .or_else(|| props.get("width"))
            .map(parse_emu)
            .transpose()?
            .unwrap_or(7_315_200);
        let h = props
            .get("h")
            .or_else(|| props.get("height"))
            .map(parse_emu)
            .transpose()?
            .unwrap_or((rows as i64) * 370_840);

        let mut tbl = XmlElement::new("a:tbl");
        let mut tblpr = el("a:tblPr", &[("bandRow", "1")]);
        if header {
            tblpr.set_attr("firstRow", "1");
        }
        tbl.push(tblpr);
        let mut grid = XmlElement::new("a:tblGrid");
        let col_w = (w / cols as i64).to_string();
        for _ in 0..cols {
            grid.push(el("a:gridCol", &[("w", col_w.as_str())]));
        }
        tbl.push(grid);
        let row_h = (h / rows as i64).to_string();
        for r in 0..rows {
            let mut tr = el("a:tr", &[("h", row_h.as_str())]);
            for c in 0..cols {
                let text = data
                    .get(r)
                    .and_then(|row| row.get(c))
                    .cloned()
                    .unwrap_or_default();
                let is_header = header && r == 0;
                let mut tc = XmlElement::new("a:tc");
                let mut tx = XmlElement::new("a:txBody");
                tx.push(XmlElement::new("a:bodyPr"));
                tx.push(XmlElement::new("a:lstStyle"));
                let mut p = XmlElement::new("a:p");
                let mut run = XmlElement::new("a:r");
                let mut rpr = el("a:rPr", &[("lang", "en-US")]);
                if is_header {
                    rpr.set_attr("b", "1");
                    rpr.push(solid_fill("FFFFFF"));
                }
                run.push(rpr);
                let mut t = XmlElement::new("a:t");
                t.push_text(&text);
                run.push(t);
                p.push(run);
                tx.push(p);
                tc.push(tx);
                let mut tcpr = XmlElement::new("a:tcPr");
                for side in ["a:lnL", "a:lnR", "a:lnT", "a:lnB"] {
                    let mut ln = el(side, &[("w", "12700")]);
                    ln.push(solid_fill("999999"));
                    tcpr.push(ln);
                }
                if is_header {
                    tcpr.push(solid_fill("4472C4"));
                }
                tc.push(tcpr);
                tr.push(tc);
            }
            tbl.push(tr);
        }

        let tree = Self::sp_tree(&self.slides[slide_idx])?;
        let id = Self::next_shape_id(tree);
        let name = props
            .get("name")
            .map(|s| s.to_string())
            .unwrap_or_else(|| format!("Table {id}"));
        let mut frame = XmlElement::new("p:graphicFrame");
        let mut nv = XmlElement::new("p:nvGraphicFramePr");
        nv.push(el(
            "p:cNvPr",
            &[("id", id.to_string().as_str()), ("name", name.as_str())],
        ));
        nv.push(XmlElement::new("p:cNvGraphicFramePr"));
        nv.push(XmlElement::new("p:nvPr"));
        frame.push(nv);
        let mut xfrm = XmlElement::new("p:xfrm");
        xfrm.push(el(
            "a:off",
            &[("x", x.to_string().as_str()), ("y", y.to_string().as_str())],
        ));
        xfrm.push(el(
            "a:ext",
            &[("cx", w.to_string().as_str()), ("cy", h.to_string().as_str())],
        ));
        frame.push(xfrm);
        let mut graphic = XmlElement::new("a:graphic");
        let mut gdata = el(
            "a:graphicData",
            &[("uri", "http://schemas.openxmlformats.org/drawingml/2006/table")],
        );
        gdata.push(tbl);
        graphic.push(gdata);
        frame.push(graphic);

        let tree = Self::sp_tree_mut(&mut self.slides[slide_idx])?;
        tree.push(frame);
        let tree = Self::sp_tree(&self.slides[slide_idx])?;
        let shape_idxs = Self::shape_indices(tree);
        let e = tree.children[*shape_idxs.last().unwrap()].as_element().unwrap();
        let spath = format!("/slide[{}]/shape[{}]", slide_idx + 1, shape_idxs.len());
        let mut info = self.shape_info(e, &spath);
        info.attr("table", format!("{rows}x{cols}"));
        Ok(Report::Nodes(vec![info]))
    }

    /// Insert a chart as a graphicFrame. Data comes from props
    /// (categories/values/series, plus values2/series2, ... for more
    /// series); the chart part stores it as cached literals.
    fn add_chart(&mut self, slide_idx: usize, props: &Props) -> Result<Report> {
        let kind = crate::chart::parse_kind(props.get("kind").unwrap_or("column"))?;
        let mut series = crate::chart::inline_series(kind, props)?;
        // Embed a real workbook with the data so "Edit Data" opens a sheet.
        let workbook = crate::chart::embedded_workbook(&mut series)?;
        let mut chart_space =
            crate::chart::build_chart_space(kind, props.get("title"), &series)?;
        crate::chart::attach_external_data(&mut chart_space, "rId1");

        let chart_num = self
            .pkg
            .part_names()
            .filter_map(|p| {
                p.strip_prefix("ppt/charts/chart")
                    .and_then(|s| s.strip_suffix(".xml"))
                    .and_then(|n| n.parse::<u32>().ok())
            })
            .max()
            .unwrap_or(0)
            + 1;
        let chart_part = format!("ppt/charts/chart{chart_num}.xml");
        self.pkg.put_xml(&chart_part, &chart_space)?;
        self.pkg
            .add_override(&chart_part, crate::chart::CHART_CONTENT_TYPE)?;
        let emb_part = format!("ppt/embeddings/Microsoft_Excel_Worksheet{chart_num}.xlsx");
        self.pkg.put_raw(&emb_part, workbook);
        self.pkg.add_default("xlsx", crate::chart::XLSX_CONTENT_TYPE)?;
        self.pkg.put_xml(
            &format!("ppt/charts/_rels/chart{chart_num}.xml.rels"),
            &crate::chart::chart_rels_xml(&format!(
                "../embeddings/Microsoft_Excel_Worksheet{chart_num}.xlsx"
            )),
        )?;
        let rid = crate::media::add_relationship(
            &mut self.slides[slide_idx].rels,
            crate::chart::CHART_REL_TYPE,
            &format!("../charts/chart{chart_num}.xml"),
        );

        let x = props.get("x").map(parse_emu).transpose()?.unwrap_or(914_400);
        let y = props.get("y").map(parse_emu).transpose()?.unwrap_or(1_371_600);
        let w = props
            .get("w")
            .or_else(|| props.get("width"))
            .map(parse_emu)
            .transpose()?
            .unwrap_or(7_315_200);
        let h = props
            .get("h")
            .or_else(|| props.get("height"))
            .map(parse_emu)
            .transpose()?
            .unwrap_or(4_114_800);

        let tree = Self::sp_tree(&self.slides[slide_idx])?;
        let id = Self::next_shape_id(tree);
        let name = props
            .get("name")
            .map(|s| s.to_string())
            .unwrap_or_else(|| format!("Chart {chart_num}"));

        let mut frame = XmlElement::new("p:graphicFrame");
        let mut nv = XmlElement::new("p:nvGraphicFramePr");
        nv.push(el(
            "p:cNvPr",
            &[("id", id.to_string().as_str()), ("name", name.as_str())],
        ));
        nv.push(XmlElement::new("p:cNvGraphicFramePr"));
        nv.push(XmlElement::new("p:nvPr"));
        frame.push(nv);
        let mut xfrm = XmlElement::new("p:xfrm");
        xfrm.push(el(
            "a:off",
            &[("x", x.to_string().as_str()), ("y", y.to_string().as_str())],
        ));
        xfrm.push(el(
            "a:ext",
            &[("cx", w.to_string().as_str()), ("cy", h.to_string().as_str())],
        ));
        frame.push(xfrm);
        let mut graphic = XmlElement::new("a:graphic");
        let mut gdata = el(
            "a:graphicData",
            &[("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart")],
        );
        gdata.push(el(
            "c:chart",
            &[
                ("xmlns:c", "http://schemas.openxmlformats.org/drawingml/2006/chart"),
                (
                    "xmlns:r",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                ),
                ("r:id", rid.as_str()),
            ],
        ));
        graphic.push(gdata);
        frame.push(graphic);

        let tree = Self::sp_tree_mut(&mut self.slides[slide_idx])?;
        tree.push(frame);
        let tree = Self::sp_tree(&self.slides[slide_idx])?;
        let shape_idxs = Self::shape_indices(tree);
        let e = tree.children[*shape_idxs.last().unwrap()].as_element().unwrap();
        let spath = format!("/slide[{}]/shape[{}]", slide_idx + 1, shape_idxs.len());
        let mut info = self.shape_info(e, &spath);
        info.attr("chart", props.get("kind").unwrap_or("column"));
        info.attr("series", series.len().to_string());
        Ok(Report::Nodes(vec![info]))
    }
}

// -------------------------------------------- transitions / animations ----

/// Insert `child` into a p:sld respecting schema order:
/// cSld, clrMapOvr, transition, timing.
fn insert_slide_child(slide_xml: &mut XmlElement, child: XmlElement) {
    let order = ["cSld", "clrMapOvr", "transition", "timing", "extLst"];
    let rank = |name: &str| order.iter().position(|o| *o == name).unwrap_or(order.len());
    let child_rank = rank(child.local_name());
    let mut at = slide_xml.children.len();
    for (i, node) in slide_xml.children.iter().enumerate() {
        if let XmlNode::Element(e) = node {
            if rank(e.local_name()) > child_rank {
                at = i;
                break;
            }
        }
    }
    slide_xml.children.insert(at, XmlNode::Element(child));
}

/// Build a `p:transition` from props. Returns None for `transition=none`.
fn build_transition(props: &Props) -> Result<Option<XmlElement>> {
    let kind = props.get("transition").unwrap().to_ascii_lowercase();
    if kind == "none" {
        return Ok(None);
    }
    // element name, dir-attribute style it accepts.
    let (name, dir_kind): (&str, &str) = match kind.as_str() {
        "fade" => ("p:fade", ""),
        "cut" => ("p:cut", ""),
        "dissolve" => ("p:dissolve", ""),
        "random" => ("p:random", ""),
        "circle" => ("p:circle", ""),
        "diamond" => ("p:diamond", ""),
        "plus" => ("p:plus", ""),
        "wedge" => ("p:wedge", ""),
        "newsflash" => ("p:newsflash", ""),
        "wheel" => ("p:wheel", ""),
        "push" => ("p:push", "lrud"),
        "wipe" => ("p:wipe", "lrud"),
        "cover" => ("p:cover", "lrud8"),
        "pull" => ("p:pull", "lrud8"),
        "zoom" => ("p:zoom", "inout"),
        "split" => ("p:split", "inout"),
        "blinds" => ("p:blinds", "hv"),
        "checker" => ("p:checker", "hv"),
        "comb" => ("p:comb", "hv"),
        "strips" => ("p:strips", "corners"),
        other => bail!(
            "unknown transition '{other}' (fade/cut/push/wipe/dissolve/circle/diamond/plus/wedge/wheel/zoom/cover/pull/split/blinds/checker/comb/strips/newsflash/random/none)"
        ),
    };
    let mut effect = XmlElement::new(name);
    if let Some(direction) = props.get("direction") {
        let d = direction.to_ascii_lowercase();
        let mapped = match (dir_kind, d.as_str()) {
            ("lrud", "left") | ("lrud8", "left") => Some("l"),
            ("lrud", "right") | ("lrud8", "right") => Some("r"),
            ("lrud", "up") | ("lrud8", "up") => Some("u"),
            ("lrud", "down") | ("lrud8", "down") => Some("d"),
            ("hv", "horizontal") => Some("horz"),
            ("hv", "vertical") => Some("vert"),
            ("inout", "in") => Some("in"),
            ("inout", "out") => Some("out"),
            ("corners", "left-down") => Some("ld"),
            ("corners", "left-up") => Some("lu"),
            ("corners", "right-down") => Some("rd"),
            ("corners", "right-up") => Some("ru"),
            _ => None,
        };
        match mapped {
            Some(m) => effect.set_attr("dir", m),
            None => bail!("direction '{direction}' does not apply to transition '{kind}'"),
        }
    }
    let mut transition = XmlElement::new("p:transition");
    if let Some(speed) = props.get("speed").or_else(|| props.get("duration")) {
        let spd = match speed.to_ascii_lowercase().as_str() {
            "slow" => "slow",
            "medium" | "med" => "med",
            "fast" => "fast",
            other => {
                // A duration is mapped onto PowerPoint's three speeds.
                let ms = parse_duration_ms(other)?;
                if ms <= 500 { "fast" } else if ms <= 1000 { "med" } else { "slow" }
            }
        };
        transition.set_attr("spd", spd);
    }
    if let Some(advance) = props.get("advance") {
        let ms = parse_duration_ms(advance)?;
        transition.set_attr("advTm", &ms.to_string());
    }
    transition.push(effect);
    Ok(Some(transition))
}

/// "500ms", "1.5s", or bare milliseconds.
fn parse_duration_ms(v: &str) -> Result<u64> {
    let v = v.trim();
    if let Some(s) = v.strip_suffix("ms") {
        return s.trim().parse().map_err(|_| anyhow::anyhow!("'{v}' is not a duration"));
    }
    if let Some(s) = v.strip_suffix('s') {
        let secs: f64 = s.trim().parse().map_err(|_| anyhow::anyhow!("'{v}' is not a duration"))?;
        return Ok((secs * 1000.0).round() as u64);
    }
    v.parse().map_err(|_| anyhow::anyhow!("'{v}' is not a duration (use 500ms, 1.5s, or milliseconds)"))
}

fn max_ctn_id(e: &XmlElement) -> u64 {
    let mut max = 0;
    if e.local_name() == "cTn" {
        if let Some(id) = e.attr_local("id").and_then(|v| v.parse::<u64>().ok()) {
            max = id;
        }
    }
    for c in e.elements() {
        max = max.max(max_ctn_id(c));
    }
    max
}

fn cond(delay: &str) -> XmlElement {
    el("p:cond", &[("delay", delay)])
}

/// Get (creating if needed) the mainSeq childTnLst of the slide's timing tree.
fn ensure_timing_main_seq(slide_xml: &mut XmlElement) -> Result<&mut XmlElement> {
    if slide_xml.child("timing").is_none() {
        let mut timing = XmlElement::new("p:timing");
        let mut tn_lst = XmlElement::new("p:tnLst");
        let mut par = XmlElement::new("p:par");
        let mut root_ctn = el(
            "p:cTn",
            &[("id", "1"), ("dur", "indefinite"), ("restart", "never"), ("nodeType", "tmRoot")],
        );
        let mut root_children = XmlElement::new("p:childTnLst");
        let mut seq = el("p:seq", &[("concurrent", "1"), ("nextAc", "seek")]);
        let mut main_ctn = el(
            "p:cTn",
            &[("id", "2"), ("dur", "indefinite"), ("nodeType", "mainSeq")],
        );
        main_ctn.push(XmlElement::new("p:childTnLst"));
        seq.push(main_ctn);
        let mut prev = XmlElement::new("p:prevCondLst");
        let mut pc = el("p:cond", &[("evt", "onPrev"), ("delay", "0")]);
        let mut tgt = XmlElement::new("p:tgtEl");
        tgt.push(XmlElement::new("p:sldTgt"));
        pc.push(tgt);
        prev.push(pc);
        seq.push(prev);
        let mut next = XmlElement::new("p:nextCondLst");
        let mut nc = el("p:cond", &[("evt", "onNext"), ("delay", "0")]);
        let mut tgt = XmlElement::new("p:tgtEl");
        tgt.push(XmlElement::new("p:sldTgt"));
        nc.push(tgt);
        next.push(nc);
        seq.push(next);
        root_children.push(seq);
        root_ctn.push(root_children);
        par.push(root_ctn);
        tn_lst.push(par);
        timing.push(tn_lst);
        insert_slide_child(slide_xml, timing);
    }
    slide_xml
        .child_mut("timing")
        .and_then(|t| t.child_mut("tnLst"))
        .and_then(|t| t.child_mut("par"))
        .and_then(|p| p.child_mut("cTn"))
        .and_then(|c| c.child_mut("childTnLst"))
        .and_then(|c| c.child_mut("seq"))
        .and_then(|s| s.child_mut("cTn"))
        .and_then(|c| c.child_mut("childTnLst"))
        .context("timing tree has no main sequence (was it created by another tool?)")
}

fn sp_target(spid: &str) -> XmlElement {
    let mut tgt = XmlElement::new("p:tgtEl");
    tgt.push(el("p:spTgt", &[("spid", spid)]));
    tgt
}

/// A `p:anim` motion behavior interpolating one position attribute
/// (ppt_x/ppt_y) from an off-slide formula to the shape's own position.
fn motion_anim(spid: &str, attr_name: &str, from: &str, to: &str, duration_ms: u64) -> XmlElement {
    let mut anim = el("p:anim", &[("calcmode", "lin"), ("valueType", "num")]);
    let mut cbhvr = el("p:cBhvr", &[("additive", "base")]);
    cbhvr.push(el(
        "p:cTn",
        &[("id", "0"), ("dur", duration_ms.to_string().as_str()), ("fill", "hold")],
    ));
    cbhvr.push(sp_target(spid));
    let mut attrs = XmlElement::new("p:attrNameLst");
    let mut a = XmlElement::new("p:attrName");
    a.push_text(attr_name);
    attrs.push(a);
    cbhvr.push(attrs);
    anim.push(cbhvr);
    let mut tavs = XmlElement::new("p:tavLst");
    for (tm, val) in [("0", from), ("100000", to)] {
        let mut tav = el("p:tav", &[("tm", tm)]);
        let mut v = XmlElement::new("p:val");
        v.push(el("p:strVal", &[("val", val)]));
        tav.push(v);
        tavs.push(tav);
    }
    anim.push(tavs);
    anim
}

/// Append one click-triggered entrance effect for shape `spid`.
fn append_entrance_animation(
    slide_xml: &mut XmlElement,
    spid: &str,
    effect: &str,
    direction: &str,
    duration_ms: u64,
    delay_ms: u64,
) -> Result<()> {
    // (presetID, animEffect filter) — appear has no filter behavior;
    // fly-in uses motion behaviors instead of a filter.
    let effect_lc = effect.to_ascii_lowercase();
    let fly_in = matches!(effect_lc.as_str(), "fly-in" | "flyin" | "fly");
    let (preset_id, filter): (u32, Option<String>) = match effect_lc.as_str() {
        "appear" => (1, None),
        "fade" | "fade-in" | "fadein" => (10, Some("fade".to_string())),
        "wipe" | "wipe-in" => (22, Some("wipe(bottom)".to_string())),
        "fly-in" | "flyin" | "fly" => (2, None),
        other => bail!("unknown animation '{other}' (appear/fade/wipe/fly-in)"),
    };
    // Fly-in start position and PowerPoint's UI subtype code per direction.
    let (preset_subtype, fly_from): (u32, (&str, &str)) =
        match direction.to_ascii_lowercase().as_str() {
            "bottom" | "up" => (4, ("#ppt_x", "1+#ppt_h/2")),
            "top" | "down" => (1, ("#ppt_x", "0-#ppt_h/2")),
            "left" => (8, ("0-#ppt_w/2", "#ppt_y")),
            "right" => (2, ("1+#ppt_w/2", "#ppt_y")),
            other => bail!("unknown direction '{other}' (left/right/top/bottom)"),
        };
    let preset_subtype = if fly_in { preset_subtype } else { 0 };
    let mut next_id = max_ctn_id(slide_xml.child("timing").unwrap_or(&XmlElement::new("x"))) + 1;
    if next_id < 3 {
        next_id = 3;
    }
    let main = ensure_timing_main_seq(slide_xml)?;
    let mut id = max_ctn_id(main).max(next_id - 1) + 1;
    let mut next = || {
        let v = id;
        id += 1;
        v.to_string()
    };

    // Effect behaviors.
    let mut behaviors: Vec<XmlElement> = Vec::new();
    let mut set = XmlElement::new("p:set");
    let mut cbhvr = XmlElement::new("p:cBhvr");
    let mut ctn = el("p:cTn", &[("id", "0"), ("dur", "1"), ("fill", "hold")]);
    let mut st = XmlElement::new("p:stCondLst");
    st.push(cond("0"));
    ctn.push(st);
    cbhvr.push(ctn);
    cbhvr.push(sp_target(spid));
    let mut attrs = XmlElement::new("p:attrNameLst");
    let mut attr = XmlElement::new("p:attrName");
    attr.push_text("style.visibility");
    attrs.push(attr);
    cbhvr.push(attrs);
    set.push(cbhvr);
    let mut to = XmlElement::new("p:to");
    let mut sv = el("p:strVal", &[("val", "visible")]);
    to.push(std::mem::take(&mut sv));
    set.push(to);
    behaviors.push(set);
    if let Some(filter) = &filter {
        let mut anim = el(
            "p:animEffect",
            &[("transition", "in"), ("filter", filter.as_str())],
        );
        let mut cbhvr = XmlElement::new("p:cBhvr");
        cbhvr.push(el("p:cTn", &[("id", "0"), ("dur", duration_ms.to_string().as_str())]));
        cbhvr.push(sp_target(spid));
        anim.push(cbhvr);
        behaviors.push(anim);
    }
    if fly_in {
        let (fx, fy) = fly_from;
        behaviors.push(motion_anim(spid, "ppt_x", fx, "#ppt_x", duration_ms));
        behaviors.push(motion_anim(spid, "ppt_y", fy, "#ppt_y", duration_ms));
    }

    // Click group scaffolding, innermost first.
    let mut effect_ctn = el(
        "p:cTn",
        &[
            ("id", "0"),
            ("presetID", preset_id.to_string().as_str()),
            ("presetClass", "entr"),
            ("presetSubtype", preset_subtype.to_string().as_str()),
            ("fill", "hold"),
            ("nodeType", "clickEffect"),
        ],
    );
    let mut st = XmlElement::new("p:stCondLst");
    st.push(cond(&delay_ms.to_string()));
    effect_ctn.push(st);
    let mut children = XmlElement::new("p:childTnLst");
    for b in behaviors {
        children.push(b);
    }
    effect_ctn.push(children);
    let mut effect_par = XmlElement::new("p:par");
    effect_par.push(effect_ctn);

    let mut group_ctn = el("p:cTn", &[("id", "0"), ("fill", "hold")]);
    let mut st = XmlElement::new("p:stCondLst");
    st.push(cond("0"));
    group_ctn.push(st);
    let mut children = XmlElement::new("p:childTnLst");
    children.push(effect_par);
    group_ctn.push(children);
    let mut group_par = XmlElement::new("p:par");
    group_par.push(group_ctn);

    let mut click_ctn = el("p:cTn", &[("id", "0"), ("fill", "hold")]);
    let mut st = XmlElement::new("p:stCondLst");
    st.push(cond("indefinite"));
    click_ctn.push(st);
    let mut children = XmlElement::new("p:childTnLst");
    children.push(group_par);
    click_ctn.push(children);
    let mut click_par = XmlElement::new("p:par");
    click_par.push(click_ctn);

    // Assign unique ids to every cTn we created (outermost first for
    // stable numbering).
    fn assign_ids(e: &mut XmlElement, next: &mut dyn FnMut() -> String) {
        if e.local_name() == "cTn" && e.attr_local("id") == Some("0") {
            let id = next();
            e.set_attr("id", &id);
        }
        for c in e.children.iter_mut() {
            if let XmlNode::Element(ce) = c {
                assign_ids(ce, next);
            }
        }
    }
    assign_ids(&mut click_par, &mut next);
    main.push(click_par);
    Ok(())
}

fn shape_cnvpr(e: &XmlElement) -> Option<&XmlElement> {
    // nvSpPr / nvPicPr / nvGraphicFramePr / nvCxnSpPr / nvGrpSpPr → cNvPr
    e.elements()
        .find(|c| c.local_name().starts_with("nv") && c.local_name().ends_with("Pr"))
        .and_then(|nv| nv.child("cNvPr"))
}


/// Shape geometry: sp/pic keep xfrm under spPr, graphicFrames directly.
fn shape_xfrm(e: &XmlElement) -> Option<&XmlElement> {
    e.child("spPr")
        .and_then(|sp| sp.child("xfrm"))
        .or_else(|| e.child("xfrm"))
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

/// The a:tbl inside a graphicFrame, if it holds a table.
fn frame_table(e: &XmlElement) -> Option<&XmlElement> {
    e.child("graphic")?.child("graphicData")?.child("tbl")
}

fn frame_table_mut(e: &mut XmlElement) -> Option<&mut XmlElement> {
    e.child_mut("graphic")?
        .child_mut("graphicData")?
        .child_mut("tbl")
}

/// Replace every cell text in a table from CSV-shaped data.
fn set_table_data(tbl: &mut XmlElement, data: &str) -> (usize, usize) {
    let grid = crate::xlsx::parse_csv(&data.replace("\\n", "\n"));
    let mut rows = 0;
    let mut cols = 0;
    let mut r = 0;
    for node in tbl.children.iter_mut().filter_map(|n| n.as_element_mut()) {
        if node.local_name() != "tr" {
            continue;
        }
        let mut c = 0;
        for tc in node.children.iter_mut().filter_map(|n| n.as_element_mut()) {
            if tc.local_name() != "tc" {
                continue;
            }
            let text = grid
                .get(r)
                .and_then(|row| row.get(c))
                .cloned()
                .unwrap_or_default();
            if let Some(t) = tc
                .child_mut("txBody")
                .and_then(|tx| tx.child_mut("p"))
                .and_then(|p| p.child_mut("r"))
                .and_then(|run| run.child_mut("t"))
            {
                t.children.clear();
                t.push_text(&text);
            }
            c += 1;
        }
        cols = cols.max(c);
        r += 1;
        rows = r;
    }
    (rows, cols)
}

/// All text in a shape's txBody, paragraphs joined with '\n'. Tables render
/// as one line per row with cells joined by ' | '.
fn shape_text(e: &XmlElement) -> String {
    if let Some(tbl) = frame_table(e) {
        let mut lines = Vec::new();
        for tr in tbl.children_named("tr") {
            let cells: Vec<String> = tr
                .children_named("tc")
                .iter()
                .map(|tc| tc.text_content())
                .collect();
            lines.push(cells.join(" | "));
        }
        return lines.join("\n");
    }
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

/// Attach `a:hlinkClick` to every run in the shape's txBody.
fn apply_hyperlink(sp: &mut XmlElement, rid: &str) {
    let Some(tx) = sp.child_mut("txBody") else { return };
    for p in tx.children.iter_mut().filter_map(|n| n.as_element_mut()) {
        if p.local_name() != "p" {
            continue;
        }
        for r in p.children.iter_mut().filter_map(|n| n.as_element_mut()) {
            if r.local_name() != "r" {
                continue;
            }
            let rpr = r.ensure_child("rPr", "a:rPr", true);
            if rpr.attr("lang").is_none() {
                rpr.set_attr("lang", "en-US");
            }
            rpr.children.retain(|n| {
                !matches!(n, XmlNode::Element(e) if e.local_name() == "hlinkClick")
            });
            rpr.push(el("a:hlinkClick", &[("r:id", rid)]));
        }
    }
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
    // Leading tabs select the list indent level.
    let line = line.replace("\\t", "\t");
    let level = line.chars().take_while(|c| *c == '\t').count().min(8);
    let text = line.trim_start_matches('\t');

    let list_kind = props
        .get("list")
        .map(|k| k.to_ascii_lowercase())
        .filter(|k| k != "none");
    let algn = props
        .get("align")
        .map(|a| parse_align(a).map(|(_, algn)| algn))
        .transpose()?;
    if list_kind.is_some() || algn.is_some() || level > 0 {
        let mut ppr = XmlElement::new("a:pPr");
        if let Some(algn) = algn {
            ppr.set_attr("algn", algn);
        }
        if let Some(kind) = &list_kind {
            if level > 0 {
                ppr.set_attr("lvl", &level.to_string());
            }
            // Hanging indent so wrapped lines align after the marker.
            let marl = 285_750 + 457_200 * level as i64;
            ppr.set_attr("marL", &marl.to_string());
            ppr.set_attr("indent", "-285750");
            ppr.push(el(
                "a:buFont",
                &[("typeface", "Arial"), ("pitchFamily", "34"), ("charset", "0")],
            ));
            match kind.as_str() {
                "bullet" | "bullets" | "ul" => ppr.push(el("a:buChar", &[("char", "•")])),
                "number" | "numbered" | "decimal" | "ol" => {
                    ppr.push(el("a:buAutoNum", &[("type", "arabicPeriod")]))
                }
                other => bail!("unknown list kind '{other}' (bullet/number/none)"),
            }
        } else if level > 0 {
            ppr.set_attr("lvl", &level.to_string());
        }
        p.push(ppr);
    }
    let mut r = XmlElement::new("a:r");
    if let Some(rpr) = build_run_props(props)? {
        r.push(rpr);
    }
    let mut t = XmlElement::new("a:t");
    t.push_text(text);
    r.push(t);
    p.push(r);
    Ok(p)
}

/// "• " or "N. " for a rendered a:p with bullet properties.
fn slide_list_marker(p: &XmlElement, counters: &mut std::collections::HashMap<u32, u32>) -> Option<String> {
    let ppr = p.child("pPr")?;
    let level: u32 = ppr.attr_local("lvl").and_then(|v| v.parse().ok()).unwrap_or(0);
    if ppr.child("buChar").is_some() {
        Some("• ".to_string())
    } else if ppr.child("buAutoNum").is_some() {
        let c = counters.entry(level).or_insert(0);
        *c += 1;
        Some(format!("{c}. "))
    } else {
        None
    }
}

/// A complete notesSlide part with the text in the body placeholder.
fn build_notes_xml(text: &str) -> XmlElement {
    let mut notes = el(
        "p:notes",
        &[
            ("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
            ("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
            ("xmlns:p", "http://schemas.openxmlformats.org/presentationml/2006/main"),
        ],
    );
    let mut csld = XmlElement::new("p:cSld");
    let mut tree = XmlElement::new("p:spTree");
    let mut nv = XmlElement::new("p:nvGrpSpPr");
    nv.push(el("p:cNvPr", &[("id", "1"), ("name", "")]));
    nv.push(XmlElement::new("p:cNvGrpSpPr"));
    nv.push(XmlElement::new("p:nvPr"));
    tree.push(nv);
    tree.push(XmlElement::new("p:grpSpPr"));

    let mut sp = XmlElement::new("p:sp");
    let mut nvsp = XmlElement::new("p:nvSpPr");
    nvsp.push(el("p:cNvPr", &[("id", "2"), ("name", "Notes Placeholder 1")]));
    let mut cnv = XmlElement::new("p:cNvSpPr");
    cnv.push(el("a:spLocks", &[("noGrp", "1")]));
    nvsp.push(cnv);
    let mut nvpr = XmlElement::new("p:nvPr");
    nvpr.push(el("p:ph", &[("type", "body"), ("idx", "1")]));
    nvsp.push(nvpr);
    sp.push(nvsp);
    sp.push(XmlElement::new("p:spPr"));
    let mut tx = XmlElement::new("p:txBody");
    tx.push(XmlElement::new("a:bodyPr"));
    tx.push(XmlElement::new("a:lstStyle"));
    for line in text.replace("\\n", "\n").split('\n') {
        let mut p = XmlElement::new("a:p");
        let mut r = XmlElement::new("a:r");
        let mut t = XmlElement::new("a:t");
        t.push_text(line);
        r.push(t);
        p.push(r);
        tx.push(p);
    }
    sp.push(tx);
    tree.push(sp);
    csld.push(tree);
    notes.push(csld);
    let mut cmo = XmlElement::new("p:clrMapOvr");
    cmo.push(XmlElement::new("a:masterClrMapping"));
    notes.push(cmo);
    notes
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
            "notes" => {
                let mut nodes = Vec::new();
                for (i, slide) in self.slides.iter().enumerate() {
                    if let Some(text) = self.notes_text(slide) {
                        let mut info = NodeInfo::new(format!("/slide[{}]", i + 1), "notes");
                        info.text = Some(text);
                        nodes.push(info);
                    }
                }
                Ok(Report::Nodes(nodes))
            }
            "html" => {
                // Render at 960px wide; 12192000 EMU (16:9 default) → 960px,
                // which conveniently makes font px = a:rPr sz / 100.
                let (sld_cx, sld_cy) = self
                    .presentation
                    .child("sldSz")
                    .and_then(|s| {
                        Some((
                            s.attr_local("cx")?.parse::<f64>().ok()?,
                            s.attr_local("cy")?.parse::<f64>().ok()?,
                        ))
                    })
                    .unwrap_or((12_192_000.0, 6_858_000.0));
                let scale = 960.0 / sld_cx;
                let slide_h = (sld_cy * scale).round() as i64;
                let mut out = String::new();
                for (i, slide) in self.slides.iter().enumerate() {
                    out.push_str(&format!("<div class=\"slide-label\">Slide {}</div>\n", i + 1));
                    let bg_css = self
                        .slide_bg_color(slide)
                        .map(|c| format!("background:#{c};"))
                        .unwrap_or_default();
                    out.push_str(&format!(
                        "<div class=\"slide\" style=\"width:960px;height:{slide_h}px;{bg_css}\">\n"
                    ));
                    let tree = Self::sp_tree(slide)?;
                    for &si in &Self::shape_indices(tree) {
                        let e = tree.children[si].as_element().unwrap();
                        render_shape_html(e, scale, &self.theme, &mut out);
                    }
                    out.push_str("</div>\n");
                }
                Ok(Report::Text(crate::html::page("Presentation", &out)))
            }
            other => bail!("unknown view mode '{other}' for pptx (text/outline/stats/html)"),
        }
    }

    fn get(&mut self, path_str: &str, depth: usize, _computed: bool) -> Result<Report> {
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
                if let Some(notes) = props.get("notes") {
                    let notes = notes.to_string();
                    self.set_notes(idx, &notes)?;
                }
                Ok(Report::Nodes(vec![self.slide_info(idx, 1)?]))
            }
            "shape" | "textbox" | "text" => {
                if dpath.segments.is_empty() {
                    bail!("shapes are added to a slide: officecli add file.pptx '/slide[1]' --type shape");
                }
                let slide_idx = self.slide_index(&dpath.segments[0])?;
                let tree = Self::sp_tree(&self.slides[slide_idx])?;
                let id = Self::next_shape_id(tree);
                let mut sp = self.build_textbox(id, props)?;
                if let Some(url) = props.get("url") {
                    let rid = crate::media::add_external_relationship(
                        &mut self.slides[slide_idx].rels,
                        crate::media::HYPERLINK_REL_TYPE,
                        url,
                    );
                    apply_hyperlink(&mut sp, &rid);
                }
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
            "image" | "picture" => {
                if dpath.segments.is_empty() {
                    bail!("images are added to a slide: officecli add file.pptx '/slide[1]' --type image --prop src=...");
                }
                let slide_idx = self.slide_index(&dpath.segments[0])?;
                let image = crate::media::image_from_props(props)?;
                let part = crate::media::store_image(&mut self.pkg, "ppt/media", &image)?;
                let target = format!("../media/{}", part.rsplit('/').next().unwrap());
                let rid = crate::media::add_relationship(
                    &mut self.slides[slide_idx].rels,
                    crate::media::IMAGE_REL_TYPE,
                    &target,
                );
                let x = props.get("x").map(parse_emu).transpose()?.unwrap_or(914_400);
                let y = props.get("y").map(parse_emu).transpose()?.unwrap_or(914_400);
                let w = props
                    .get("w")
                    .or_else(|| props.get("width"))
                    .map(parse_emu)
                    .transpose()?
                    .unwrap_or(image.width_emu);
                let h = match props
                    .get("h")
                    .or_else(|| props.get("height"))
                    .map(parse_emu)
                    .transpose()?
                {
                    Some(h) => h,
                    None if w != image.width_emu && image.width_emu > 0 => {
                        image.height_emu * w / image.width_emu
                    }
                    None => image.height_emu,
                };
                let tree = Self::sp_tree(&self.slides[slide_idx])?;
                let id = Self::next_shape_id(tree);
                let name = props
                    .get("name")
                    .map(|s| s.to_string())
                    .unwrap_or_else(|| format!("Picture {id}"));

                let mut pic = XmlElement::new("p:pic");
                let mut nv = XmlElement::new("p:nvPicPr");
                nv.push(el(
                    "p:cNvPr",
                    &[("id", id.to_string().as_str()), ("name", name.as_str())],
                ));
                nv.push(XmlElement::new("p:cNvPicPr"));
                nv.push(XmlElement::new("p:nvPr"));
                pic.push(nv);
                let mut fill = XmlElement::new("p:blipFill");
                fill.push(el("a:blip", &[("r:embed", rid.as_str())]));
                let mut stretch = XmlElement::new("a:stretch");
                stretch.push(XmlElement::new("a:fillRect"));
                fill.push(stretch);
                pic.push(fill);
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
                pic.push(sppr);

                let tree = Self::sp_tree_mut(&mut self.slides[slide_idx])?;
                tree.push(pic);
                let tree = Self::sp_tree(&self.slides[slide_idx])?;
                let shape_idxs = Self::shape_indices(tree);
                let e = tree.children[*shape_idxs.last().unwrap()].as_element().unwrap();
                let spath = format!("/slide[{}]/shape[{}]", slide_idx + 1, shape_idxs.len());
                Ok(Report::Nodes(vec![self.shape_info(e, &spath)]))
            }
            "chart" => {
                if dpath.segments.is_empty() {
                    bail!("charts are added to a slide: officecli add file.pptx '/slide[1]' --type chart --prop values=...");
                }
                let slide_idx = self.slide_index(&dpath.segments[0])?;
                self.add_chart(slide_idx, props)
            }
            "table" | "tbl" => {
                if dpath.segments.is_empty() {
                    bail!("tables are added to a slide: officecli add file.pptx '/slide[1]' --type table --prop data=\"a,b\\nc,d\"");
                }
                let slide_idx = self.slide_index(&dpath.segments[0])?;
                self.add_table(slide_idx, props)
            }
            other => bail!("unsupported pptx element type '{other}' (slide/shape/image/chart)"),
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
            let mut changed: Vec<String> = Vec::new();
            if let Some(background) = props.get("background") {
                let color = parse_color(background)?;
                set_slide_background(&mut self.slides[slide_idx].xml, &color)?;
                changed.push(format!("background={color}"));
            }
            if props.has("transition") {
                let slide_xml = &mut self.slides[slide_idx].xml;
                slide_xml.children.retain(|n| {
                    !matches!(n, XmlNode::Element(e) if e.local_name() == "transition")
                });
                if let Some(transition) = build_transition(props)? {
                    insert_slide_child(slide_xml, transition);
                }
                changed.push(format!("transition={}", props.get("transition").unwrap()));
            }
            if let Some(notes) = props.get("notes") {
                let notes = notes.to_string();
                self.set_notes(slide_idx, &notes)?;
                changed.push("notes".to_string());
            }
            if changed.is_empty() {
                bail!("slide-level set supports --prop background=COLOR, --prop notes=\"...\", and --prop transition=fade|push|... [--prop direction=..] [--prop speed=..] [--prop advance=..]");
            }
            return Ok(Report::Data {
                text: format!("slide {}: set {}", slide_idx + 1, changed.join(", ")),
                data: json!({ "slide": slide_idx + 1, "set": changed }),
            });
        }

        let shape_child_idx = self.find_shape(slide_idx, &dpath.segments[1])?;

        // Entrance animation (extends the slide's timing tree).
        if let Some(effect) = props.get("animation") {
            let spid = {
                let tree = Self::sp_tree(&self.slides[slide_idx])?;
                let e = tree.children[shape_child_idx].as_element().unwrap();
                shape_cnvpr(e)
                    .and_then(|c| c.attr_local("id"))
                    .map(|s| s.to_string())
                    .context("shape has no stable id to animate")?
            };
            let duration = props
                .get("duration")
                .map(parse_duration_ms)
                .transpose()?
                .unwrap_or(500);
            let delay = props.get("delay").map(parse_duration_ms).transpose()?.unwrap_or(0);
            let direction = props.get("direction").unwrap_or("bottom").to_string();
            append_entrance_animation(
                &mut self.slides[slide_idx].xml,
                &spid,
                effect,
                &direction,
                duration,
                delay,
            )?;
        }

        // Hyperlink rel must be created before the shape borrow below.
        let link_rid = match props.get("url") {
            Some(url) => Some(crate::media::add_external_relationship(
                &mut self.slides[slide_idx].rels,
                crate::media::HYPERLINK_REL_TYPE,
                url,
            )),
            None => None,
        };

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

        // Table data replacement.
        if let Some(data) = props.get("data") {
            match frame_table_mut(sp) {
                Some(tbl) => {
                    let (rows, cols) = set_table_data(tbl, data);
                    let _ = (rows, cols);
                }
                None => bail!("--prop data= applies to table shapes"),
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
        if let Some(rid) = &link_rid {
            apply_hyperlink(sp, rid);
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
            // Notes slide goes with its slide.
            if let Some(notes_part) = Self::notes_part_for_slide(&self.slides[slide_idx]) {
                let notes_rels =
                    notes_part.replace("ppt/notesSlides/", "ppt/notesSlides/_rels/") + ".rels";
                let mut ct = self.pkg.xml(CONTENT_TYPES_PART)?;
                let part_name = format!("/{notes_part}");
                ct.children.retain(|n| {
                    !matches!(n, XmlNode::Element(e)
                        if e.local_name() == "Override"
                            && e.attr_local("PartName") == Some(&part_name))
                });
                self.pkg.put_xml(CONTENT_TYPES_PART, &ct)?;
                self.pkg.remove_part(&notes_part);
                self.pkg.remove_part(&notes_rels);
            }
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
            self.pkg.put_xml(&slide_rels_part(&slide.part), &slide.rels)?;
        }
        self.pkg.save(path)
    }

    fn tree(&mut self) -> Result<Vec<NodeInfo>> {
        let mut roots = Vec::new();
        for i in 0..self.slides.len() {
            roots.push(self.slide_info(i, 1)?);
        }
        Ok(roots)
    }

    fn move_el(&mut self, path_str: &str, _to: Option<&str>, pos: &Position) -> Result<Report> {
        let dpath = path::parse(path_str)?;
        let slide_idx = self.slide_index(&dpath.segments[0])?;

        if dpath.segments.len() == 1 {
            // Reorder slides.
            let to = match pos {
                Position::Index(n) => (*n).min(self.slides.len() - 1),
                Position::Before(p) | Position::After(p) => {
                    let anchor =
                        self.slide_index(&path::parse(p)?.segments[0])? as i64;
                    let after = matches!(pos, Position::After(_)) as i64;
                    // Position after removing the source slide.
                    let mut t = anchor + after;
                    if (slide_idx as i64) < t {
                        t -= 1;
                    }
                    t.max(0) as usize
                }
                Position::Append => self.slides.len() - 1,
            };
            let slide = self.slides.remove(slide_idx);
            let rid = slide.rid.clone();
            self.slides.insert(to.min(self.slides.len()), slide);
            // Mirror in sldIdLst.
            if let Some(lst) = self.presentation.child_mut("sldIdLst") {
                let from_raw = lst.children.iter().position(|n| {
                    matches!(n, XmlNode::Element(e)
                        if e.local_name() == "sldId"
                            && (e.attr("r:id") == Some(&rid)
                                || e.attr_local("id") == Some(rid.as_str())))
                });
                if let Some(i) = from_raw {
                    let entry = lst.children.remove(i);
                    // Raw index of the `to`-th sldId (or end).
                    let mut seen = 0;
                    let mut insert_at = lst.children.len();
                    for (ci, n) in lst.children.iter().enumerate() {
                        if matches!(n, XmlNode::Element(e) if e.local_name() == "sldId") {
                            if seen == to {
                                insert_at = ci;
                                break;
                            }
                            seen += 1;
                        }
                    }
                    lst.children.insert(insert_at, entry);
                }
            }
            return Ok(Report::Data {
                text: format!("moved slide {} to position {}", slide_idx + 1, to + 1),
                data: json!({ "moved": format!("/slide[{}]", slide_idx + 1), "position": to + 1 }),
            });
        }

        // Reorder a shape within its slide.
        let shape_child_idx = self.find_shape(slide_idx, &dpath.segments[1])?;
        let tree = Self::sp_tree_mut(&mut self.slides[slide_idx])?;
        let node = tree.children.remove(shape_child_idx);
        let shape_idxs = Self::shape_indices(tree);
        let insert_at = match pos {
            Position::Index(n) => shape_idxs
                .get(*n)
                .copied()
                .unwrap_or(tree.children.len()),
            Position::Append => tree.children.len(),
            _ => bail!("shape move supports --index N (0-based among shapes)"),
        };
        tree.children.insert(insert_at, node);
        Ok(Report::Data {
            text: format!("moved shape on slide {}", slide_idx + 1),
            data: json!({ "moved": path_str }),
        })
    }

    fn copy_el(&mut self, path_str: &str, pos: &Position) -> Result<Report> {
        let dpath = path::parse(path_str)?;
        if dpath.is_root() {
            bail!("cannot copy the presentation root");
        }
        let slide_idx = self.slide_index(&dpath.segments[0])?;

        // Duplicate a whole slide (content, rels, background, notes).
        if dpath.segments.len() == 1 {
            let src_xml = self.slides[slide_idx].xml.clone();
            let mut src_rels = self.slides[slide_idx].rels.clone();
            // A notes slide belongs to exactly one slide; the copy gets its
            // own below.
            src_rels.children.retain(|n| {
                !matches!(n, XmlNode::Element(e)
                    if e.attr_local("Type") == Some(NOTES_SLIDE_REL_TYPE))
            });
            let notes = self.notes_text(&self.slides[slide_idx]);

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
            let rid = crate::media::add_relationship(
                &mut self.rels,
                SLIDE_REL_TYPE,
                &format!("slides/slide{next_part_num}.xml"),
            );
            self.pkg.add_override(&part, SLIDE_CONTENT_TYPE)?;

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
            // Default: right after the original.
            let insert_at = match pos {
                Position::Append => sld_id_lst
                    .nth_child_index("sldId", slide_idx)
                    .map(|i| i + 1)
                    .unwrap_or(sld_id_lst.children.len()),
                Position::Index(n) => {
                    let count = sld_id_lst.children_named("sldId").len();
                    sld_id_lst
                        .nth_child_index("sldId", (*n).min(count.saturating_sub(1)))
                        .unwrap_or(sld_id_lst.children.len())
                }
                _ => bail!("copy slide supports the default position (after the original) or --index N"),
            };
            let entry = el(
                "p:sldId",
                &[
                    ("id", next_sld_id.to_string().as_str()),
                    ("r:id", rid.as_str()),
                ],
            );
            sld_id_lst.children.insert(insert_at, XmlNode::Element(entry));
            let slide_pos = {
                let lst = self.presentation.child("sldIdLst").unwrap();
                lst.children[..insert_at]
                    .iter()
                    .filter(|n| matches!(n, XmlNode::Element(e) if e.local_name() == "sldId"))
                    .count()
            };
            self.slides.insert(
                slide_pos,
                Slide {
                    part,
                    rid,
                    xml: src_xml,
                    rels: src_rels,
                },
            );
            if let Some(notes) = notes {
                self.set_notes(slide_pos, &notes)?;
            }
            let mut info = self.slide_info(slide_pos, 1)?;
            info.attr("copied-from", path_str);
            return Ok(Report::Nodes(vec![info]));
        }

        // Duplicate a shape within its slide.
        let shape_child_idx = self.find_shape(slide_idx, &dpath.segments[1])?;
        let tree = Self::sp_tree(&self.slides[slide_idx])?;
        let mut clone = tree.children[shape_child_idx]
            .as_element()
            .context("shape vanished")?
            .clone();
        let new_id = Self::next_shape_id(tree);
        if let Some(cnvpr) = clone
            .children
            .iter_mut()
            .filter_map(|n| n.as_element_mut())
            .find(|e| e.local_name().starts_with("nv") && e.local_name().ends_with("Pr"))
            .and_then(|nv| nv.child_mut("cNvPr"))
        {
            cnvpr.set_attr("id", &new_id.to_string());
            let name = cnvpr.attr_local("name").unwrap_or("Shape").to_string();
            cnvpr.set_attr("name", &format!("{name} Copy"));
        }
        let insert_at = match pos {
            Position::Append => shape_child_idx + 1,
            Position::Index(n) => {
                let tree = Self::sp_tree(&self.slides[slide_idx])?;
                let idxs = Self::shape_indices(tree);
                idxs.get(*n).copied().unwrap_or(tree.children.len())
            }
            _ => bail!("copy shape supports the default position (after the original) or --index N"),
        };
        let tree = Self::sp_tree_mut(&mut self.slides[slide_idx])?;
        tree.children.insert(insert_at, XmlNode::Element(clone));
        let tree = Self::sp_tree(&self.slides[slide_idx])?;
        let e = tree.children[insert_at].as_element().unwrap();
        let position = Self::shape_indices(tree)
            .iter()
            .position(|&i| i == insert_at)
            .unwrap_or(0);
        let spath = format!("/slide[{}]/shape[{}]", slide_idx + 1, position + 1);
        let mut info = self.shape_info(e, &spath);
        info.attr("copied-from", path_str);
        Ok(Report::Nodes(vec![info]))
    }

    fn swap(&mut self, path1: &str, path2: &str) -> Result<Report> {
        let d1 = path::parse(path1)?;
        let d2 = path::parse(path2)?;
        let s1 = self.slide_index(&d1.segments[0])?;
        let s2 = self.slide_index(&d2.segments[0])?;
        match (d1.segments.len(), d2.segments.len()) {
            (1, 1) => {
                // Swap slides = swap their sldIdLst entries + vec order.
                self.slides.swap(s1, s2);
                if let Some(lst) = self.presentation.child_mut("sldIdLst") {
                    let raw: Vec<usize> = lst
                        .children
                        .iter()
                        .enumerate()
                        .filter(|(_, n)| {
                            matches!(n, XmlNode::Element(e) if e.local_name() == "sldId")
                        })
                        .map(|(i, _)| i)
                        .collect();
                    if let (Some(&a), Some(&b)) = (raw.get(s1), raw.get(s2)) {
                        lst.children.swap(a, b);
                    }
                }
            }
            (2, 2) => {
                let c1 = self.find_shape(s1, &d1.segments[1])?;
                let c2 = self.find_shape(s2, &d2.segments[1])?;
                if s1 == s2 {
                    let tree = Self::sp_tree_mut(&mut self.slides[s1])?;
                    tree.children.swap(c1, c2);
                } else {
                    let n1 = Self::sp_tree(&self.slides[s1])?.children[c1].clone();
                    let n2 = Self::sp_tree(&self.slides[s2])?.children[c2].clone();
                    Self::sp_tree_mut(&mut self.slides[s1])?.children[c1] = n2;
                    Self::sp_tree_mut(&mut self.slides[s2])?.children[c2] = n1;
                }
            }
            _ => bail!("swap needs two slides or two shapes, not a mix"),
        }
        Ok(Report::Data {
            text: format!("swapped {path1} and {path2}"),
            data: json!({ "swapped": [path1, path2] }),
        })
    }

    fn screenshot(&mut self) -> Result<Vec<Vec<u8>>> {
        use crate::render::{Canvas, Color, Span, BLACK, MUTED, WHITE};
        let (sld_cx, sld_cy) = self
            .presentation
            .child("sldSz")
            .and_then(|s| {
                Some((
                    s.attr_local("cx")?.parse::<f64>().ok()?,
                    s.attr_local("cy")?.parse::<f64>().ok()?,
                ))
            })
            .unwrap_or((12_192_000.0, 6_858_000.0));
        let width = 1280u32;
        let scale = width as f64 / sld_cx;
        let height = (sld_cy * scale).round() as u32;
        let px_per_pt = (scale * 12_700.0) as f32;

        let mut images = Vec::new();
        for slide in &self.slides {
            let bg = self
                .slide_bg_color(slide)
                .and_then(|hex| Color::from_hex(&hex))
                .unwrap_or(WHITE);
            // Dark backgrounds get light default text.
            let default_text = if (bg.r as u32 + bg.g as u32 + bg.b as u32) < 3 * 110 {
                WHITE
            } else {
                BLACK
            };
            let mut canvas = Canvas::new(width, height, bg)?;
            let tree = Self::sp_tree(slide)?;
            for &si in &Self::shape_indices(tree) {
                let e = tree.children[si].as_element().unwrap();
                let xfrm = shape_xfrm(e);
                let emu = |el: Option<&XmlElement>, a: &str| -> f32 {
                    el.and_then(|e| e.attr_local(a))
                        .and_then(|v| v.parse::<f64>().ok())
                        .map(|v| (v * scale) as f32)
                        .unwrap_or(0.0)
                };
                let x = emu(xfrm.and_then(|x| x.child("off")), "x");
                let y = emu(xfrm.and_then(|x| x.child("off")), "y");
                let w = emu(xfrm.and_then(|x| x.child("ext")), "cx");
                let h = emu(xfrm.and_then(|x| x.child("ext")), "cy");
                if let Some(fill) = e
                    .child("spPr")
                    .and_then(|sp| resolve_fill_color(sp, &self.theme))
                    .and_then(|hex| Color::from_hex(&hex))
                {
                    canvas.fill_rect(x, y, w, h, fill);
                }
                match e.local_name() {
                    "pic" => {
                        let bytes = self
                            .dump_picture(slide, e)
                            .and_then(|p| p.get("srcdata").and_then(|v| v.as_str()).map(String::from))
                            .and_then(|b64| {
                                use base64::Engine;
                                base64::engine::general_purpose::STANDARD.decode(b64).ok()
                            })
                            .unwrap_or_default();
                        canvas.draw_image(&bytes, x, y, w.max(2.0), h.max(2.0));
                    }
                    "graphicFrame" => {
                        if let Some(tbl) = frame_table(e) {
                            let trs = tbl.children_named("tr");
                            let rows = trs.len().max(1);
                            let cols = trs
                                .first()
                                .map(|tr| tr.children_named("tc").len())
                                .unwrap_or(1)
                                .max(1);
                            let (cw, rh) = (w / cols as f32, h / rows as f32);
                            let font = (rh * 0.45).min(16.0);
                            for (r, tr) in trs.iter().enumerate() {
                                for (c, tc) in tr.children_named("tc").iter().enumerate() {
                                    let (cx, cy) = (x + cw * c as f32, y + rh * r as f32);
                                    let fill = tc
                                        .child("tcPr")
                                        .and_then(|p| resolve_fill_color(p, &self.theme))
                                        .and_then(|hex| Color::from_hex(&hex));
                                    if let Some(fill) = fill {
                                        canvas.fill_rect(cx, cy, cw, rh, fill);
                                    }
                                    canvas.stroke_rect(cx, cy, cw, rh, MUTED);
                                    let text = tc.text_content();
                                    let color = tc
                                        .child("txBody")
                                        .and_then(|tx| tx.child("p"))
                                        .and_then(|p| p.child("r"))
                                        .and_then(|run| run.child("rPr"))
                                        .and_then(|rpr| resolve_fill_color(rpr, &self.theme))
                                        .and_then(|hex| Color::from_hex(&hex))
                                        .unwrap_or(default_text);
                                    canvas.draw_text(
                                        &text,
                                        cx + 6.0,
                                        cy + rh / 2.0 + font / 3.0,
                                        font,
                                        color,
                                        false,
                                    );
                                }
                            }
                        } else {
                            let drawn = self
                                .chart_space_for_frame(slide, e)
                                .map(|space| {
                                    crate::chartdraw::draw_chart(&mut canvas, &space, x, y, w, h)
                                })
                                .unwrap_or(false);
                            if !drawn {
                                canvas.stroke_rect(x, y, w, h, MUTED);
                                canvas.draw_text("[chart]", x + 8.0, y + 20.0, 14.0, MUTED, false);
                            }
                        }
                    }
                    _ => {
                        let Some(tx) = e.child("txBody") else { continue };
                        let mut top = y;
                        let pad = 4.0;
                        let mut bullet_counters: std::collections::HashMap<u32, u32> =
                            Default::default();
                        for p in tx.children_named("p") {
                            let algn = p
                                .child("pPr")
                                .and_then(|pr| pr.attr_local("algn"))
                                .unwrap_or("l");
                            let list_level: u32 = p
                                .child("pPr")
                                .and_then(|pr| pr.attr_local("lvl"))
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0);
                            let marker = slide_list_marker(p, &mut bullet_counters);
                            let mut spans: Vec<Span> = Vec::new();
                            for r in p.children_named("r") {
                                let rpr = r.child("rPr");
                                let size_pt = rpr
                                    .and_then(|rp| rp.attr_local("sz"))
                                    .and_then(|v| v.parse::<f32>().ok())
                                    .map(|v| v / 100.0)
                                    .unwrap_or(18.0);
                                let color = rpr
                                    .and_then(|rp| resolve_fill_color(rp, &self.theme))
                                    .and_then(|hex| Color::from_hex(&hex))
                                    .unwrap_or(default_text);
                                let bold =
                                    rpr.map(|rp| rp.attr_local("b") == Some("1")).unwrap_or(false);
                                let italic =
                                    rpr.map(|rp| rp.attr_local("i") == Some("1")).unwrap_or(false);
                                let text =
                                    r.child("t").map(|t| t.text_content()).unwrap_or_default();
                                spans.push(Span {
                                    text,
                                    size: size_pt * px_per_pt,
                                    color,
                                    bold,
                                    italic,
                                });
                            }
                            if spans.is_empty() {
                                top += 18.0 * px_per_pt * 1.25;
                                continue;
                            }
                            let mut extra_indent = 0.0f32;
                            if let Some(marker) = marker {
                                let size = spans.first().map(|s| s.size).unwrap_or(18.0);
                                let color = spans.first().map(|s| s.color).unwrap_or(default_text);
                                extra_indent = size * 1.2 * list_level as f32;
                                spans.insert(0, Span { text: marker, size, color, bold: false, italic: false });
                            }
                            let max_size = spans.iter().map(|s| s.size).fold(12.0f32, f32::max);
                            let line_h = max_size * 1.25;
                            let usable = (w - 2.0 * pad - extra_indent).max(20.0);
                            for line in canvas.layout_spans(&spans, usable) {
                                let lw = canvas.spans_width(&line);
                                let lx = match algn {
                                    "ctr" => x + pad + (usable - lw) / 2.0,
                                    "r" => x + pad + usable - lw,
                                    _ => x + pad,
                                } + extra_indent;
                                let asc = canvas.ascent(max_size, false);
                                canvas.draw_spans_line(&line, lx, top + pad + asc);
                                top += line_h;
                            }
                        }
                    }
                }
            }
            images.push(canvas.png()?);
        }
        Ok(images)
    }

    fn dump(&mut self) -> Result<serde_json::Value> {
        let mut ops = Vec::new();
        for (i, slide) in self.slides.iter().enumerate() {
            let mut slide_props = serde_json::Map::new();
            if let Some(bg_color) = self.slide_bg_color(slide) {
                slide_props.insert("background".into(), json!(bg_color));
            }
            if let Some(notes) = self.notes_text(slide) {
                slide_props.insert("notes".into(), json!(notes));
            }
            ops.push(json!({
                "command": "add", "parent": "/", "type": "slide",
                "props": slide_props,
            }));
            if let Some(tprops) = dump_transition(&slide.xml) {
                ops.push(json!({
                    "command": "set",
                    "path": format!("/slide[{}]", i + 1),
                    "props": tprops,
                }));
            }
            let tree = Self::sp_tree(slide)?;
            for &si in &Self::shape_indices(tree) {
                let e = tree.children[si].as_element().unwrap();
                if e.local_name() == "pic" {
                    // Pictures replay with their bytes embedded.
                    if let Some(props) = self.dump_picture(slide, e) {
                        ops.push(json!({
                            "command": "add",
                            "parent": format!("/slide[{}]", i + 1),
                            "type": "image", "props": props,
                        }));
                    }
                    continue;
                }
                if e.local_name() != "sp" {
                    continue; // charts/groups are not replayable
                }
                let mut props = serde_json::Map::new();
                props.insert("text".into(), json!(shape_text(e)));
                if let Some(cnvpr) = shape_cnvpr(e) {
                    if let Some(name) = cnvpr.attr_local("name") {
                        props.insert("name".into(), json!(name));
                    }
                }
                if let Some(xfrm) = e.child("spPr").and_then(|sp| sp.child("xfrm")) {
                    if let Some(off) = xfrm.child("off") {
                        if let (Some(x), Some(y)) = (off.attr_local("x"), off.attr_local("y")) {
                            props.insert("x".into(), json!(x));
                            props.insert("y".into(), json!(y));
                        }
                    }
                    if let Some(ext) = xfrm.child("ext") {
                        if let (Some(cx), Some(cy)) =
                            (ext.attr_local("cx"), ext.attr_local("cy"))
                        {
                            props.insert("w".into(), json!(cx));
                            props.insert("h".into(), json!(cy));
                        }
                    }
                }
                if let Some(fill) = e
                    .child("spPr")
                    .and_then(|sp| resolve_fill_color(sp, &self.theme))
                {
                    props.insert("fill".into(), json!(fill));
                }
                if let Some(algn) = e
                    .child("txBody")
                    .and_then(|tx| tx.child("p"))
                    .and_then(|p| p.child("pPr"))
                    .and_then(|pr| pr.attr_local("algn"))
                {
                    let align = match algn {
                        "ctr" => "center",
                        "r" => "right",
                        "just" => "justify",
                        _ => "left",
                    };
                    props.insert("align".into(), json!(align));
                }
                // First run's formatting as a shape-level approximation.
                if let Some(rpr) = e
                    .child("txBody")
                    .and_then(|tx| tx.child("p"))
                    .and_then(|p| p.child("r"))
                    .and_then(|r| r.child("rPr"))
                {
                    if let Some(sz) = rpr.attr_local("sz") {
                        if let Ok(hundredths) = sz.parse::<f64>() {
                            props.insert("size".into(), json!(format!("{}", hundredths / 100.0)));
                        }
                    }
                    if rpr.attr_local("b") == Some("1") {
                        props.insert("bold".into(), json!("true"));
                    }
                    if let Some(color) = resolve_fill_color(rpr, &self.theme) {
                        props.insert("color".into(), json!(color));
                    }
                    if let Some(font) = rpr.child("latin").and_then(|l| l.attr_local("typeface"))
                    {
                        props.insert("font".into(), json!(font));
                    }
                }
                ops.push(json!({
                    "command": "add",
                    "parent": format!("/slide[{}]", i + 1),
                    "type": "shape",
                    "props": props,
                }));
            }
        }
        Ok(serde_json::Value::Array(ops))
    }
}

/// New slides live only in memory until save; treat them as present.
fn slide_is_new(_part: &str) -> bool {
    true
}

/// Reverse-map a p:transition into replayable set props.
fn dump_transition(slide_xml: &XmlElement) -> Option<serde_json::Map<String, serde_json::Value>> {
    let transition = slide_xml.child("transition")?;
    let effect = transition.elements().next()?;
    let mut props = serde_json::Map::new();
    props.insert("transition".into(), json!(effect.local_name()));
    if let Some(dir) = effect.attr_local("dir") {
        let word = match dir {
            "l" => "left",
            "r" => "right",
            "u" => "up",
            "d" => "down",
            "horz" => "horizontal",
            "vert" => "vertical",
            "in" => "in",
            "out" => "out",
            "ld" => "left-down",
            "lu" => "left-up",
            "rd" => "right-down",
            "ru" => "right-up",
            other => other,
        };
        props.insert("direction".into(), json!(word));
    }
    if let Some(spd) = transition.attr_local("spd") {
        let word = match spd {
            "med" => "medium",
            other => other,
        };
        props.insert("speed".into(), json!(word));
    }
    if let Some(adv) = transition.attr_local("advTm") {
        props.insert("advance".into(), json!(format!("{adv}ms")));
    }
    Some(props)
}

impl Pptx {
    /// Replayable `add image` props (bytes embedded as srcdata) for a p:pic.
    fn dump_picture(
        &self,
        slide: &Slide,
        pic: &XmlElement,
    ) -> Option<serde_json::Map<String, serde_json::Value>> {
        let rid = pic
            .child("blipFill")
            .and_then(|f| f.child("blip"))
            .and_then(|b| b.attr("r:embed").or_else(|| b.attr_local("embed")))?;
        let target = slide
            .rels
            .children_named("Relationship")
            .into_iter()
            .find(|r| r.attr_local("Id") == Some(rid))
            .and_then(|r| r.attr_local("Target"))?;
        let part = resolve_target("ppt/slides", target);
        let bytes = self.pkg.raw(&part).ok()?;
        use base64::Engine;
        let mut props = serde_json::Map::new();
        props.insert(
            "srcdata".into(),
            json!(base64::engine::general_purpose::STANDARD.encode(bytes)),
        );
        if let Some(cnvpr) = shape_cnvpr(pic) {
            if let Some(name) = cnvpr.attr_local("name") {
                props.insert("name".into(), json!(name));
            }
        }
        if let Some(xfrm) = pic.child("spPr").and_then(|sp| sp.child("xfrm")) {
            if let Some(off) = xfrm.child("off") {
                if let (Some(x), Some(y)) = (off.attr_local("x"), off.attr_local("y")) {
                    props.insert("x".into(), json!(x));
                    props.insert("y".into(), json!(y));
                }
            }
            if let Some(ext) = xfrm.child("ext") {
                if let (Some(cx), Some(cy)) = (ext.attr_local("cx"), ext.attr_local("cy")) {
                    props.insert("w".into(), json!(cx));
                    props.insert("h".into(), json!(cy));
                }
            }
        }
        Some(props)
    }
}

fn render_shape_html(
    e: &XmlElement,
    scale: f64,
    theme: &std::collections::HashMap<String, String>,
    out: &mut String,
) {
    let xfrm = shape_xfrm(e);
    let get = |el: Option<&XmlElement>, a: &str| -> f64 {
        el.and_then(|e| e.attr_local(a))
            .and_then(|v| v.parse::<f64>().ok())
            .unwrap_or(0.0)
    };
    let x = get(xfrm.and_then(|x| x.child("off")), "x") * scale;
    let y = get(xfrm.and_then(|x| x.child("off")), "y") * scale;
    let w = get(xfrm.and_then(|x| x.child("ext")), "cx") * scale;
    let h = get(xfrm.and_then(|x| x.child("ext")), "cy") * scale;
    let mut css = format!(
        "left:{:.0}px;top:{:.0}px;width:{:.0}px;height:{:.0}px;",
        x, y, w, h
    );
    if let Some(fill) = e.child("spPr").and_then(|sp| resolve_fill_color(sp, theme)) {
        css.push_str(&format!("background:#{fill};"));
    }
    out.push_str(&format!("<div class=\"shape\" style=\"{css}\">"));
    if let Some(tx) = e.child("txBody") {
        for p in tx.children_named("p") {
            let algn = p
                .child("pPr")
                .and_then(|pr| pr.attr_local("algn"))
                .unwrap_or("l");
            let align_css = match algn {
                "ctr" => "text-align:center;",
                "r" => "text-align:right;",
                "just" => "text-align:justify;",
                _ => "",
            };
            out.push_str(&format!("<div style=\"{align_css}\">"));
            for r in p.children_named("r") {
                let mut span = String::new();
                if let Some(rpr) = r.child("rPr") {
                    if let Some(sz) = rpr.attr_local("sz").and_then(|v| v.parse::<f64>().ok()) {
                        // At 960px slide width, px ≈ sz/100 (see above).
                        span.push_str(&format!("font-size:{:.0}px;", sz / 100.0));
                    }
                    if rpr.attr_local("b") == Some("1") {
                        span.push_str("font-weight:bold;");
                    }
                    if rpr.attr_local("i") == Some("1") {
                        span.push_str("font-style:italic;");
                    }
                    if let Some(color) = resolve_fill_color(rpr, theme) {
                        span.push_str(&format!("color:#{color};"));
                    }
                    if let Some(font) = rpr.child("latin").and_then(|l| l.attr_local("typeface")) {
                        span.push_str(&format!("font-family:'{font}';"));
                    }
                }
                let text = r
                    .child("t")
                    .map(|t| crate::html::escape(&t.text_content()))
                    .unwrap_or_default();
                if span.is_empty() {
                    out.push_str(&text);
                } else {
                    out.push_str(&format!("<span style=\"{span}\">{text}</span>"));
                }
            }
            out.push_str("</div>");
        }
    } else if e.local_name() == "pic" {
        out.push_str("<em style=\"color:#888\">[image]</em>");
    }
    out.push_str("</div>\n");
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

#[cfg(test)]
mod tests {
    use super::apply_color_mods;
    use crate::xml::el;

    #[test]
    fn color_mods_approximate_office() {
        // accent1 4472C4 at lumMod 75% — PowerPoint's "Darker 25%" is
        // 2F5496; HSL luminance math lands within a hair of it.
        let mut clr = el("a:schemeClr", &[("val", "accent1")]);
        clr.push(el("a:lumMod", &[("val", "75000")]));
        assert_eq!(apply_color_mods("4472C4", &clr), "2F5597");
        // lumMod 60% + lumOff 40% — PowerPoint's "Lighter 40%" is 8EAADB.
        let mut clr = el("a:schemeClr", &[("val", "accent1")]);
        clr.push(el("a:lumMod", &[("val", "60000")]));
        clr.push(el("a:lumOff", &[("val", "40000")]));
        assert_eq!(apply_color_mods("4472C4", &clr), "8FAADC");
        // No mods pass through.
        assert_eq!(apply_color_mods("FF0000", &el("a:srgbClr", &[])), "FF0000");
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
