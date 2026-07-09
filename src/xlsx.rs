//! Excel (.xlsx) handler: sheets, rows, cells, formulas, basic styling.
//!
//! Reads resolve shared strings; writes use inline strings (valid OOXML,
//! avoids shared-string table bookkeeping). Formulas are stored uncalculated
//! and the workbook is flagged `fullCalcOnLoad` so Excel/LibreOffice
//! recalculate on open.

use anyhow::{bail, Context, Result};
use regex::Regex;
use serde_json::json;
use std::path::Path;

use crate::handler::{Handler, Position};
use crate::out::{NodeInfo, Report};
use crate::path::{self, cell_name, col_letters, parse_cell_ref};
use crate::pkg::Package;
use crate::props::{parse_bool, parse_color, parse_pt, Props};
use crate::templates::XLSX_BLANK_SHEET;
use crate::xml::{el, XmlElement, XmlNode};

const WORKBOOK_PART: &str = "xl/workbook.xml";
const WORKBOOK_RELS_PART: &str = "xl/_rels/workbook.xml.rels";
const STYLES_PART: &str = "xl/styles.xml";
const CONTENT_TYPES_PART: &str = "[Content_Types].xml";

struct Sheet {
    name: String,
    part: String,
    xml: XmlElement,
}

pub struct Xlsx {
    pkg: Package,
    workbook: XmlElement,
    rels: XmlElement,
    styles: XmlElement,
    shared: Vec<String>,
    sheets: Vec<Sheet>,
}

impl Xlsx {
    pub fn new(pkg: Package) -> Result<Xlsx> {
        let workbook = pkg.xml(WORKBOOK_PART)?;
        let rels = pkg.xml(WORKBOOK_RELS_PART)?;
        let styles = pkg.xml(STYLES_PART)?;
        let shared = load_shared_strings(&pkg)?;

        let mut sheets = Vec::new();
        if let Some(sheets_el) = workbook.child("sheets") {
            for sheet in sheets_el.children_named("sheet") {
                let name = sheet.attr_local("name").unwrap_or("?").to_string();
                let rid = sheet
                    .attr("r:id")
                    .or_else(|| sheet.attr_local("id"))
                    .context("workbook sheet entry has no r:id")?;
                let target = rel_target(&rels, rid)
                    .with_context(|| format!("no workbook relationship '{rid}'"))?;
                let part = resolve_target("xl", &target);
                let xml = pkg.xml(&part)?;
                sheets.push(Sheet { name, part, xml });
            }
        }
        Ok(Xlsx {
            pkg,
            workbook,
            rels,
            styles,
            shared,
            sheets,
        })
    }

    fn sheet_index(&self, seg: &path::Segment) -> Result<usize> {
        // `/sheet[2]` positional form.
        if seg.name.eq_ignore_ascii_case("sheet") {
            let n = seg.index().unwrap_or(1);
            if n == 0 || n > self.sheets.len() {
                bail!("sheet index {n} out of range (1..{})", self.sheets.len());
            }
            return Ok(n - 1);
        }
        // `/Sheet1` by name.
        self.sheets
            .iter()
            .position(|s| s.name == seg.name)
            .with_context(|| {
                let names: Vec<&str> = self.sheets.iter().map(|s| s.name.as_str()).collect();
                format!("no sheet named '{}' (sheets: {})", seg.name, names.join(", "))
            })
    }

    fn cell_display(&self, c: &XmlElement) -> (String, &'static str) {
        let t = c.attr_local("t").unwrap_or("n");
        let v = c.child("v").map(|v| v.text_content()).unwrap_or_default();
        match t {
            "s" => {
                let idx: usize = v.trim().parse().unwrap_or(usize::MAX);
                (
                    self.shared.get(idx).cloned().unwrap_or_default(),
                    "string",
                )
            }
            "str" => (v, "string"),
            "inlineStr" => (
                c.child("is").map(|is| is.text_content()).unwrap_or_default(),
                "string",
            ),
            "b" => ((v.trim() == "1").to_string(), "boolean"),
            "e" => (v, "error"),
            _ => (v, "number"),
        }
    }

    fn cell_info(&self, sheet: &Sheet, c: &XmlElement, path: &str) -> NodeInfo {
        let mut info = NodeInfo::new(path, "cell");
        let (value, typ) = self.cell_display(c);
        info.text = Some(value);
        info.attr("type", typ);
        if let Some(f) = c.child("f") {
            info.attr("formula", format!("={}", f.text_content()));
        }
        if let Some(r) = c.attr_local("r") {
            info.attr("ref", format!("{}!{}", sheet.name, r));
        }
        info
    }

    /// Used range of a sheet: ((min_col,min_row),(max_col,max_row)) 0-based.
    fn used_range(sheet: &Sheet) -> Option<((u32, u32), (u32, u32))> {
        let sheet_data = sheet.xml.child("sheetData")?;
        let mut range: Option<((u32, u32), (u32, u32))> = None;
        for row in sheet_data.children_named("row") {
            for c in row.children_named("c") {
                let Some((col, r)) = c.attr_local("r").and_then(parse_cell_ref_opt) else {
                    continue;
                };
                range = Some(match range {
                    None => ((col, r), (col, r)),
                    Some(((c0, r0), (c1, r1))) => {
                        ((c0.min(col), r0.min(r)), (c1.max(col), r1.max(r)))
                    }
                });
            }
        }
        range
    }

    fn render_sheet_grid(&self, sheet: &Sheet) -> String {
        let mut out = String::new();
        let Some(((c0, r0), (c1, r1))) = Self::used_range(sheet) else {
            out.push_str(&format!("{} (empty)\n", sheet.name));
            return out;
        };
        out.push_str(&format!(
            "{} ({}:{})\n",
            sheet.name,
            cell_name(c0, r0),
            cell_name(c1, r1)
        ));
        // Header row with column letters.
        out.push('\t');
        let header: Vec<String> = (c0..=c1).map(col_letters).collect();
        out.push_str(&header.join("\t"));
        out.push('\n');
        let sheet_data = sheet.xml.child("sheetData").unwrap();
        for r in r0..=r1 {
            let mut cells = vec![String::new(); (c1 - c0 + 1) as usize];
            if let Some(row) = sheet_data
                .children_named("row")
                .into_iter()
                .find(|row| row.attr_local("r").map(|v| v == (r + 1).to_string()).unwrap_or(false))
            {
                for c in row.children_named("c") {
                    if let Some((col, _)) = c.attr_local("r").and_then(parse_cell_ref_opt) {
                        if col >= c0 && col <= c1 {
                            cells[(col - c0) as usize] = self.cell_display(c).0;
                        }
                    }
                }
            }
            out.push_str(&format!("{}\t{}\n", r + 1, cells.join("\t")));
        }
        out
    }

    fn set_cell(
        &mut self,
        sheet_idx: usize,
        col: u32,
        row: u32,
        props: &Props,
    ) -> Result<NodeInfo> {
        // Style props are computed against the current cell first.
        let style_props = StyleRequest::from_props(props)?;
        let new_style = if style_props.any() {
            let current_s = self.sheets[sheet_idx]
                .xml
                .child("sheetData")
                .and_then(|sd| find_row(sd, row + 1))
                .and_then(|r| find_cell(r, col, row))
                .and_then(|c| c.attr_local("s"))
                .and_then(|s| s.parse::<usize>().ok())
                .unwrap_or(0);
            Some(ensure_style(&mut self.styles, current_s, &style_props)?)
        } else {
            None
        };

        {
            let sheet = &mut self.sheets[sheet_idx];
            let sheet_data = sheet
                .xml
                .child_mut("sheetData")
                .context("worksheet has no <sheetData>")?;
            let c = ensure_cell(sheet_data, col, row);

            if let Some(value) = props.get("value") {
                write_cell_value(c, value, props.get("type"))?;
            } else if let Some(formula) = props.get("formula") {
                write_cell_formula(c, formula);
            }
            if let Some(s) = new_style {
                c.set_attr("s", &s.to_string());
            }
        }
        let sheet = &self.sheets[sheet_idx];
        let path = format!("/{}/{}", sheet.name, cell_name(col, row));
        let c = sheet
            .xml
            .child("sheetData")
            .and_then(|sd| find_row(sd, row + 1))
            .and_then(|r| find_cell(r, col, row))
            .context("internal: cell vanished after write")?;
        Ok(self.cell_info(sheet, c, &path))
    }

    fn save_sheets(&mut self) -> Result<()> {
        // Ensure recalculation of formulas on open.
        let calc = self.workbook.ensure_child("calcPr", "calcPr", false);
        calc.set_attr("fullCalcOnLoad", "1");
        self.pkg.put_xml(WORKBOOK_PART, &self.workbook)?;
        self.pkg.put_xml(WORKBOOK_RELS_PART, &self.rels)?;
        self.pkg.put_xml(STYLES_PART, &self.styles)?;
        for sheet in &self.sheets {
            self.pkg.put_xml(&sheet.part, &sheet.xml)?;
        }
        Ok(())
    }
}

fn parse_cell_ref_opt(r: &str) -> Option<(u32, u32)> {
    parse_cell_ref(r)
}

fn load_shared_strings(pkg: &Package) -> Result<Vec<String>> {
    if !pkg.has_part("xl/sharedStrings.xml") {
        return Ok(Vec::new());
    }
    let sst = pkg.xml("xl/sharedStrings.xml")?;
    Ok(sst
        .children_named("si")
        .into_iter()
        .map(|si| si.text_content())
        .collect())
}

fn rel_target(rels: &XmlElement, rid: &str) -> Option<String> {
    rels.children_named("Relationship")
        .into_iter()
        .find(|r| r.attr_local("Id") == Some(rid))
        .and_then(|r| r.attr_local("Target"))
        .map(|t| t.to_string())
}

/// Resolve a relationship target relative to a base directory.
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

fn find_row(sheet_data: &XmlElement, row_num: u32) -> Option<&XmlElement> {
    sheet_data
        .children_named("row")
        .into_iter()
        .find(|r| r.attr_local("r").map(|v| v == row_num.to_string()).unwrap_or(false))
}

fn find_cell<'a>(row: &'a XmlElement, col: u32, row0: u32) -> Option<&'a XmlElement> {
    let target = cell_name(col, row0);
    row.children_named("c")
        .into_iter()
        .find(|c| c.attr_local("r") == Some(target.as_str()))
}

/// Get or create the row (1-based `row_num` = row0+1) and cell, keeping both
/// sorted by position.
fn ensure_cell(sheet_data: &mut XmlElement, col: u32, row0: u32) -> &mut XmlElement {
    let row_num = row0 + 1;
    // Find or insert the row, keeping rows ordered by r.
    let row_pos = {
        let mut insert_at = sheet_data.children.len();
        let mut found = None;
        for (i, node) in sheet_data.children.iter().enumerate() {
            let XmlNode::Element(e) = node else { continue };
            if e.local_name() != "row" {
                continue;
            }
            let r: u32 = e
                .attr_local("r")
                .and_then(|v| v.parse().ok())
                .unwrap_or(u32::MAX);
            if r == row_num {
                found = Some(i);
                break;
            }
            if r > row_num {
                insert_at = i;
                break;
            }
        }
        match found {
            Some(i) => i,
            None => {
                let row = el("row", &[("r", row_num.to_string().as_str())]);
                sheet_data.children.insert(insert_at, XmlNode::Element(row));
                insert_at
            }
        }
    };
    let row = sheet_data.children[row_pos].as_element_mut().unwrap();
    let target = cell_name(col, row0);
    let cell_pos = {
        let mut insert_at = row.children.len();
        let mut found = None;
        for (i, node) in row.children.iter().enumerate() {
            let XmlNode::Element(e) = node else { continue };
            if e.local_name() != "c" {
                continue;
            }
            let c = e
                .attr_local("r")
                .and_then(parse_cell_ref_opt)
                .map(|(c, _)| c)
                .unwrap_or(u32::MAX);
            if c == col {
                found = Some(i);
                break;
            }
            if c > col {
                insert_at = i;
                break;
            }
        }
        match found {
            Some(i) => i,
            None => {
                let cell = el("c", &[("r", target.as_str())]);
                row.children.insert(insert_at, XmlNode::Element(cell));
                insert_at
            }
        }
    };
    row.children[cell_pos].as_element_mut().unwrap()
}

fn clear_cell_content(c: &mut XmlElement) {
    c.children.clear();
    c.remove_attr("t");
}

fn write_cell_value(c: &mut XmlElement, value: &str, type_hint: Option<&str>) -> Result<()> {
    clear_cell_content(c);
    let hint = type_hint.map(|t| t.to_ascii_lowercase());
    let hint = hint.as_deref();
    if hint == Some("formula") || (hint.is_none() && value.starts_with('=')) {
        write_cell_formula(c, value);
        return Ok(());
    }
    let as_number = value.parse::<f64>().is_ok() && !value.is_empty();
    match hint {
        Some("string") | Some("text") => write_inline_string(c, value),
        Some("number") => {
            if !as_number {
                bail!("'{value}' is not a number");
            }
            let mut v = XmlElement::new("v");
            v.push_text(value.trim());
            c.push(v);
        }
        Some("boolean") | Some("bool") => {
            let b = parse_bool(value)?;
            c.set_attr("t", "b");
            let mut v = XmlElement::new("v");
            v.push_text(if b { "1" } else { "0" });
            c.push(v);
        }
        Some(other) => bail!("unknown cell type '{other}' (string/number/boolean/formula)"),
        None => {
            if as_number {
                let mut v = XmlElement::new("v");
                v.push_text(value.trim());
                c.push(v);
            } else {
                write_inline_string(c, value);
            }
        }
    }
    Ok(())
}

fn write_inline_string(c: &mut XmlElement, value: &str) {
    c.set_attr("t", "inlineStr");
    let mut is = XmlElement::new("is");
    let mut t = XmlElement::new("t");
    t.set_attr("xml:space", "preserve");
    t.push_text(value);
    is.push(t);
    c.push(is);
}

fn write_cell_formula(c: &mut XmlElement, formula: &str) {
    clear_cell_content(c);
    let mut f = XmlElement::new("f");
    f.push_text(formula.strip_prefix('=').unwrap_or(formula));
    c.push(f);
}

// ------------------------------------------------------------- styles ----

#[derive(Default)]
struct StyleRequest {
    bold: Option<bool>,
    italic: Option<bool>,
    color: Option<String>,
    size: Option<f64>,
    fill: Option<String>,
    font: Option<String>,
}

impl StyleRequest {
    fn from_props(props: &Props) -> Result<StyleRequest> {
        let get = |k: &str| props.get(k).or_else(|| props.get(&format!("font.{k}")));
        Ok(StyleRequest {
            bold: get("bold").map(parse_bool).transpose()?,
            italic: get("italic").map(parse_bool).transpose()?,
            color: get("color").map(parse_color).transpose()?,
            size: get("size").map(parse_pt).transpose()?,
            fill: props
                .get("fill")
                .or_else(|| props.get("background"))
                .map(parse_color)
                .transpose()?,
            font: get("font").map(|s| s.to_string()),
        })
    }

    fn any(&self) -> bool {
        self.bold.is_some()
            || self.italic.is_some()
            || self.color.is_some()
            || self.size.is_some()
            || self.fill.is_some()
            || self.font.is_some()
    }
}

/// Derive a new cellXfs entry from the cell's current one with the requested
/// font/fill changes applied. Returns the new xf index.
fn ensure_style(styles: &mut XmlElement, current_xf: usize, req: &StyleRequest) -> Result<usize> {
    // --- font ---
    let current_font_id: usize = styles
        .child("cellXfs")
        .and_then(|xfs| xfs.children_named("xf").get(current_xf).copied())
        .and_then(|xf| xf.attr_local("fontId"))
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    let mut font = styles
        .child("fonts")
        .and_then(|fonts| fonts.children_named("font").get(current_font_id).copied())
        .cloned()
        .unwrap_or_else(|| {
            let mut f = XmlElement::new("font");
            f.push(el("sz", &[("val", "11")]));
            f.push(el("name", &[("val", "Calibri")]));
            f
        });
    let toggle = |font: &mut XmlElement, name: &str, on: bool| {
        if on {
            font.ensure_child(name, name, true);
        } else {
            font.children
                .retain(|n| !matches!(n, XmlNode::Element(e) if e.local_name() == name));
        }
    };
    if let Some(b) = req.bold {
        toggle(&mut font, "b", b);
    }
    if let Some(i) = req.italic {
        toggle(&mut font, "i", i);
    }
    if let Some(color) = &req.color {
        let c = font.ensure_child("color", "color", false);
        c.remove_attr("theme");
        c.set_attr("rgb", &format!("FF{color}"));
    }
    if let Some(size) = req.size {
        font.ensure_child("sz", "sz", true)
            .set_attr("val", &format_num(size));
    }
    if let Some(name) = &req.font {
        font.ensure_child("name", "name", false).set_attr("val", name);
    }
    let font_id = find_or_push(styles.ensure_child("fonts", "fonts", false), "font", font);

    // --- fill ---
    let fill_id = if let Some(fill) = &req.fill {
        let mut f = XmlElement::new("fill");
        let mut pat = el("patternFill", &[("patternType", "solid")]);
        pat.push(el("fgColor", &[("rgb", format!("FF{fill}").as_str())]));
        pat.push(el("bgColor", &[("indexed", "64")]));
        f.push(pat);
        Some(find_or_push(
            styles.ensure_child("fills", "fills", false),
            "fill",
            f,
        ))
    } else {
        None
    };

    // --- xf ---
    let mut xf = styles
        .child("cellXfs")
        .and_then(|xfs| xfs.children_named("xf").get(current_xf).copied())
        .cloned()
        .unwrap_or_else(|| {
            el(
                "xf",
                &[("numFmtId", "0"), ("fontId", "0"), ("fillId", "0"), ("borderId", "0"), ("xfId", "0")],
            )
        });
    xf.set_attr("fontId", &font_id.to_string());
    xf.set_attr("applyFont", "1");
    if let Some(fid) = fill_id {
        xf.set_attr("fillId", &fid.to_string());
        xf.set_attr("applyFill", "1");
    }
    let xf_id = find_or_push(styles.ensure_child("cellXfs", "cellXfs", false), "xf", xf);
    Ok(xf_id)
}

/// Find an existing identical element under `parent` or append it; returns
/// the element's index among `child_name` children. Keeps `count` accurate.
fn find_or_push(parent: &mut XmlElement, child_name: &str, candidate: XmlElement) -> usize {
    let existing = parent
        .children_named(child_name)
        .into_iter()
        .position(|e| *e == candidate);
    let idx = match existing {
        Some(i) => i,
        None => {
            parent.push(candidate);
            parent.children_named(child_name).len() - 1
        }
    };
    let count = parent.children_named(child_name).len();
    parent.set_attr("count", &count.to_string());
    idx
}

fn format_num(n: f64) -> String {
    if n.fract() == 0.0 {
        format!("{}", n as i64)
    } else {
        format!("{n}")
    }
}

// ------------------------------------------------------------ Handler ----

impl Handler for Xlsx {
    fn view(&mut self, mode: &str) -> Result<Report> {
        match mode {
            "text" => {
                let mut out = String::new();
                for sheet in &self.sheets {
                    out.push_str(&self.render_sheet_grid(sheet));
                    out.push('\n');
                }
                Ok(Report::Text(out))
            }
            "outline" => {
                let mut out = String::new();
                for (i, sheet) in self.sheets.iter().enumerate() {
                    let dims = Self::used_range(sheet)
                        .map(|((c0, r0), (c1, r1))| {
                            format!("{}:{}", cell_name(c0, r0), cell_name(c1, r1))
                        })
                        .unwrap_or_else(|| "empty".to_string());
                    out.push_str(&format!("Sheet {}: {} ({})\n", i + 1, sheet.name, dims));
                }
                Ok(Report::Text(out))
            }
            "stats" => {
                let mut cells = 0usize;
                let mut formulas = 0usize;
                for sheet in &self.sheets {
                    if let Some(sd) = sheet.xml.child("sheetData") {
                        for row in sd.children_named("row") {
                            for c in row.children_named("c") {
                                cells += 1;
                                if c.child("f").is_some() {
                                    formulas += 1;
                                }
                            }
                        }
                    }
                }
                Ok(Report::Data {
                    text: format!(
                        "sheets: {}\ncells: {cells}\nformulas: {formulas}",
                        self.sheets.len()
                    ),
                    data: json!({
                        "sheets": self.sheets.len(),
                        "cells": cells,
                        "formulas": formulas,
                    }),
                })
            }
            other => bail!("unknown view mode '{other}' for xlsx (text/outline/stats)"),
        }
    }

    fn get(&mut self, path_str: &str, depth: usize) -> Result<Report> {
        let dpath = path::parse(path_str)?;
        if dpath.is_root() {
            let mut info = NodeInfo::new("/", "workbook");
            info.attr("sheets", self.sheets.len().to_string());
            for (i, sheet) in self.sheets.iter().enumerate() {
                let mut child = NodeInfo::new(format!("/{}", sheet.name), "sheet");
                child.attr("index", (i + 1).to_string());
                if let Some(((c0, r0), (c1, r1))) = Self::used_range(sheet) {
                    child.attr("range", format!("{}:{}", cell_name(c0, r0), cell_name(c1, r1)));
                }
                info.children.push(child);
            }
            return Ok(Report::Nodes(vec![info]));
        }

        let sheet_idx = self.sheet_index(&dpath.segments[0])?;
        let sheet = &self.sheets[sheet_idx];

        if dpath.segments.len() == 1 {
            let mut info = NodeInfo::new(format!("/{}", sheet.name), "sheet");
            if let Some(((c0, r0), (c1, r1))) = Self::used_range(sheet) {
                info.attr("range", format!("{}:{}", cell_name(c0, r0), cell_name(c1, r1)));
            }
            if depth > 0 {
                if let Some(sd) = sheet.xml.child("sheetData") {
                    for row in sd.children_named("row") {
                        let rnum = row.attr_local("r").unwrap_or("?");
                        let mut row_info =
                            NodeInfo::new(format!("/{}/row[{}]", sheet.name, rnum), "row");
                        if depth > 1 {
                            for c in row.children_named("c") {
                                let r = c.attr_local("r").unwrap_or("?");
                                let cpath = format!("/{}/{}", sheet.name, r);
                                row_info.children.push(self.cell_info(sheet, c, &cpath));
                            }
                        }
                        info.children.push(row_info);
                    }
                }
            }
            return Ok(Report::Nodes(vec![info]));
        }

        let seg = &dpath.segments[1];
        if dpath.segments.len() == 2 {
            // row[N]
            if seg.name.eq_ignore_ascii_case("row") {
                let n = seg.index().context("row needs an index, e.g. row[5]")? as u32;
                let sd = sheet.xml.child("sheetData").context("no sheetData")?;
                let row = find_row(sd, n)
                    .with_context(|| format!("row {n} has no content in {}", sheet.name))?;
                let mut info = NodeInfo::new(format!("/{}/row[{}]", sheet.name, n), "row");
                for c in row.children_named("c") {
                    let r = c.attr_local("r").unwrap_or("?");
                    let cpath = format!("/{}/{}", sheet.name, r);
                    info.children.push(self.cell_info(sheet, c, &cpath));
                }
                return Ok(Report::Nodes(vec![info]));
            }
            // A1 or A1:C3
            if let Some((from, to)) = seg.name.split_once(':') {
                let (c0, r0) = parse_cell_ref(from)
                    .with_context(|| format!("'{from}' is not a cell reference"))?;
                let (c1, r1) = parse_cell_ref(to)
                    .with_context(|| format!("'{to}' is not a cell reference"))?;
                let mut nodes = Vec::new();
                let sd = sheet.xml.child("sheetData").context("no sheetData")?;
                for r in r0.min(r1)..=r0.max(r1) {
                    let Some(row) = find_row(sd, r + 1) else { continue };
                    for col in c0.min(c1)..=c0.max(c1) {
                        if let Some(c) = find_cell(row, col, r) {
                            let cpath = format!("/{}/{}", sheet.name, cell_name(col, r));
                            nodes.push(self.cell_info(sheet, c, &cpath));
                        }
                    }
                }
                return Ok(Report::Nodes(nodes));
            }
            if let Some((col, row0)) = parse_cell_ref(&seg.name) {
                let sd = sheet.xml.child("sheetData").context("no sheetData")?;
                let cpath = format!("/{}/{}", sheet.name, cell_name(col, row0));
                let info = match find_row(sd, row0 + 1).and_then(|r| find_cell(r, col, row0)) {
                    Some(c) => self.cell_info(sheet, c, &cpath),
                    None => {
                        let mut info = NodeInfo::new(cpath, "cell");
                        info.text = Some(String::new());
                        info.attr("type", "empty");
                        info
                    }
                };
                return Ok(Report::Nodes(vec![info]));
            }
        }
        bail!(
            "cannot resolve xlsx path '{path_str}' (expected /SheetName, /SheetName/A1, /SheetName/A1:C3, or /SheetName/row[N])"
        )
    }

    fn add(&mut self, parent: &str, typ: &str, props: &Props, pos: &Position) -> Result<Report> {
        let dpath = path::parse(parent)?;
        match typ.to_ascii_lowercase().as_str() {
            "sheet" => {
                if !dpath.is_root() {
                    bail!("sheets are added at the root: officecli add file.xlsx / --type sheet");
                }
                let default_name = format!("Sheet{}", self.sheets.len() + 1);
                let name = props.get("name").unwrap_or(&default_name).to_string();
                if self.sheets.iter().any(|s| s.name == name) {
                    bail!("a sheet named '{name}' already exists");
                }
                // New part number = max existing sheetN + 1.
                let next_part_num = self
                    .pkg
                    .part_names()
                    .filter_map(|p| {
                        p.strip_prefix("xl/worksheets/sheet")
                            .and_then(|s| s.strip_suffix(".xml"))
                            .and_then(|n| n.parse::<u32>().ok())
                    })
                    .max()
                    .unwrap_or(0)
                    + 1;
                let part = format!("xl/worksheets/sheet{next_part_num}.xml");
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
                let sheet_id = self
                    .workbook
                    .child("sheets")
                    .map(|s| {
                        s.children_named("sheet")
                            .into_iter()
                            .filter_map(|e| {
                                e.attr_local("sheetId").and_then(|v| v.parse::<u32>().ok())
                            })
                            .max()
                            .unwrap_or(0)
                    })
                    .unwrap_or(0)
                    + 1;

                self.rels.push(el(
                    "Relationship",
                    &[
                        ("Id", rid.as_str()),
                        (
                            "Type",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                        ),
                        ("Target", format!("worksheets/sheet{next_part_num}.xml").as_str()),
                    ],
                ));
                let sheets_el = self.workbook.ensure_child("sheets", "sheets", false);
                sheets_el.push(el(
                    "sheet",
                    &[
                        ("name", name.as_str()),
                        ("sheetId", sheet_id.to_string().as_str()),
                        ("r:id", rid.as_str()),
                    ],
                ));
                // Content type override for the new part.
                let mut ct = self.pkg.xml(CONTENT_TYPES_PART)?;
                ct.push(el(
                    "Override",
                    &[
                        ("PartName", format!("/{part}").as_str()),
                        (
                            "ContentType",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                        ),
                    ],
                ));
                self.pkg.put_xml(CONTENT_TYPES_PART, &ct)?;
                self.pkg.put_raw(&part, XLSX_BLANK_SHEET.as_bytes().to_vec());
                let xml = crate::xml::parse(XLSX_BLANK_SHEET.as_bytes())?;
                self.sheets.push(Sheet { name: name.clone(), part, xml });
                let mut info = NodeInfo::new(format!("/{name}"), "sheet");
                info.attr("index", self.sheets.len().to_string());
                Ok(Report::Nodes(vec![info]))
            }
            "row" => {
                if dpath.segments.is_empty() {
                    bail!("add row needs a sheet path, e.g. /Sheet1");
                }
                let sheet_idx = self.sheet_index(&dpath.segments[0])?;
                // Row number: --index is 1-based here (matches OOXML row index).
                let row_num: u32 = match pos {
                    Position::Index(n) => *n as u32,
                    Position::Append => {
                        let sheet = &self.sheets[sheet_idx];
                        Self::used_range(sheet)
                            .map(|((_, _), (_, r1))| r1 + 2)
                            .unwrap_or(1)
                    }
                    _ => bail!("add row supports --index N (1-based row number) or append"),
                };
                if row_num == 0 {
                    bail!("row numbers are 1-based");
                }
                // Shift existing rows at or below the insertion point down.
                {
                    let sheet = &mut self.sheets[sheet_idx];
                    let sd = sheet
                        .xml
                        .child_mut("sheetData")
                        .context("worksheet has no <sheetData>")?;
                    let mut needs_shift = false;
                    for row in sd.children.iter().filter_map(|n| n.as_element()) {
                        if row.local_name() == "row" {
                            if let Some(r) = row.attr_local("r").and_then(|v| v.parse::<u32>().ok()) {
                                if r >= row_num {
                                    needs_shift = true;
                                    break;
                                }
                            }
                        }
                    }
                    if needs_shift {
                        shift_rows_down(sd, row_num);
                    }
                }
                // Fill values: values="a,b,c" and/or cN=value props.
                let mut wrote_any = false;
                if let Some(values) = props.get("values") {
                    let values = values.to_string();
                    for (i, v) in values.split(',').enumerate() {
                        let value_props = Props::from_pairs(vec![(
                            "value".to_string(),
                            v.trim().to_string(),
                        )]);
                        self.set_cell(sheet_idx, i as u32, row_num - 1, &value_props)?;
                        wrote_any = true;
                    }
                }
                let cell_props: Vec<(u32, String)> = props
                    .iter()
                    .filter_map(|(k, v)| {
                        k.strip_prefix('c')
                            .and_then(|n| n.parse::<u32>().ok())
                            .filter(|n| *n >= 1)
                            .map(|n| (n - 1, v.to_string()))
                    })
                    .collect();
                for (col, v) in cell_props {
                    let value_props =
                        Props::from_pairs(vec![("value".to_string(), v)]);
                    self.set_cell(sheet_idx, col, row_num - 1, &value_props)?;
                    wrote_any = true;
                }
                if !wrote_any {
                    // Materialize an empty row element.
                    let sheet = &mut self.sheets[sheet_idx];
                    let sd = sheet.xml.child_mut("sheetData").unwrap();
                    ensure_row(sd, row_num);
                }
                let sheet = &self.sheets[sheet_idx];
                let mut info =
                    NodeInfo::new(format!("/{}/row[{}]", sheet.name, row_num), "row");
                info.attr("row", row_num.to_string());
                Ok(Report::Nodes(vec![info]))
            }
            other => bail!("unsupported xlsx element type '{other}' (sheet/row; cells are created via set)"),
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
            let replace = replace.context("xlsx find requires --replace (find+format is not supported)")?;
            let is_regex = props.get_bool("regex")?.unwrap_or(false);
            let pattern = if is_regex {
                find.to_string()
            } else {
                regex::escape(find)
            };
            let re = Regex::new(&pattern).with_context(|| format!("invalid regex '{find}'"))?;
            let scope: Vec<usize> = if dpath.is_root() {
                (0..self.sheets.len()).collect()
            } else {
                vec![self.sheet_index(&dpath.segments[0])?]
            };
            let mut matched = 0usize;
            for si in scope {
                let mut updates: Vec<(u32, u32, String)> = Vec::new();
                {
                    let sheet = &self.sheets[si];
                    if let Some(sd) = sheet.xml.child("sheetData") {
                        for row in sd.children_named("row") {
                            for c in row.children_named("c") {
                                let t = c.attr_local("t").unwrap_or("n");
                                if !matches!(t, "s" | "str" | "inlineStr") {
                                    continue;
                                }
                                let (value, _) = self.cell_display(c);
                                let hits = re.find_iter(&value).count();
                                if hits > 0 {
                                    matched += hits;
                                    let new_value = re.replace_all(&value, replace).into_owned();
                                    if let Some((col, r)) =
                                        c.attr_local("r").and_then(parse_cell_ref_opt)
                                    {
                                        updates.push((col, r, new_value));
                                    }
                                }
                            }
                        }
                    }
                }
                for (col, r, new_value) in updates {
                    let value_props = Props::from_pairs(vec![
                        ("value".to_string(), new_value),
                        ("type".to_string(), "string".to_string()),
                    ]);
                    self.set_cell(si, col, r, &value_props)?;
                }
            }
            return Ok(Report::Data {
                text: format!("matched: {matched}"),
                data: json!({ "matched": matched }),
            });
        }

        if dpath.is_root() {
            bail!("set on '/' needs --find/--replace for xlsx");
        }
        let sheet_idx = self.sheet_index(&dpath.segments[0])?;

        if dpath.segments.len() == 1 {
            // Sheet-level props: rename.
            if let Some(new_name) = props.get("name") {
                let new_name = new_name.to_string();
                if self.sheets.iter().any(|s| s.name == new_name) {
                    bail!("a sheet named '{new_name}' already exists");
                }
                let old = self.sheets[sheet_idx].name.clone();
                if let Some(sheets_el) = self.workbook.child_mut("sheets") {
                    for node in sheets_el.children.iter_mut() {
                        if let XmlNode::Element(e) = node {
                            if e.local_name() == "sheet" && e.attr_local("name") == Some(&old) {
                                e.set_attr("name", &new_name);
                            }
                        }
                    }
                }
                self.sheets[sheet_idx].name = new_name.clone();
                return Ok(Report::Data {
                    text: format!("renamed sheet '{old}' to '{new_name}'"),
                    data: json!({ "renamed": { "from": old, "to": new_name } }),
                });
            }
            bail!("sheet-level set supports --prop name=NewName");
        }

        let seg = &dpath.segments[1];
        let (col, row0) = parse_cell_ref(&seg.name)
            .with_context(|| format!("'{}' is not a cell reference like B2", seg.name))?;
        if !props.has("value") && !props.has("formula") && !StyleRequest::from_props(props)?.any() {
            bail!("set on a cell needs --prop value=... , formula=... , or style props (bold/italic/color/size/fill/font)");
        }
        let info = self.set_cell(sheet_idx, col, row0, props)?;
        Ok(Report::Nodes(vec![info]))
    }

    fn remove(&mut self, path_str: &str) -> Result<Report> {
        let dpath = path::parse(path_str)?;
        if dpath.is_root() {
            bail!("cannot remove the workbook root");
        }
        let sheet_idx = self.sheet_index(&dpath.segments[0])?;

        if dpath.segments.len() == 1 {
            if self.sheets.len() == 1 {
                bail!("cannot remove the last sheet in a workbook");
            }
            let sheet = self.sheets.remove(sheet_idx);
            // Remove workbook entry.
            if let Some(sheets_el) = self.workbook.child_mut("sheets") {
                sheets_el.children.retain(|n| {
                    !matches!(n, XmlNode::Element(e)
                        if e.local_name() == "sheet" && e.attr_local("name") == Some(&sheet.name))
                });
            }
            // Remove the relationship pointing at this part.
            let target_rel = sheet.part.strip_prefix("xl/").unwrap_or(&sheet.part).to_string();
            self.rels.children.retain(|n| {
                !matches!(n, XmlNode::Element(e)
                    if e.local_name() == "Relationship"
                        && e.attr_local("Target") == Some(&target_rel))
            });
            // Remove content-type override and the part itself.
            let mut ct = self.pkg.xml(CONTENT_TYPES_PART)?;
            let part_name = format!("/{}", sheet.part);
            ct.children.retain(|n| {
                !matches!(n, XmlNode::Element(e)
                    if e.local_name() == "Override"
                        && e.attr_local("PartName") == Some(&part_name))
            });
            self.pkg.put_xml(CONTENT_TYPES_PART, &ct)?;
            self.pkg.remove_part(&sheet.part);
            return Ok(Report::Data {
                text: format!("removed sheet '{}'", sheet.name),
                data: json!({ "removed": sheet.name }),
            });
        }

        let seg = &dpath.segments[1];
        let sheet = &mut self.sheets[sheet_idx];
        let sd = sheet
            .xml
            .child_mut("sheetData")
            .context("worksheet has no <sheetData>")?;
        if seg.name.eq_ignore_ascii_case("row") {
            let n = seg.index().context("row needs an index, e.g. row[5]")? as u32;
            let before = sd.children.len();
            sd.children.retain(|node| {
                !matches!(node, XmlNode::Element(e)
                    if e.local_name() == "row"
                        && e.attr_local("r").and_then(|v| v.parse::<u32>().ok()) == Some(n))
            });
            if sd.children.len() == before {
                bail!("row {n} has no content in {}", sheet.name);
            }
            shift_rows_up(sd, n);
            return Ok(Report::Data {
                text: format!("removed row {n} from {}", sheet.name),
                data: json!({ "removed": format!("/{}/row[{}]", sheet.name, n) }),
            });
        }
        if let Some((col, row0)) = parse_cell_ref(&seg.name) {
            let target = cell_name(col, row0);
            let mut removed = false;
            for node in sd.children.iter_mut() {
                let XmlNode::Element(row) = node else { continue };
                if row.local_name() != "row" {
                    continue;
                }
                let before = row.children.len();
                row.children.retain(|n| {
                    !matches!(n, XmlNode::Element(e)
                        if e.local_name() == "c" && e.attr_local("r") == Some(target.as_str()))
                });
                if row.children.len() != before {
                    removed = true;
                }
            }
            if !removed {
                bail!("cell {target} is already empty in {}", sheet.name);
            }
            return Ok(Report::Data {
                text: format!("cleared cell {} in {}", target, sheet.name),
                data: json!({ "removed": format!("/{}/{}", sheet.name, target) }),
            });
        }
        bail!("cannot resolve xlsx path '{path_str}' for remove");
    }

    fn validate(&mut self) -> Result<Report> {
        let mut problems = Vec::new();
        if self.sheets.is_empty() {
            problems.push("workbook has no sheets".to_string());
        }
        for part in [CONTENT_TYPES_PART, "_rels/.rels", STYLES_PART] {
            if !self.pkg.has_part(part) {
                problems.push(format!("missing package part {part}"));
            }
        }
        let ok = problems.is_empty();
        Ok(Report::Data {
            text: if ok { "valid".to_string() } else { problems.join("\n") },
            data: json!({ "valid": ok, "problems": problems }),
        })
    }

    fn save(&mut self, path: &Path) -> Result<()> {
        self.save_sheets()?;
        self.pkg.save(path)
    }
}

fn ensure_row(sheet_data: &mut XmlElement, row_num: u32) {
    let exists = sheet_data.children.iter().any(|n| {
        matches!(n, XmlNode::Element(e)
            if e.local_name() == "row"
                && e.attr_local("r").and_then(|v| v.parse::<u32>().ok()) == Some(row_num))
    });
    if exists {
        return;
    }
    let mut insert_at = sheet_data.children.len();
    for (i, node) in sheet_data.children.iter().enumerate() {
        if let XmlNode::Element(e) = node {
            if e.local_name() == "row" {
                let r: u32 = e
                    .attr_local("r")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(u32::MAX);
                if r > row_num {
                    insert_at = i;
                    break;
                }
            }
        }
    }
    let row = el("row", &[("r", row_num.to_string().as_str())]);
    sheet_data.children.insert(insert_at, XmlNode::Element(row));
}

/// Renumber rows >= `from` one down the sheet (insert). Cell refs follow.
/// Note: formula references are NOT rewritten (documented limitation).
fn shift_rows_down(sheet_data: &mut XmlElement, from: u32) {
    renumber_rows(sheet_data, from, 1);
}

/// Renumber rows > `removed` one up (delete).
fn shift_rows_up(sheet_data: &mut XmlElement, removed: u32) {
    renumber_rows(sheet_data, removed + 1, -1);
}

fn renumber_rows(sheet_data: &mut XmlElement, from: u32, delta: i64) {
    for node in sheet_data.children.iter_mut() {
        let XmlNode::Element(row) = node else { continue };
        if row.local_name() != "row" {
            continue;
        }
        let Some(r) = row.attr_local("r").and_then(|v| v.parse::<u32>().ok()) else {
            continue;
        };
        if r < from {
            continue;
        }
        let new_r = (r as i64 + delta) as u32;
        row.set_attr("r", &new_r.to_string());
        for cnode in row.children.iter_mut() {
            let XmlNode::Element(c) = cnode else { continue };
            if c.local_name() != "c" {
                continue;
            }
            if let Some((col, _)) = c.attr_local("r").and_then(parse_cell_ref_opt) {
                c.set_attr("r", &cell_name(col, new_r - 1));
            }
        }
    }
}
