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
            _ => {
                // Numbers styled with a date format read back as ISO.
                if let Some(kind) = self.cell_date_kind(c) {
                    if let Ok(serial) = v.trim().parse::<f64>() {
                        let (wd, wt, label) = match kind {
                            DateKind::Date => (true, false, "date"),
                            DateKind::DateTime => (true, true, "datetime"),
                            DateKind::Time => (false, true, "time"),
                        };
                        return (serial_to_iso(serial, wd, wt), label);
                    }
                }
                (v, "number")
            }
        }
    }

    /// Date-like number format applied to this cell, if any.
    fn cell_date_kind(&self, c: &XmlElement) -> Option<DateKind> {
        let s: usize = c.attr_local("s")?.parse().ok()?;
        let xf = self
            .styles
            .child("cellXfs")?
            .children_named("xf")
            .get(s)
            .copied()?;
        let id: u32 = xf.attr_local("numFmtId")?.parse().ok()?;
        if let Some(kind) = builtin_date_kind(id) {
            return Some(kind);
        }
        if id >= 164 {
            let code = self
                .styles
                .child("numFmts")?
                .children_named("numFmt")
                .into_iter()
                .find(|f| {
                    f.attr_local("numFmtId").and_then(|v| v.parse::<u32>().ok()) == Some(id)
                })
                .and_then(|f| f.attr_local("formatCode"))?
                .to_string();
            return code_date_kind(&code);
        }
        None
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
    /// (from_col, from_row, to_col, to_row, chartSpace) for every chart
    /// anchored on the sheet's drawing part — 0-based cell coordinates,
    /// used by the screenshot renderer.
    fn sheet_chart_anchors(&self, sheet: &Sheet) -> Vec<(u32, u32, u32, u32, XmlElement)> {
        let mut out = Vec::new();
        let rels_part = sheet.part.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels";
        if !self.pkg.has_part(&rels_part) {
            return out;
        }
        let Ok(rels) = self.pkg.xml(&rels_part) else { return out };
        let Some(rid) = sheet
            .xml
            .child("drawing")
            .and_then(|d| d.attr("r:id").or_else(|| d.attr_local("id")))
        else {
            return out;
        };
        let Some(target) = rel_target(&rels, rid) else { return out };
        let drawing_part = resolve_target("xl/worksheets", &target);
        let Ok(drawing) = self.pkg.xml(&drawing_part) else { return out };
        let drawing_rels_part = drawing_part.replace("xl/drawings/", "xl/drawings/_rels/") + ".rels";
        let Ok(drels) = self.pkg.xml(&drawing_rels_part) else { return out };
        for anchor in drawing.children_named("twoCellAnchor") {
            let cell = |tag: &str| -> Option<(u32, u32)> {
                let e = anchor.child(tag)?;
                Some((
                    e.child("col")?.text_content().trim().parse().ok()?,
                    e.child("row")?.text_content().trim().parse().ok()?,
                ))
            };
            let (Some((fc, fr)), Some((tc, tr))) = (cell("from"), cell("to")) else {
                continue;
            };
            let Some(chart_rid) = anchor
                .child("graphicFrame")
                .and_then(|f| f.child("graphic"))
                .and_then(|g| g.child("graphicData"))
                .and_then(|d| d.child("chart"))
                .and_then(|c| c.attr_local("id"))
            else {
                continue;
            };
            let Some(t) = rel_target(&drels, chart_rid) else { continue };
            if let Ok(space) = self.pkg.xml(&resolve_target("xl/drawings", &t)) {
                out.push((fc, fr, tc, tr, space));
            }
        }
        out
    }

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
        let mut style_props = StyleRequest::from_props(props)?;

        // ISO dates/times become serial numbers with a date number format
        // (opt out with type=string, force with type=date).
        let mut value_override: Option<String> = None;
        if let Some(value) = props.get("value") {
            let hint = props.get("type").map(|t| t.to_ascii_lowercase());
            let hint = hint.as_deref();
            let date_requested = matches!(hint, Some("date") | Some("datetime") | Some("time"));
            if (hint.is_none() || date_requested) && !value.starts_with('=') {
                if let Some((serial, has_date, has_time)) = parse_iso_datetime(value) {
                    value_override = Some(format!("{serial}"));
                    if style_props.num_fmt.is_none() {
                        style_props.num_fmt = Some(
                            match (has_date, has_time) {
                                (true, true) => "datetime",
                                (false, true) => "time",
                                _ => "date",
                            }
                            .to_string(),
                        );
                    }
                } else if date_requested {
                    bail!(
                        "'{value}' is not a date/time (YYYY-MM-DD, YYYY-MM-DD HH:MM[:SS], or HH:MM:SS)"
                    );
                }
            }
        }

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

            if let Some(serial) = &value_override {
                write_cell_value(c, serial, Some("number"))?;
            } else if let Some(value) = props.get("value") {
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

    /// After rows or columns of one sheet shift, update A1-style references
    /// in every formula of the workbook.
    fn rewrite_all_formulas(&mut self, target_sheet_idx: usize, shift: &Shift) {
        let target_name = self.sheets[target_sheet_idx].name.clone();
        for (i, sheet) in self.sheets.iter_mut().enumerate() {
            let own = i == target_sheet_idx;
            let Some(sd) = sheet.xml.child_mut("sheetData") else { continue };
            for row in sd.children.iter_mut() {
                let XmlNode::Element(row) = row else { continue };
                for cnode in row.children.iter_mut() {
                    let XmlNode::Element(c) = cnode else { continue };
                    if c.local_name() != "c" {
                        continue;
                    }
                    let Some(f) = c.child_mut("f") else { continue };
                    let old = f.text_content();
                    let new = rewrite_formula_refs(&old, &target_name, own, shift);
                    if new != old {
                        f.children.clear();
                        f.push_text(&new);
                    }
                }
            }
        }
    }

    /// Create a blank sheet named `name`, wiring workbook, rels, and content
    /// types. Returns the new sheet index.
    fn create_sheet(&mut self, name: &str) -> Result<usize> {
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
        let rid = next_rid(&self.rels);
        let sheet_id = self
            .workbook
            .child("sheets")
            .map(|s| {
                s.children_named("sheet")
                    .into_iter()
                    .filter_map(|e| e.attr_local("sheetId").and_then(|v| v.parse::<u32>().ok()))
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
                ("name", name),
                ("sheetId", sheet_id.to_string().as_str()),
                ("r:id", rid.as_str()),
            ],
        ));
        self.pkg.add_override(
            &part,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        )?;
        self.pkg.put_raw(&part, XLSX_BLANK_SHEET.as_bytes().to_vec());
        let xml = crate::xml::parse(XLSX_BLANK_SHEET.as_bytes())?;
        self.sheets.push(Sheet {
            name: name.to_string(),
            part,
            xml,
        });
        Ok(self.sheets.len() - 1)
    }

    /// Parse `A1:C10` or `Sheet2!A1:C10` into (sheet index, corners).
    fn parse_range_spec(
        &self,
        default_sheet: usize,
        spec: &str,
    ) -> Result<(usize, (u32, u32), (u32, u32))> {
        let (sheet_idx, range) = match spec.split_once('!') {
            Some((sheet, range)) => {
                let name = sheet.trim_matches('\'');
                let idx = self
                    .sheets
                    .iter()
                    .position(|s| s.name == name)
                    .with_context(|| format!("no sheet named '{name}'"))?;
                (idx, range)
            }
            None => (default_sheet, spec),
        };
        let (from, to) = match range.split_once(':') {
            Some((a, b)) => (a, b),
            None => (range, range),
        };
        let strip = |s: &str| s.replace('$', "");
        let (c0, r0) = parse_cell_ref(&strip(from))
            .with_context(|| format!("'{from}' is not a cell reference"))?;
        let (c1, r1) = parse_cell_ref(&strip(to))
            .with_context(|| format!("'{to}' is not a cell reference"))?;
        Ok((
            sheet_idx,
            (c0.min(c1), r0.min(r1)),
            (c0.max(c1), r0.max(r1)),
        ))
    }

    /// Display value at (col, row0), or empty when the cell doesn't exist.
    fn value_at(&self, sheet_idx: usize, col: u32, row0: u32) -> String {
        self.sheets[sheet_idx]
            .xml
            .child("sheetData")
            .and_then(|sd| find_row(sd, row0 + 1))
            .and_then(|r| find_cell(r, col, row0))
            .map(|c| self.cell_display(c).0)
            .unwrap_or_default()
    }

    /// Absolute reference like `Sheet1!$A$1:$A$5` (single cell when equal).
    fn abs_ref(&self, sheet_idx: usize, c0: u32, r0: u32, c1: u32, r1: u32) -> String {
        let name = &self.sheets[sheet_idx].name;
        let quoted = if name.chars().all(|c| c.is_ascii_alphanumeric() || c == '_') {
            name.clone()
        } else {
            format!("'{}'", name.replace('\'', "''"))
        };
        let a = format!("${}${}", col_letters(c0), r0 + 1);
        if (c0, r0) == (c1, r1) {
            format!("{quoted}!{a}")
        } else {
            format!("{quoted}!{a}:${}${}", col_letters(c1), r1 + 1)
        }
    }

    /// Insert a chart fed by a cell range, anchored on the sheet.
    fn add_chart(&mut self, sheet_idx: usize, props: &Props) -> Result<Report> {
        let data = props
            .get("data")
            .or_else(|| props.get("source"))
            .context("chart needs --prop data=A1:B5 (first column = categories, first row = headers)")?;
        let (src_idx, (c0, r0), (c1, r1)) = self.parse_range_spec(sheet_idx, data)?;
        let kind = crate::chart::parse_kind(props.get("kind").unwrap_or("column"))?;

        // Header row detection: any non-numeric first cell in the value
        // columns, unless --prop headers= forces it.
        let has_headers = match props.get_bool("headers")? {
            Some(b) => b,
            None => {
                let mut non_numeric = false;
                for col in (c0 + 1).max(c0)..=c1 {
                    let v = self.value_at(src_idx, col, r0);
                    if !v.is_empty() && v.parse::<f64>().is_err() {
                        non_numeric = true;
                    }
                }
                non_numeric && r1 > r0
            }
        };
        let data_r0 = if has_headers { r0 + 1 } else { r0 };
        if data_r0 > r1 {
            bail!("chart range {data} has no data rows");
        }

        // Single-column ranges are a bare value series; wider ranges use the
        // first column as categories (scatter: as numeric x values).
        let (cat_col, first_val_col) = if c0 == c1 { (None, c0) } else { (Some(c0), c0 + 1) };
        let scatter = kind == crate::chart::ChartKind::Scatter;
        let cats: Vec<String> = match cat_col {
            Some(cc) if !scatter => {
                (data_r0..=r1).map(|r| self.value_at(src_idx, cc, r)).collect()
            }
            _ => Vec::new(),
        };
        let xs: Vec<f64> = match cat_col {
            Some(cc) if scatter => (data_r0..=r1)
                .enumerate()
                .map(|(i, r)| {
                    self.value_at(src_idx, cc, r)
                        .parse::<f64>()
                        .unwrap_or((i + 1) as f64)
                })
                .collect(),
            _ => Vec::new(),
        };
        let cats_ref = cat_col.map(|cc| self.abs_ref(src_idx, cc, data_r0, cc, r1));

        let mut series = Vec::new();
        for (i, col) in (first_val_col..=c1).enumerate() {
            let name = if has_headers {
                let h = self.value_at(src_idx, col, r0);
                if h.is_empty() { format!("Series {}", i + 1) } else { h }
            } else {
                format!("Series {}", i + 1)
            };
            let vals: Vec<f64> = (data_r0..=r1)
                .map(|r| self.value_at(src_idx, col, r).parse::<f64>().unwrap_or(0.0))
                .collect();
            series.push(crate::chart::Series {
                name,
                name_ref: has_headers.then(|| self.abs_ref(src_idx, col, r0, col, r0)),
                cats: cats.clone(),
                cats_ref: if scatter { None } else { cats_ref.clone() },
                vals,
                vals_ref: Some(self.abs_ref(src_idx, col, data_r0, col, r1)),
                xs: xs.clone(),
                xs_ref: if scatter { cats_ref.clone() } else { None },
            });
        }
        let chart_space =
            crate::chart::build_chart_space(kind, props.get("title"), &series)?;

        // Chart part.
        let chart_num = self
            .pkg
            .part_names()
            .filter_map(|p| {
                p.strip_prefix("xl/charts/chart")
                    .and_then(|s| s.strip_suffix(".xml"))
                    .and_then(|n| n.parse::<u32>().ok())
            })
            .max()
            .unwrap_or(0)
            + 1;
        let chart_part = format!("xl/charts/chart{chart_num}.xml");
        self.pkg.put_xml(&chart_part, &chart_space)?;
        self.pkg
            .add_override(&chart_part, crate::chart::CHART_CONTENT_TYPE)?;

        // Drawing part: reuse the sheet's existing drawing or create one.
        let sheet_part = self.sheets[sheet_idx].part.clone();
        let sheet_rels_part = sheet_part.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels";
        let mut sheet_rels = if self.pkg.has_part(&sheet_rels_part) {
            self.pkg.xml(&sheet_rels_part)?
        } else {
            crate::xml::parse(EMPTY_RELS_XML.as_bytes())?
        };
        let existing_drawing = self.sheets[sheet_idx]
            .xml
            .child("drawing")
            .and_then(|d| d.attr("r:id").or_else(|| d.attr_local("id")))
            .and_then(|rid| rel_target(&sheet_rels, rid))
            .map(|t| resolve_target("xl/worksheets", &t));
        let (drawing_part, mut drawing) = match existing_drawing {
            Some(part) => {
                let xml = self.pkg.xml(&part)?;
                (part, xml)
            }
            None => {
                let n = self
                    .pkg
                    .part_names()
                    .filter_map(|p| {
                        p.strip_prefix("xl/drawings/drawing")
                            .and_then(|s| s.strip_suffix(".xml"))
                            .and_then(|n| n.parse::<u32>().ok())
                    })
                    .max()
                    .unwrap_or(0)
                    + 1;
                let part = format!("xl/drawings/drawing{n}.xml");
                let root = el(
                    "xdr:wsDr",
                    &[
                        (
                            "xmlns:xdr",
                            "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                        ),
                        ("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
                    ],
                );
                let rid = crate::media::add_relationship(
                    &mut sheet_rels,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                    &format!("../drawings/drawing{n}.xml"),
                );
                self.pkg
                    .add_override(&part, "application/vnd.openxmlformats-officedocument.drawing+xml")?;
                let sheet_xml = &mut self.sheets[sheet_idx].xml;
                if sheet_xml.attr("xmlns:r").is_none() {
                    sheet_xml.set_attr(
                        "xmlns:r",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    );
                }
                sheet_xml.push(el("drawing", &[("r:id", rid.as_str())]));
                (part, root)
            }
        };

        // Drawing → chart relationship.
        let drawing_rels_part =
            drawing_part.replace("xl/drawings/", "xl/drawings/_rels/") + ".rels";
        let mut drawing_rels = if self.pkg.has_part(&drawing_rels_part) {
            self.pkg.xml(&drawing_rels_part)?
        } else {
            crate::xml::parse(EMPTY_RELS_XML.as_bytes())?
        };
        let chart_rid = crate::media::add_relationship(
            &mut drawing_rels,
            crate::chart::CHART_REL_TYPE,
            &format!("../charts/chart{chart_num}.xml"),
        );

        // Anchor: --prop at=E2 top-left cell, spanning w x h cells.
        let (at_col, at_row) = props
            .get("at")
            .map(|a| parse_cell_ref(a).with_context(|| format!("'{a}' is not a cell reference")))
            .transpose()?
            .unwrap_or((c1 + 2, r0));
        let w_cells: u32 = props.get("w").map(|v| v.parse()).transpose().context("w must be a number of columns")?.unwrap_or(8);
        let h_cells: u32 = props.get("h").map(|v| v.parse()).transpose().context("h must be a number of rows")?.unwrap_or(15);

        let anchor_id = drawing.children_named("twoCellAnchor").len() as u64 + 2;
        let mut anchor = XmlElement::new("xdr:twoCellAnchor");
        for (tag, col, row) in [
            ("xdr:from", at_col, at_row),
            ("xdr:to", at_col + w_cells, at_row + h_cells),
        ] {
            let mut corner = XmlElement::new(tag);
            for (n, v) in [
                ("xdr:col", col.to_string()),
                ("xdr:colOff", "0".into()),
                ("xdr:row", row.to_string()),
                ("xdr:rowOff", "0".into()),
            ] {
                let mut e = XmlElement::new(n);
                e.push_text(&v);
                corner.push(e);
            }
            anchor.push(corner);
        }
        let mut frame = el("xdr:graphicFrame", &[("macro", "")]);
        let mut nv = XmlElement::new("xdr:nvGraphicFramePr");
        nv.push(el(
            "xdr:cNvPr",
            &[
                ("id", anchor_id.to_string().as_str()),
                ("name", format!("Chart {chart_num}").as_str()),
            ],
        ));
        nv.push(XmlElement::new("xdr:cNvGraphicFramePr"));
        frame.push(nv);
        let mut xfrm = XmlElement::new("xdr:xfrm");
        xfrm.push(el("a:off", &[("x", "0"), ("y", "0")]));
        xfrm.push(el("a:ext", &[("cx", "0"), ("cy", "0")]));
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
                ("r:id", chart_rid.as_str()),
            ],
        ));
        graphic.push(gdata);
        frame.push(graphic);
        anchor.push(frame);
        anchor.push(XmlElement::new("xdr:clientData"));
        drawing.push(anchor);

        self.pkg.put_xml(&drawing_part, &drawing)?;
        self.pkg.put_xml(&drawing_rels_part, &drawing_rels)?;
        self.pkg.put_xml(&sheet_rels_part, &sheet_rels)?;

        let mut info = NodeInfo::new(
            format!("/{}/chart[{}]", self.sheets[sheet_idx].name, chart_num),
            "chart",
        );
        info.attr("kind", props.get("kind").unwrap_or("column"));
        info.attr("data", data);
        info.attr("series", series.len().to_string());
        Ok(Report::Nodes(vec![info]))
    }

    /// Computed pivot: group-by aggregation written to a new sheet. This is
    /// a static summary table, not a native interactive PivotTable.
    fn add_pivot(&mut self, default_sheet: usize, props: &Props) -> Result<Report> {
        let source = props
            .get("source")
            .or_else(|| props.get("data"))
            .context("pivot needs --prop source=A1:C10 (first row = headers)")?;
        let (src_idx, (c0, r0), (c1, r1)) = self.parse_range_spec(default_sheet, source)?;
        if r1 <= r0 {
            bail!("pivot source {source} needs a header row plus data rows");
        }
        let agg = props.get("agg").unwrap_or("sum").to_ascii_lowercase();
        if !matches!(agg.as_str(), "sum" | "count" | "avg" | "average" | "min" | "max") {
            bail!("unknown aggregation '{agg}' (sum/count/avg/min/max)");
        }
        let headers: Vec<String> = (c0..=c1).map(|c| self.value_at(src_idx, c, r0)).collect();
        let find_col = |key: &str| -> Result<u32> {
            if let Some(i) = headers.iter().position(|h| h.eq_ignore_ascii_case(key)) {
                return Ok(c0 + i as u32);
            }
            if let Some(c) = letters_to_col(key) {
                if c >= c0 && c <= c1 {
                    return Ok(c);
                }
            }
            bail!("no column '{key}' in {source} (headers: {})", headers.join(", "))
        };
        let rows_key = props.get("rows").context("pivot needs --prop rows=HeaderName")?;
        let rows_col = find_col(rows_key)?;
        let values_col = match props.get("values") {
            Some(v) => Some(find_col(v)?),
            None if agg == "count" => None,
            None => bail!("pivot needs --prop values=HeaderName (or agg=count)"),
        };

        // Group in first-seen order.
        let mut order: Vec<String> = Vec::new();
        let mut groups: std::collections::HashMap<String, Vec<f64>> = Default::default();
        for r in (r0 + 1)..=r1 {
            let key = self.value_at(src_idx, rows_col, r);
            if key.is_empty() {
                continue;
            }
            let entry = match groups.get_mut(&key) {
                Some(e) => e,
                None => {
                    order.push(key.clone());
                    groups.entry(key.clone()).or_default()
                }
            };
            match values_col {
                Some(vc) => {
                    if let Ok(v) = self.value_at(src_idx, vc, r).parse::<f64>() {
                        entry.push(v);
                    }
                }
                None => entry.push(1.0),
            }
        }
        let aggregate = |vals: &[f64]| -> f64 {
            match agg.as_str() {
                "count" => vals.len() as f64,
                "avg" | "average" => {
                    if vals.is_empty() { 0.0 } else { vals.iter().sum::<f64>() / vals.len() as f64 }
                }
                "min" => vals.iter().cloned().fold(f64::INFINITY, f64::min),
                "max" => vals.iter().cloned().fold(f64::NEG_INFINITY, f64::max),
                _ => vals.iter().sum(),
            }
        };

        // Target sheet.
        let base = props.get("name").unwrap_or("Pivot").to_string();
        let mut name = base.clone();
        let mut n = 1;
        while self.sheets.iter().any(|s| s.name == name) {
            n += 1;
            name = format!("{base}{n}");
        }
        let pivot_idx = self.create_sheet(&name)?;

        let value_label = match (values_col, agg.as_str()) {
            (_, "count") => "Count".to_string(),
            (Some(vc), a) => {
                let header = &headers[(vc - c0) as usize];
                let a = format!("{}{}", a[..1].to_ascii_uppercase(), &a[1..]);
                format!("{a} of {header}")
            }
            _ => "Value".to_string(),
        };
        let set = |me: &mut Self, col: u32, row: u32, value: String, bold: bool| -> Result<()> {
            let mut pairs = vec![("value".to_string(), value)];
            if bold {
                pairs.push(("bold".to_string(), "true".to_string()));
            }
            me.set_cell(pivot_idx, col, row, &Props::from_pairs(pairs))?;
            Ok(())
        };
        set(self, 0, 0, headers[(rows_col - c0) as usize].clone(), true)?;
        set(self, 1, 0, value_label, true)?;
        let mut grand: Vec<f64> = Vec::new();
        for (i, key) in order.iter().enumerate() {
            let vals = &groups[key];
            grand.extend_from_slice(vals);
            set(self, 0, 1 + i as u32, key.clone(), false)?;
            set(self, 1, 1 + i as u32, format_num(aggregate(vals)), false)?;
        }
        let total_row = 1 + order.len() as u32;
        set(self, 0, total_row, "Grand Total".to_string(), true)?;
        set(self, 1, total_row, format_num(aggregate(&grand)), true)?;

        let mut info = NodeInfo::new(format!("/{name}"), "sheet");
        info.attr("kind", "pivot");
        info.attr("groups", order.len().to_string());
        info.attr("agg", agg);
        Ok(Report::Nodes(vec![info]))
    }

    /// Attach a hyperlink to a cell: worksheet <hyperlinks> entry, external
    /// relationship in the sheet's rels, and link styling.
    fn set_cell_hyperlink(&mut self, sheet_idx: usize, col: u32, row0: u32, url: &str) -> Result<()> {
        // Cells need visible text; default to the URL itself.
        if self.value_at(sheet_idx, col, row0).is_empty() {
            let value_props = Props::from_pairs(vec![
                ("value".to_string(), url.to_string()),
                ("type".to_string(), "string".to_string()),
            ]);
            self.set_cell(sheet_idx, col, row0, &value_props)?;
        }
        let style_props = Props::from_pairs(vec![
            ("color".to_string(), "0563C1".to_string()),
            ("underline".to_string(), "true".to_string()),
        ]);
        self.set_cell(sheet_idx, col, row0, &style_props)?;

        let sheet_part = self.sheets[sheet_idx].part.clone();
        let rels_part = sheet_part.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels";
        let mut rels = if self.pkg.has_part(&rels_part) {
            self.pkg.xml(&rels_part)?
        } else {
            crate::xml::parse(EMPTY_RELS_XML.as_bytes())?
        };
        let rid = crate::media::add_external_relationship(
            &mut rels,
            crate::media::HYPERLINK_REL_TYPE,
            url,
        );
        self.pkg.put_xml(&rels_part, &rels)?;

        let target = cell_name(col, row0);
        let sheet_xml = &mut self.sheets[sheet_idx].xml;
        if sheet_xml.attr("xmlns:r").is_none() {
            sheet_xml.set_attr(
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            );
        }
        let links = worksheet_child(sheet_xml, "hyperlinks");
        links.children.retain(|n| {
            !matches!(n, XmlNode::Element(e)
                if e.local_name() == "hyperlink" && e.attr_local("ref") == Some(target.as_str()))
        });
        links.push(el(
            "hyperlink",
            &[("ref", target.as_str()), ("r:id", rid.as_str())],
        ));
        Ok(())
    }

    /// Set (or replace) a cell comment. Writes the comments part plus the
    /// legacy VML note shape Excel needs to actually display it.
    fn set_cell_comment(
        &mut self,
        sheet_idx: usize,
        col: u32,
        row0: u32,
        text: &str,
        author: &str,
    ) -> Result<()> {
        let sheet_part = self.sheets[sheet_idx].part.clone();
        let sheet_rels_part = sheet_part.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels";
        let mut sheet_rels = if self.pkg.has_part(&sheet_rels_part) {
            self.pkg.xml(&sheet_rels_part)?
        } else {
            crate::xml::parse(EMPTY_RELS_XML.as_bytes())?
        };

        // Comments part: reuse the sheet's existing one or allocate the next.
        let existing = sheet_rels
            .children_named("Relationship")
            .into_iter()
            .find(|r| r.attr_local("Type") == Some(COMMENTS_REL_TYPE))
            .and_then(|r| r.attr_local("Target"))
            .map(|t| resolve_target("xl/worksheets", t));
        let (comments_part, mut comments) = match existing {
            Some(p) => {
                let xml = self.pkg.xml(&p)?;
                (p, xml)
            }
            None => {
                let n = self
                    .pkg
                    .part_names()
                    .filter_map(|p| {
                        p.strip_prefix("xl/comments")
                            .and_then(|s| s.strip_suffix(".xml"))
                            .and_then(|n| n.parse::<u32>().ok())
                    })
                    .max()
                    .unwrap_or(0)
                    + 1;
                let part = format!("xl/comments{n}.xml");
                crate::media::add_relationship(
                    &mut sheet_rels,
                    COMMENTS_REL_TYPE,
                    &format!("../comments{n}.xml"),
                );
                self.pkg.add_override(&part, COMMENTS_CONTENT_TYPE)?;
                let mut root = el(
                    "comments",
                    &[("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")],
                );
                root.push(XmlElement::new("authors"));
                root.push(XmlElement::new("commentList"));
                (part, root)
            }
        };

        let authors = comments.ensure_child("authors", "authors", false);
        let author_id = match authors
            .children_named("author")
            .iter()
            .position(|a| a.text_content() == author)
        {
            Some(i) => i,
            None => {
                let mut a = XmlElement::new("author");
                a.push_text(author);
                authors.push(a);
                authors.children_named("author").len() - 1
            }
        };

        let ref_name = cell_name(col, row0);
        let list = comments.ensure_child("commentList", "commentList", false);
        list.children.retain(|n| {
            !matches!(n, XmlNode::Element(e)
                if e.local_name() == "comment" && e.attr_local("ref") == Some(ref_name.as_str()))
        });
        let mut cm = el(
            "comment",
            &[("ref", ref_name.as_str()), ("authorId", author_id.to_string().as_str())],
        );
        let mut text_el = XmlElement::new("text");
        let mut run = XmlElement::new("r");
        let mut t = el("t", &[("xml:space", "preserve")]);
        t.push_text(text);
        run.push(t);
        text_el.push(run);
        cm.push(text_el);
        list.push(cm);
        self.pkg.put_xml(&comments_part, &comments)?;

        // VML note shape (hidden until hover), one per commented cell.
        let vml_existing = sheet_rels
            .children_named("Relationship")
            .into_iter()
            .find(|r| r.attr_local("Type") == Some(VML_REL_TYPE))
            .and_then(|r| Some((r.attr_local("Id")?.to_string(), r.attr_local("Target")?.to_string())));
        let (vml_part, vml_rid, mut vml) = match vml_existing {
            Some((rid, target)) => {
                let part = resolve_target("xl/worksheets", &target);
                let xml = self.pkg.xml(&part)?;
                (part, rid, xml)
            }
            None => {
                let n = self
                    .pkg
                    .part_names()
                    .filter_map(|p| {
                        p.strip_prefix("xl/drawings/vmlDrawing")
                            .and_then(|s| s.strip_suffix(".vml"))
                            .and_then(|n| n.parse::<u32>().ok())
                    })
                    .max()
                    .unwrap_or(0)
                    + 1;
                let part = format!("xl/drawings/vmlDrawing{n}.vml");
                let rid = crate::media::add_relationship(
                    &mut sheet_rels,
                    VML_REL_TYPE,
                    &format!("../drawings/vmlDrawing{n}.vml"),
                );
                self.pkg.add_default("vml", VML_CONTENT_TYPE)?;
                let mut root = el(
                    "xml",
                    &[
                        ("xmlns:v", "urn:schemas-microsoft-com:vml"),
                        ("xmlns:o", "urn:schemas-microsoft-com:office:office"),
                        ("xmlns:x", "urn:schemas-microsoft-com:office:excel"),
                    ],
                );
                let mut layout = el("o:shapelayout", &[("v:ext", "edit")]);
                layout.push(el("o:idmap", &[("v:ext", "edit"), ("data", "1")]));
                root.push(layout);
                let mut st = el(
                    "v:shapetype",
                    &[
                        ("id", "_x0000_t202"),
                        ("coordsize", "21600,21600"),
                        ("o:spt", "202"),
                        ("path", "m,l,21600r21600,l21600,xe"),
                    ],
                );
                st.push(el("v:stroke", &[("joinstyle", "miter")]));
                st.push(el("v:path", &[("gradientshapeok", "t"), ("o:connecttype", "rect")]));
                root.push(st);
                (part, rid, root)
            }
        };
        let row_s = row0.to_string();
        let col_s = col.to_string();
        vml.children.retain(|n| {
            !matches!(n, XmlNode::Element(e)
                if e.local_name() == "shape"
                    && e.child("ClientData").map(|cd| {
                        cd.child("Row").map(|r| r.text_content().trim() == row_s).unwrap_or(false)
                            && cd.child("Column").map(|c| c.text_content().trim() == col_s).unwrap_or(false)
                    }).unwrap_or(false))
        });
        let max_id = vml
            .children_named("shape")
            .iter()
            .filter_map(|s| {
                s.attr_local("id")?
                    .strip_prefix("_x0000_s")?
                    .parse::<u32>()
                    .ok()
            })
            .max()
            .unwrap_or(1024);
        let sid = format!("_x0000_s{}", max_id + 1);
        let mut shape = el(
            "v:shape",
            &[
                ("id", sid.as_str()),
                ("type", "#_x0000_t202"),
                (
                    "style",
                    "position:absolute;margin-left:80pt;margin-top:2pt;width:108pt;height:60pt;z-index:1;visibility:hidden",
                ),
                ("fillcolor", "#ffffe1"),
                ("o:insetmode", "auto"),
            ],
        );
        shape.push(el("v:fill", &[("color2", "#ffffe1")]));
        shape.push(el("v:shadow", &[("on", "t"), ("color", "black"), ("obscured", "t")]));
        shape.push(el("v:path", &[("o:connecttype", "none")]));
        shape.push(el("v:textbox", &[("style", "mso-direction-alt:auto")]));
        let mut cd = el("x:ClientData", &[("ObjectType", "Note")]);
        cd.push(XmlElement::new("x:MoveWithCells"));
        cd.push(XmlElement::new("x:SizeWithCells"));
        let mut anchor = XmlElement::new("x:Anchor");
        anchor.push_text(&format!(
            "{}, 15, {}, 2, {}, 15, {}, 2",
            col + 1,
            row0,
            col + 3,
            row0 + 4
        ));
        cd.push(anchor);
        let mut autofill = XmlElement::new("x:AutoFill");
        autofill.push_text("False");
        cd.push(autofill);
        let mut r_el = XmlElement::new("x:Row");
        r_el.push_text(&row_s);
        cd.push(r_el);
        let mut c_el = XmlElement::new("x:Column");
        c_el.push_text(&col_s);
        cd.push(c_el);
        shape.push(cd);
        vml.push(shape);
        self.pkg.put_xml(&vml_part, &vml)?;
        self.pkg.put_xml(&sheet_rels_part, &sheet_rels)?;

        let sheet_xml = &mut self.sheets[sheet_idx].xml;
        if sheet_xml.attr("xmlns:r").is_none() {
            sheet_xml.set_attr(
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            );
        }
        let ld = worksheet_child(sheet_xml, "legacyDrawing");
        ld.set_attr("r:id", &vml_rid);
        Ok(())
    }

    /// Delete a cell's comment (and its VML note shape); no-op when absent.
    fn remove_cell_comment(&mut self, sheet_idx: usize, col: u32, row0: u32) -> Result<()> {
        let sheet_part = self.sheets[sheet_idx].part.clone();
        let sheet_rels_part = sheet_part.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels";
        if !self.pkg.has_part(&sheet_rels_part) {
            return Ok(());
        }
        let sheet_rels = self.pkg.xml(&sheet_rels_part)?;
        let find = |rel_type: &str| -> Option<String> {
            sheet_rels
                .children_named("Relationship")
                .into_iter()
                .find(|r| r.attr_local("Type") == Some(rel_type))
                .and_then(|r| r.attr_local("Target"))
                .map(|t| resolve_target("xl/worksheets", t))
        };
        let ref_name = cell_name(col, row0);
        if let Some(part) = find(COMMENTS_REL_TYPE) {
            let mut comments = self.pkg.xml(&part)?;
            if let Some(list) = comments.child_mut("commentList") {
                list.children.retain(|n| {
                    !matches!(n, XmlNode::Element(e)
                        if e.local_name() == "comment"
                            && e.attr_local("ref") == Some(ref_name.as_str()))
                });
            }
            self.pkg.put_xml(&part, &comments)?;
        }
        if let Some(part) = find(VML_REL_TYPE) {
            let mut vml = self.pkg.xml(&part)?;
            let (row_s, col_s) = (row0.to_string(), col.to_string());
            vml.children.retain(|n| {
                !matches!(n, XmlNode::Element(e)
                    if e.local_name() == "shape"
                        && e.child("ClientData").map(|cd| {
                            cd.child("Row").map(|r| r.text_content().trim() == row_s).unwrap_or(false)
                                && cd.child("Column").map(|c| c.text_content().trim() == col_s).unwrap_or(false)
                        }).unwrap_or(false))
            });
            self.pkg.put_xml(&part, &vml)?;
        }
        Ok(())
    }

    /// (cell_ref, author, text) of every comment on the sheet.
    fn sheet_comments(&self, sheet: &Sheet) -> Vec<(String, String, String)> {
        let mut out = Vec::new();
        let rels_part = sheet.part.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels";
        if !self.pkg.has_part(&rels_part) {
            return out;
        }
        let Ok(rels) = self.pkg.xml(&rels_part) else { return out };
        let Some(part) = rels
            .children_named("Relationship")
            .into_iter()
            .find(|r| r.attr_local("Type") == Some(COMMENTS_REL_TYPE))
            .and_then(|r| r.attr_local("Target"))
            .map(|t| resolve_target("xl/worksheets", t))
        else {
            return out;
        };
        let Ok(comments) = self.pkg.xml(&part) else { return out };
        let authors: Vec<String> = comments
            .child("authors")
            .map(|a| {
                a.children_named("author")
                    .iter()
                    .map(|x| x.text_content())
                    .collect()
            })
            .unwrap_or_default();
        if let Some(list) = comments.child("commentList") {
            for c in list.children_named("comment") {
                let cell = c.attr_local("ref").unwrap_or_default().to_string();
                let author = c
                    .attr_local("authorId")
                    .and_then(|id| id.parse::<usize>().ok())
                    .and_then(|id| authors.get(id).cloned())
                    .unwrap_or_default();
                let text = c.child("text").map(|t| t.text_content()).unwrap_or_default();
                out.push((cell, author, text));
            }
        }
        out
    }

    /// Replace a cell's text with its computed value when it's a formula.
    fn attach_computed(&self, sheet_idx: usize, col: u32, row: u32, info: &mut NodeInfo) {
        let sheet = &self.sheets[sheet_idx];
        let formula = sheet
            .xml
            .child("sheetData")
            .and_then(|sd| find_row(sd, row + 1))
            .and_then(|r| find_cell(r, col, row))
            .and_then(|c| c.child("f"))
            .map(|f| f.text_content());
        let Some(formula) = formula else { return };
        let ev = crate::formula::Evaluator::new(self);
        let v = ev.eval_formula(&sheet.name, &formula);
        info.attr("computed", "true");
        info.attr("value-type", v.type_name());
        info.text = Some(v.display());
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

impl crate::formula::CellSource for Xlsx {
    fn cell(&self, sheet: &str, col: u32, row: u32) -> crate::formula::CellContent {
        use crate::formula::CellContent;
        let Some(idx) = self
            .sheets
            .iter()
            .position(|s| s.name.eq_ignore_ascii_case(sheet))
        else {
            return CellContent::Error("#REF!".into());
        };
        let s = &self.sheets[idx];
        let Some(c) = s
            .xml
            .child("sheetData")
            .and_then(|sd| find_row(sd, row + 1))
            .and_then(|r| find_cell(r, col, row))
        else {
            return CellContent::Empty;
        };
        if let Some(f) = c.child("f") {
            return CellContent::Formula(f.text_content());
        }
        let t = c.attr_local("t").unwrap_or("n");
        let v = c.child("v").map(|v| v.text_content()).unwrap_or_default();
        match t {
            "s" => {
                let i: usize = v.trim().parse().unwrap_or(usize::MAX);
                CellContent::Text(self.shared.get(i).cloned().unwrap_or_default())
            }
            "str" => CellContent::Text(v),
            "inlineStr" => {
                CellContent::Text(c.child("is").map(|is| is.text_content()).unwrap_or_default())
            }
            "b" => CellContent::Bool(v.trim() == "1"),
            "e" => CellContent::Error(v),
            _ => match v.trim().parse::<f64>() {
                Ok(n) => CellContent::Number(n),
                Err(_) if v.trim().is_empty() => CellContent::Empty,
                Err(_) => CellContent::Error("#VALUE!".into()),
            },
        }
    }

    fn has_sheet(&self, sheet: &str) -> bool {
        self.sheets.iter().any(|s| s.name.eq_ignore_ascii_case(sheet))
    }
}

fn parse_cell_ref_opt(r: &str) -> Option<(u32, u32)> {
    parse_cell_ref(r)
}

// -------------------------------------------------------------- csv ----

/// RFC 4180-ish CSV parser: quoted fields, doubled quotes, CRLF or LF.
pub(crate) fn parse_csv(text: &str) -> Vec<Vec<String>> {
    let mut rows = Vec::new();
    let mut row: Vec<String> = Vec::new();
    let mut field = String::new();
    let mut field_started = false;
    let mut in_quotes = false;
    let mut chars = text.chars().peekable();
    while let Some(c) = chars.next() {
        if in_quotes {
            match c {
                '"' => {
                    if chars.peek() == Some(&'"') {
                        field.push('"');
                        chars.next();
                    } else {
                        in_quotes = false;
                    }
                }
                c => field.push(c),
            }
        } else {
            match c {
                '"' if field.is_empty() => {
                    in_quotes = true;
                    field_started = true;
                }
                ',' => {
                    row.push(std::mem::take(&mut field));
                    field_started = false;
                }
                '\r' => {}
                '\n' => {
                    row.push(std::mem::take(&mut field));
                    rows.push(std::mem::take(&mut row));
                    field_started = false;
                }
                c => field.push(c),
            }
        }
    }
    if !field.is_empty() || field_started || !row.is_empty() {
        row.push(field);
        rows.push(row);
    }
    rows
}

fn csv_field(s: &str) -> String {
    if s.contains([',', '"', '\n', '\r']) {
        format!("\"{}\"", s.replace('"', "\"\""))
    } else {
        s.to_string()
    }
}

// ------------------------------------------------------------- dates ----
// Excel stores dates as serial numbers: days since 1899-12-30, with the
// time of day as the fraction.

/// Days from 1970-01-01 (Howard Hinnant's days-from-civil algorithm).
fn days_from_civil(y: i64, m: u32, d: u32) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let doy = (153 * (if m > 2 { m - 3 } else { m + 9 }) as i64 + 2) / 5 + d as i64 - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

const UNIX_EPOCH_SERIAL: f64 = 25_569.0; // 1970-01-01 as an Excel serial

/// Excel serial for a civil date (shared with the formula evaluator).
pub(crate) fn date_serial(y: i64, m: u32, d: u32) -> f64 {
    days_from_civil(y, m, d) as f64 + UNIX_EPOCH_SERIAL
}

/// (year, month, day) of an Excel serial (shared with the evaluator).
pub(crate) fn serial_civil(serial: f64) -> (i64, u32, u32) {
    let days = serial.floor() as i64;
    let z = days - UNIX_EPOCH_SERIAL as i64 + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = if m <= 2 { y + 1 } else { y };
    (y, m as u32, d as u32)
}

/// Parse `YYYY-MM-DD`, `YYYY-MM-DD HH:MM[:SS]` (or `T` separator), or a
/// bare `HH:MM:SS` into (serial, has_date, has_time).
fn parse_iso_datetime(s: &str) -> Option<(f64, bool, bool)> {
    let s = s.trim();
    let parse_time = |t: &str| -> Option<f64> {
        let parts: Vec<&str> = t.split(':').collect();
        if parts.len() < 2 || parts.len() > 3 {
            return None;
        }
        let h: u32 = parts[0].parse().ok().filter(|h| *h < 24)?;
        let m: u32 = parts[1].parse().ok().filter(|m| *m < 60)?;
        let sec: f64 = if parts.len() == 3 {
            parts[2].parse().ok().filter(|s| *s < 60.0)?
        } else {
            0.0
        };
        if parts[0].len() > 2 || parts[1].len() != 2 {
            return None;
        }
        Some((h as f64 * 3600.0 + m as f64 * 60.0 + sec) / 86_400.0)
    };
    // Bare time.
    if s.contains(':') && !s.contains('-') {
        return parse_time(s).map(|frac| (frac, false, true));
    }
    let (date_part, time_part) = match s.split_once([' ', 'T']) {
        Some((d, t)) => (d, Some(t)),
        None => (s, None),
    };
    let dp: Vec<&str> = date_part.split('-').collect();
    if dp.len() != 3 || dp[0].len() != 4 || dp[1].len() != 2 || dp[2].len() != 2 {
        return None;
    }
    let y: i64 = dp[0].parse().ok()?;
    let m: u32 = dp[1].parse().ok().filter(|m| (1..=12).contains(m))?;
    let d: u32 = dp[2].parse().ok().filter(|d| (1..=31).contains(d))?;
    let serial = days_from_civil(y, m, d) as f64 + UNIX_EPOCH_SERIAL;
    match time_part {
        Some(t) => parse_time(t).map(|frac| (serial + frac, true, true)),
        None => Some((serial, true, false)),
    }
}

/// Render a serial back as ISO (date / datetime / time as appropriate).
fn serial_to_iso(serial: f64, with_date: bool, with_time: bool) -> String {
    let days = serial.floor();
    let mut secs = ((serial - days) * 86_400.0).round() as i64;
    let mut days = days;
    if secs >= 86_400 {
        days += 1.0;
        secs -= 86_400;
    }
    let (y, m, d) = serial_civil(days);
    let date = format!("{:04}-{:02}-{:02}", y, m, d);
    let time = format!("{:02}:{:02}:{:02}", secs / 3600, (secs / 60) % 60, secs % 60);
    match (with_date, with_time) {
        (true, true) => format!("{date} {time}"),
        (false, true) => time,
        _ => date,
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum DateKind {
    Date,
    DateTime,
    Time,
}

/// Classify a number format code as date-like.
fn code_date_kind(code: &str) -> Option<DateKind> {
    // Strip literals in quotes and color/condition brackets.
    let mut clean = String::new();
    let mut in_quote = false;
    let mut in_bracket = false;
    for ch in code.chars() {
        match ch {
            '"' => in_quote = !in_quote,
            '[' if !in_quote => in_bracket = true,
            ']' if !in_quote => in_bracket = false,
            c if !in_quote && !in_bracket => clean.push(c.to_ascii_lowercase()),
            _ => {}
        }
    }
    let has_time = clean.contains('h') || clean.contains("ss");
    let has_date = clean.contains('y') || clean.contains('d');
    match (has_date, has_time) {
        (true, true) => Some(DateKind::DateTime),
        (true, false) => Some(DateKind::Date),
        (false, true) => Some(DateKind::Time),
        _ => None,
    }
}

fn builtin_date_kind(id: u32) -> Option<DateKind> {
    match id {
        14..=17 => Some(DateKind::Date),
        18..=21 | 45..=47 => Some(DateKind::Time),
        22 => Some(DateKind::DateTime),
        _ => None,
    }
}

/// Resolve a friendly format name (or raw code) to a numFmtId, creating a
/// custom <numFmt> entry when needed.
fn ensure_num_fmt(styles: &mut XmlElement, name: &str) -> u32 {
    let builtin = match name.to_ascii_lowercase().as_str() {
        "general" => Some(0),
        "integer" | "int" | "0" => Some(1),
        "decimal" | "0.00" => Some(2),
        "thousands" | "#,##0" => Some(3),
        "percent" | "0%" => Some(9),
        "percent2" | "0.00%" => Some(10),
        "date" => Some(14),
        "time" => Some(21),
        "datetime" => Some(22),
        "text" => Some(49),
        _ => None,
    };
    if let Some(id) = builtin {
        return id;
    }
    let code = match name.to_ascii_lowercase().as_str() {
        "currency" => "$#,##0.00",
        _ => name,
    };
    // numFmts must be the first child of styleSheet.
    let num_fmts = styles.ensure_child("numFmts", "numFmts", true);
    if let Some(existing) = num_fmts
        .children_named("numFmt")
        .into_iter()
        .find(|f| f.attr_local("formatCode") == Some(code))
        .and_then(|f| f.attr_local("numFmtId"))
        .and_then(|v| v.parse::<u32>().ok())
    {
        return existing;
    }
    let id = num_fmts
        .children_named("numFmt")
        .into_iter()
        .filter_map(|f| f.attr_local("numFmtId").and_then(|v| v.parse::<u32>().ok()))
        .max()
        .unwrap_or(163)
        .max(163)
        + 1;
    num_fmts.push(el(
        "numFmt",
        &[("numFmtId", id.to_string().as_str()), ("formatCode", code)],
    ));
    let count = num_fmts.children_named("numFmt").len();
    num_fmts.set_attr("count", &count.to_string());
    id
}

/// Friendly name (or raw code) for a numFmtId, for dump replay.
fn num_fmt_name(styles: &XmlElement, id: u32) -> Option<String> {
    let name = match id {
        0 => return None,
        1 => "integer",
        2 => "0.00",
        3 => "thousands",
        9 => "percent",
        10 => "0.00%",
        14 => "date",
        21 => "time",
        22 => "datetime",
        49 => "text",
        _ => {
            return styles
                .child("numFmts")
                .and_then(|nf| {
                    nf.children_named("numFmt")
                        .into_iter()
                        .find(|f| {
                            f.attr_local("numFmtId").and_then(|v| v.parse::<u32>().ok())
                                == Some(id)
                        })
                        .and_then(|f| f.attr_local("formatCode"))
                        .map(|c| c.to_string())
                })
        }
    };
    Some(name.to_string())
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

const EMPTY_RELS_XML: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

const COMMENTS_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const VML_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const VML_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";

/// Get-or-create a worksheet child at its schema-mandated position.
fn worksheet_child<'a>(sheet_xml: &'a mut XmlElement, name: &str) -> &'a mut XmlElement {
    const ORDER: &[&str] = &[
        "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
        "sheetCalcPr", "sheetProtection", "autoFilter", "sortState", "mergeCells",
        "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
        "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "drawing",
        "legacyDrawing",
    ];
    if sheet_xml.child(name).is_none() {
        let rank = ORDER.iter().position(|o| *o == name).unwrap_or(ORDER.len());
        let mut at = sheet_xml.children.len();
        for (i, node) in sheet_xml.children.iter().enumerate() {
            if let XmlNode::Element(e) = node {
                let r = ORDER
                    .iter()
                    .position(|o| *o == e.local_name())
                    .unwrap_or(ORDER.len());
                if r > rank {
                    at = i;
                    break;
                }
            }
        }
        sheet_xml
            .children
            .insert(at, XmlNode::Element(XmlElement::new(name)));
    }
    sheet_xml.child_mut(name).unwrap()
}

fn next_rid(rels: &XmlElement) -> String {
    let n = rels
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
    format!("rId{n}")
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

fn find_cell(row: &XmlElement, col: u32, row0: u32) -> Option<&XmlElement> {
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
    underline: Option<bool>,
    /// Number format: friendly name or raw format code.
    num_fmt: Option<String>,
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
            underline: get("underline").map(parse_bool).transpose()?,
            num_fmt: props.get("format").map(|s| s.to_string()),
        })
    }

    fn any(&self) -> bool {
        self.bold.is_some()
            || self.italic.is_some()
            || self.color.is_some()
            || self.size.is_some()
            || self.fill.is_some()
            || self.font.is_some()
            || self.underline.is_some()
            || self.num_fmt.is_some()
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
    if let Some(u) = req.underline {
        toggle(&mut font, "u", u);
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
    if let Some(fmt) = &req.num_fmt {
        let id = ensure_num_fmt(styles, fmt);
        xf.set_attr("numFmtId", &id.to_string());
        xf.set_attr("applyNumberFormat", "1");
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

/// Replayable style props (bold/italic/color/size/font/fill) for a cellXfs
/// index. Default font/fill entries are skipped to keep dumps quiet.
fn style_props(styles: &XmlElement, xf_idx: usize) -> Vec<(String, String)> {
    let mut out = Vec::new();
    let Some(xf) = styles
        .child("cellXfs")
        .and_then(|xfs| xfs.children_named("xf").get(xf_idx).copied())
    else {
        return out;
    };
    let font_id: usize = xf
        .attr_local("fontId")
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    if font_id != 0 {
        if let Some(font) = styles
            .child("fonts")
            .and_then(|f| f.children_named("font").get(font_id).copied())
        {
            if font.child("b").is_some() {
                out.push(("bold".to_string(), "true".to_string()));
            }
            if font.child("i").is_some() {
                out.push(("italic".to_string(), "true".to_string()));
            }
            if font.child("u").is_some() {
                out.push(("underline".to_string(), "true".to_string()));
            }
            if let Some(rgb) = font.child("color").and_then(|c| c.attr_local("rgb")) {
                let hex = rgb.strip_prefix("FF").unwrap_or(rgb);
                out.push(("color".to_string(), hex.to_string()));
            }
            if let Some(sz) = font.child("sz").and_then(|s| s.attr_local("val")) {
                if sz != "11" {
                    out.push(("size".to_string(), sz.to_string()));
                }
            }
            if let Some(name) = font.child("name").and_then(|n| n.attr_local("val")) {
                if name != "Calibri" {
                    out.push(("font".to_string(), name.to_string()));
                }
            }
        }
    }
    if let Some(id) = xf
        .attr_local("numFmtId")
        .and_then(|v| v.parse::<u32>().ok())
        .filter(|id| *id != 0)
    {
        // Date-like builtins are implied by the ISO value in the dump, but
        // custom codes must be carried explicitly.
        if let Some(name) = num_fmt_name(styles, id) {
            if !matches!(name.as_str(), "date" | "datetime" | "time") {
                out.push(("format".to_string(), name));
            }
        }
    }
    let fill_id: usize = xf
        .attr_local("fillId")
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    // Fill 0 (none) and 1 (gray125) are the mandatory defaults.
    if fill_id > 1 {
        if let Some(rgb) = styles
            .child("fills")
            .and_then(|f| f.children_named("fill").get(fill_id).copied())
            .and_then(|f| f.child("patternFill"))
            .and_then(|p| p.child("fgColor"))
            .and_then(|c| c.attr_local("rgb"))
        {
            let hex = rgb.strip_prefix("FF").unwrap_or(rgb);
            out.push(("fill".to_string(), hex.to_string()));
        }
    }
    out
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
            "html" => {
                let mut out = String::new();
                for sheet in &self.sheets {
                    out.push_str(&format!(
                        "<h2 class=\"sheet-name\">{}</h2>\n",
                        crate::html::escape(&sheet.name)
                    ));
                    let Some(((c0, r0), (c1, r1))) = Self::used_range(sheet) else {
                        out.push_str("<p><em>(empty)</em></p>\n");
                        continue;
                    };
                    out.push_str("<table>\n<tr><th></th>");
                    for c in c0..=c1 {
                        out.push_str(&format!("<th>{}</th>", col_letters(c)));
                    }
                    out.push_str("</tr>\n");
                    let sd = sheet.xml.child("sheetData").unwrap();
                    for r in r0..=r1 {
                        out.push_str(&format!("<tr><th class=\"row-num\">{}</th>", r + 1));
                        for col in c0..=c1 {
                            let value = find_row(sd, r + 1)
                                .and_then(|row| find_cell(row, col, r))
                                .map(|c| self.cell_display(c).0)
                                .unwrap_or_default();
                            out.push_str(&format!("<td>{}</td>", crate::html::escape(&value)));
                        }
                        out.push_str("</tr>\n");
                    }
                    out.push_str("</table>\n");
                }
                Ok(Report::Text(crate::html::page("Workbook", &out)))
            }
            "comments" => {
                let mut out = String::new();
                let mut items = Vec::new();
                for sheet in &self.sheets {
                    for (cell, author, text) in self.sheet_comments(sheet) {
                        out.push_str(&format!("{}!{} [{}] {}\n", sheet.name, cell, author, text));
                        items.push(json!({
                            "sheet": sheet.name, "cell": cell,
                            "author": author, "text": text,
                        }));
                    }
                }
                if out.is_empty() {
                    out = "(no comments)".into();
                }
                Ok(Report::Data { text: out, data: json!(items) })
            }
            other => bail!("unknown view mode '{other}' for xlsx (text/outline/stats/html/comments)"),
        }
    }

    fn get(&mut self, path_str: &str, depth: usize, computed: bool) -> Result<Report> {
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
                            let mut info = self.cell_info(sheet, c, &cpath);
                            if computed {
                                self.attach_computed(sheet_idx, col, r, &mut info);
                            }
                            nodes.push(info);
                        }
                    }
                }
                return Ok(Report::Nodes(nodes));
            }
            if let Some((col, row0)) = parse_cell_ref(&seg.name) {
                let sd = sheet.xml.child("sheetData").context("no sheetData")?;
                let cpath = format!("/{}/{}", sheet.name, cell_name(col, row0));
                let mut info = match find_row(sd, row0 + 1).and_then(|r| find_cell(r, col, row0)) {
                    Some(c) => self.cell_info(sheet, c, &cpath),
                    None => {
                        let mut info = NodeInfo::new(cpath, "cell");
                        info.text = Some(String::new());
                        info.attr("type", "empty");
                        info
                    }
                };
                if computed {
                    self.attach_computed(sheet_idx, col, row0, &mut info);
                }
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
                self.create_sheet(&name)?;
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
                let mut needs_shift = false;
                {
                    let sheet = &mut self.sheets[sheet_idx];
                    let sd = sheet
                        .xml
                        .child_mut("sheetData")
                        .context("worksheet has no <sheetData>")?;
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
                if needs_shift {
                    self.rewrite_all_formulas(sheet_idx, &Shift::Rows { from: row_num, delta: 1 });
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
            "column" | "col" => {
                if dpath.segments.is_empty() {
                    bail!("add column needs a sheet path, e.g. /Sheet1");
                }
                let sheet_idx = self.sheet_index(&dpath.segments[0])?;
                // Column: --index is the 1-based column number (like row add),
                // or --prop at=B by letter.
                let col: u32 = if let Some(at) = props.get("at") {
                    letters_to_col(at)
                        .with_context(|| format!("'{at}' is not a column letter like B"))?
                } else {
                    match pos {
                        Position::Index(n) => {
                            if *n == 0 {
                                bail!("column numbers are 1-based");
                            }
                            *n as u32 - 1
                        }
                        Position::Append => {
                            let sheet = &self.sheets[sheet_idx];
                            Self::used_range(sheet).map(|((_, _), (c1, _))| c1 + 1).unwrap_or(0)
                        }
                        _ => bail!(
                            "add column supports --index N (1-based column number), --prop at=B, or append"
                        ),
                    }
                };
                // Shift existing cells at or right of the insertion point.
                let mut needs_shift = false;
                {
                    let sheet = &mut self.sheets[sheet_idx];
                    let sd = sheet
                        .xml
                        .child_mut("sheetData")
                        .context("worksheet has no <sheetData>")?;
                    'outer: for row in sd.children.iter().filter_map(|n| n.as_element()) {
                        if row.local_name() != "row" {
                            continue;
                        }
                        for c in row.children_named("c") {
                            if let Some((cc, _)) = c.attr_local("r").and_then(parse_cell_ref_opt) {
                                if cc >= col {
                                    needs_shift = true;
                                    break 'outer;
                                }
                            }
                        }
                    }
                    if needs_shift {
                        shift_cols(sd, col, 1);
                    }
                }
                if needs_shift {
                    self.rewrite_all_formulas(sheet_idx, &Shift::Cols { from: col, delta: 1 });
                }
                // Fill values: values="a,b,c" fills down from row 1, and/or
                // rN=value props for specific rows.
                if let Some(values) = props.get("values") {
                    let values = values.to_string();
                    for (i, v) in values.split(',').enumerate() {
                        let value_props =
                            Props::from_pairs(vec![("value".to_string(), v.trim().to_string())]);
                        self.set_cell(sheet_idx, col, i as u32, &value_props)?;
                    }
                }
                let row_props: Vec<(u32, String)> = props
                    .iter()
                    .filter_map(|(k, v)| {
                        k.strip_prefix('r')
                            .and_then(|n| n.parse::<u32>().ok())
                            .filter(|n| *n >= 1)
                            .map(|n| (n - 1, v.to_string()))
                    })
                    .collect();
                for (row0, v) in row_props {
                    let value_props = Props::from_pairs(vec![("value".to_string(), v)]);
                    self.set_cell(sheet_idx, col, row0, &value_props)?;
                }
                let sheet = &self.sheets[sheet_idx];
                let letters = col_letters(col);
                let mut info =
                    NodeInfo::new(format!("/{}/col[{}]", sheet.name, col + 1), "column");
                info.attr("column", letters);
                Ok(Report::Nodes(vec![info]))
            }
            "chart" => {
                if dpath.segments.is_empty() {
                    bail!("charts are added to a sheet, e.g. officecli add data.xlsx /Sheet1 --type chart --prop data=A1:B5");
                }
                let sheet_idx = self.sheet_index(&dpath.segments[0])?;
                self.add_chart(sheet_idx, props)
            }
            "csv" => {
                if dpath.segments.is_empty() {
                    bail!("csv is imported into a sheet, e.g. officecli add data.xlsx /Sheet1 --type csv --prop src=file.csv");
                }
                let sheet_idx = self.sheet_index(&dpath.segments[0])?;
                let text = if let Some(src) = props.get("src") {
                    std::fs::read_to_string(src)
                        .with_context(|| format!("cannot read CSV file '{src}'"))?
                } else if let Some(data) = props.get("data") {
                    data.replace("\\n", "\n")
                } else {
                    bail!("csv needs --prop src=file.csv (or inline --prop data=...)");
                };
                let (at_col, at_row) = props
                    .get("at")
                    .map(|a| {
                        parse_cell_ref(a).with_context(|| format!("'{a}' is not a cell reference"))
                    })
                    .transpose()?
                    .unwrap_or((0, 0));
                let rows = parse_csv(&text);
                let mut cells = 0usize;
                let mut max_cols = 0usize;
                for (ri, row) in rows.iter().enumerate() {
                    max_cols = max_cols.max(row.len());
                    for (ci, field) in row.iter().enumerate() {
                        if field.is_empty() {
                            continue;
                        }
                        let value_props =
                            Props::from_pairs(vec![("value".to_string(), field.clone())]);
                        self.set_cell(sheet_idx, at_col + ci as u32, at_row + ri as u32, &value_props)?;
                        cells += 1;
                    }
                }
                let sheet_name = self.sheets[sheet_idx].name.clone();
                let to = cell_name(
                    at_col + max_cols.saturating_sub(1) as u32,
                    at_row + rows.len().saturating_sub(1) as u32,
                );
                let mut info = NodeInfo::new(
                    format!("/{}/{}:{}", sheet_name, cell_name(at_col, at_row), to),
                    "range",
                );
                info.attr("rows", rows.len().to_string());
                info.attr("cells", cells.to_string());
                Ok(Report::Nodes(vec![info]))
            }
            "pivot" => {
                let source_sheet = if dpath.segments.is_empty() {
                    0
                } else {
                    self.sheet_index(&dpath.segments[0])?
                };
                self.add_pivot(source_sheet, props)
            }
            other => bail!("unsupported xlsx element type '{other}' (sheet/row/column/chart/pivot; cells are created via set)"),
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

        // Range: merge / unmerge.
        if seg.name.contains(':') {
            let merge = props
                .get_bool("merge")?
                .context("set on a range supports --prop merge=true|false")?;
            let (_, (c0, r0), (c1, r1)) =
                self.parse_range_spec(sheet_idx, &format!("{}!{}", self.sheets[sheet_idx].name, seg.name))?;
            let reference = format!("{}:{}", cell_name(c0, r0), cell_name(c1, r1));
            let sheet_xml = &mut self.sheets[sheet_idx].xml;
            let merges = worksheet_child(sheet_xml, "mergeCells");
            merges.children.retain(|n| {
                !matches!(n, XmlNode::Element(e)
                    if e.local_name() == "mergeCell"
                        && e.attr_local("ref") == Some(reference.as_str()))
            });
            if merge {
                merges.push(el("mergeCell", &[("ref", reference.as_str())]));
            }
            let count = merges.children_named("mergeCell").len();
            merges.set_attr("count", &count.to_string());
            if count == 0 {
                sheet_xml.children.retain(|n| {
                    !matches!(n, XmlNode::Element(e) if e.local_name() == "mergeCells")
                });
            }
            return Ok(Report::Data {
                text: format!(
                    "{} {} in {}",
                    if merge { "merged" } else { "unmerged" },
                    reference,
                    self.sheets[sheet_idx].name
                ),
                data: json!({ "merged": merge, "ref": reference }),
            });
        }

        // Row: height.
        if seg.name.eq_ignore_ascii_case("row") {
            let n = seg.index().context("row needs an index, e.g. row[5]")? as u32;
            let height: f64 = props
                .get("height")
                .context("set on a row supports --prop height=N (points)")?
                .parse()
                .context("height must be a number of points")?;
            let sheet_xml = &mut self.sheets[sheet_idx].xml;
            let sd = sheet_xml
                .child_mut("sheetData")
                .context("worksheet has no <sheetData>")?;
            ensure_row(sd, n);
            for node in sd.children.iter_mut().filter_map(|x| x.as_element_mut()) {
                if node.local_name() == "row"
                    && node.attr_local("r").and_then(|v| v.parse::<u32>().ok()) == Some(n)
                {
                    node.set_attr("ht", &format_num(height));
                    node.set_attr("customHeight", "1");
                }
            }
            return Ok(Report::Data {
                text: format!("set row {n} height to {} in {}", format_num(height), self.sheets[sheet_idx].name),
                data: json!({ "row": n, "height": height }),
            });
        }

        // Column: width.
        let col_target: Option<u32> = if seg.name.eq_ignore_ascii_case("col")
            || seg.name.eq_ignore_ascii_case("column")
        {
            Some(seg.index().context("column needs an index, e.g. col[2]")? as u32 - 1)
        } else if seg.name.len() <= 3
            && seg.name.chars().all(|c| c.is_ascii_alphabetic())
            && parse_cell_ref(&seg.name).is_none()
        {
            letters_to_col(&seg.name)
        } else {
            None
        };
        if let Some(col) = col_target {
            let width: f64 = props
                .get("width")
                .context("set on a column supports --prop width=N (characters)")?
                .parse()
                .context("width must be a number of characters")?;
            let n = (col + 1).to_string();
            let sheet_xml = &mut self.sheets[sheet_idx].xml;
            let cols = worksheet_child(sheet_xml, "cols");
            cols.children.retain(|node| {
                !matches!(node, XmlNode::Element(e)
                    if e.local_name() == "col"
                        && e.attr_local("min") == Some(n.as_str())
                        && e.attr_local("max") == Some(n.as_str()))
            });
            cols.push(el(
                "col",
                &[
                    ("min", n.as_str()),
                    ("max", n.as_str()),
                    ("width", format_num(width).as_str()),
                    ("customWidth", "1"),
                ],
            ));
            return Ok(Report::Data {
                text: format!(
                    "set column {} width to {} in {}",
                    col_letters(col),
                    format_num(width),
                    self.sheets[sheet_idx].name
                ),
                data: json!({ "column": col_letters(col), "width": width }),
            });
        }

        let (col, row0) = parse_cell_ref(&seg.name)
            .with_context(|| format!("'{}' is not a cell reference like B2", seg.name))?;
        if !props.has("value")
            && !props.has("formula")
            && !props.has("url")
            && !props.has("comment")
            && !StyleRequest::from_props(props)?.any()
        {
            bail!("set on a cell needs --prop value=... , formula=... , url=... , comment=... , or style props (bold/italic/color/size/fill/font/format)");
        }
        let mut info = if props.has("value") || props.has("formula") || StyleRequest::from_props(props)?.any() {
            self.set_cell(sheet_idx, col, row0, props)?
        } else {
            NodeInfo::new(
                format!("/{}/{}", self.sheets[sheet_idx].name, cell_name(col, row0)),
                "cell",
            )
        };
        if let Some(url) = props.get("url") {
            let url = url.to_string();
            self.set_cell_hyperlink(sheet_idx, col, row0, &url)?;
            info.attr("url", url);
        }
        if let Some(comment) = props.get("comment") {
            let comment = comment.to_string();
            if comment.is_empty() {
                self.remove_cell_comment(sheet_idx, col, row0)?;
                info.attr("comment", "(removed)");
            } else {
                let author = props.get("author").unwrap_or("officecli").to_string();
                self.set_cell_comment(sheet_idx, col, row0, &comment, &author)?;
                info.attr("comment", comment);
            }
        }
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
            self.rewrite_all_formulas(sheet_idx, &Shift::Rows { from: n + 1, delta: -1 });
            let sheet = &self.sheets[sheet_idx];
            return Ok(Report::Data {
                text: format!("removed row {n} from {}", sheet.name),
                data: json!({ "removed": format!("/{}/row[{}]", sheet.name, n) }),
            });
        }
        // Column: /Sheet1/col[2] (1-based) or /Sheet1/B (letters).
        let col_target: Option<u32> = if seg.name.eq_ignore_ascii_case("col")
            || seg.name.eq_ignore_ascii_case("column")
        {
            let n = seg.index().context("column needs an index, e.g. col[2]")? as u32;
            Some(n - 1)
        } else if seg.name.len() <= 3 && seg.name.chars().all(|c| c.is_ascii_alphabetic()) {
            letters_to_col(&seg.name)
        } else {
            None
        };
        if let Some(col) = col_target {
            let removed = remove_col_cells(sd, col);
            if removed == 0 {
                bail!("column {} has no content in {}", col_letters(col), sheet.name);
            }
            shift_cols(sd, col + 1, -1);
            self.rewrite_all_formulas(sheet_idx, &Shift::Cols { from: col + 1, delta: -1 });
            let sheet = &self.sheets[sheet_idx];
            return Ok(Report::Data {
                text: format!("removed column {} from {}", col_letters(col), sheet.name),
                data: json!({ "removed": format!("/{}/col[{}]", sheet.name, col + 1) }),
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

    fn tree(&mut self) -> Result<Vec<NodeInfo>> {
        let mut roots = Vec::new();
        for sheet in &self.sheets {
            let mut s = NodeInfo::new(format!("/{}", sheet.name), "sheet");
            s.attr("name", &sheet.name);
            if let Some(sd) = sheet.xml.child("sheetData") {
                for row in sd.children_named("row") {
                    let rnum = row.attr_local("r").unwrap_or("?");
                    let mut ri =
                        NodeInfo::new(format!("/{}/row[{}]", sheet.name, rnum), "row");
                    for c in row.children_named("c") {
                        let r = c.attr_local("r").unwrap_or("?");
                        let cpath = format!("/{}/{}", sheet.name, r);
                        ri.children.push(self.cell_info(sheet, c, &cpath));
                    }
                    s.children.push(ri);
                }
            }
            roots.push(s);
        }
        Ok(roots)
    }

    fn move_el(&mut self, path_str: &str, _to: Option<&str>, pos: &Position) -> Result<Report> {
        let dpath = path::parse(path_str)?;
        if dpath.segments.len() != 1 {
            bail!("xlsx move supports reordering sheets only, e.g. move /Sheet2 --index 0");
        }
        let from = self.sheet_index(&dpath.segments[0])?;
        let to = match pos {
            Position::Index(n) => (*n).min(self.sheets.len() - 1),
            _ => bail!("xlsx move needs --index N (0-based target position)"),
        };
        let sheet = self.sheets.remove(from);
        let name = sheet.name.clone();
        self.sheets.insert(to, sheet);
        // Mirror the order in workbook.xml <sheets>.
        if let Some(sheets_el) = self.workbook.child_mut("sheets") {
            let pos_in_children = sheets_el.children.iter().position(|n| {
                matches!(n, XmlNode::Element(e)
                    if e.local_name() == "sheet" && e.attr_local("name") == Some(&name))
            });
            if let Some(i) = pos_in_children {
                let entry = sheets_el.children.remove(i);
                // Find raw child index of the `to`-th sheet element (or end).
                let mut seen = 0;
                let mut insert_at = sheets_el.children.len();
                for (ci, n) in sheets_el.children.iter().enumerate() {
                    if matches!(n, XmlNode::Element(e) if e.local_name() == "sheet") {
                        if seen == to {
                            insert_at = ci;
                            break;
                        }
                        seen += 1;
                    }
                }
                sheets_el.children.insert(insert_at, entry);
            }
        }
        Ok(Report::Data {
            text: format!("moved sheet '{name}' to position {}", to + 1),
            data: json!({ "moved": name, "position": to + 1 }),
        })
    }

    fn calc(&mut self, expr: &str) -> Result<Report> {
        let sheet_name = self
            .sheets
            .first()
            .map(|s| s.name.clone())
            .context("workbook has no sheets")?;
        let ev = crate::formula::Evaluator::new(&*self);
        let v = ev.eval_formula(&sheet_name, expr);
        Ok(Report::Data {
            text: v.display(),
            data: json!({ "value": v.display(), "type": v.type_name() }),
        })
    }

    fn sort(&mut self, range: &str, by: &str, descending: bool) -> Result<Report> {
        let (idx, (c0, r0), (c1, r1)) = self.parse_range_spec(0, range)?;
        let by_col = letters_to_col(by)
            .with_context(|| format!("--by needs a column letter like B, got '{by}'"))?;
        if by_col < c0 || by_col > c1 {
            bail!("--by column {by} is outside the range {range}");
        }
        // Sorting rearranges values; formulas would silently break.
        {
            let sheet = &self.sheets[idx];
            let sd = sheet.xml.child("sheetData").context("no sheetData")?;
            for r in r0..=r1 {
                let Some(row) = find_row(sd, r + 1) else { continue };
                for c in c0..=c1 {
                    if find_cell(row, c, r).map(|c| c.child("f").is_some()).unwrap_or(false) {
                        bail!(
                            "range {range} contains formulas at {} — sort works on values only",
                            cell_name(c, r)
                        );
                    }
                }
            }
        }
        // Snapshot the range's cell elements row by row.
        let mut rows: Vec<Vec<Option<XmlElement>>> = Vec::new();
        {
            let sheet = &self.sheets[idx];
            let sd = sheet.xml.child("sheetData").context("no sheetData")?;
            for r in r0..=r1 {
                let mut cells = Vec::new();
                for c in c0..=c1 {
                    cells.push(
                        find_row(sd, r + 1)
                            .and_then(|row| find_cell(row, c, r))
                            .cloned(),
                    );
                }
                rows.push(cells);
            }
        }
        // Sort keys: numeric when possible, else case-insensitive text;
        // empty cells always sink to the bottom.
        let key_idx = (by_col - c0) as usize;
        let display = |cell: &Option<XmlElement>| -> String {
            cell.as_ref().map(|c| self.cell_display(c).0).unwrap_or_default()
        };
        rows.sort_by(|a, b| {
            let (ka, kb) = (display(&a[key_idx]), display(&b[key_idx]));
            match (ka.is_empty(), kb.is_empty()) {
                (true, true) => return std::cmp::Ordering::Equal,
                (true, false) => return std::cmp::Ordering::Greater,
                (false, true) => return std::cmp::Ordering::Less,
                _ => {}
            }
            let ord = match (ka.parse::<f64>(), kb.parse::<f64>()) {
                (Ok(x), Ok(y)) => x.partial_cmp(&y).unwrap_or(std::cmp::Ordering::Equal),
                _ => ka.to_lowercase().cmp(&kb.to_lowercase()),
            };
            if descending { ord.reverse() } else { ord }
        });
        // Write the rearranged cells back.
        {
            let sheet = &mut self.sheets[idx];
            let sd = sheet.xml.child_mut("sheetData").context("no sheetData")?;
            for (offset, cells) in rows.into_iter().enumerate() {
                let target_row = r0 + offset as u32;
                for (ci, cell) in cells.into_iter().enumerate() {
                    let col = c0 + ci as u32;
                    match cell {
                        Some(mut cell) => {
                            cell.set_attr("r", &cell_name(col, target_row));
                            let slot = ensure_cell(sd, col, target_row);
                            *slot = cell;
                        }
                        None => {
                            // Clear any leftover value in this slot.
                            if let Some(row) = sd
                                .children
                                .iter_mut()
                                .filter_map(|n| n.as_element_mut())
                                .find(|e| {
                                    e.local_name() == "row"
                                        && e.attr_local("r")
                                            .and_then(|v| v.parse::<u32>().ok())
                                            == Some(target_row + 1)
                                })
                            {
                                let target = cell_name(col, target_row);
                                row.children.retain(|n| {
                                    !matches!(n, XmlNode::Element(e)
                                        if e.local_name() == "c"
                                            && e.attr_local("r") == Some(target.as_str()))
                                });
                            }
                        }
                    }
                }
            }
        }
        let sheet_name = &self.sheets[idx].name;
        Ok(Report::Data {
            text: format!(
                "sorted {} row(s) in {}!{}:{} by column {}{}",
                r1 - r0 + 1,
                sheet_name,
                cell_name(c0, r0),
                cell_name(c1, r1),
                by.to_uppercase(),
                if descending { " (descending)" } else { "" }
            ),
            data: json!({
                "sorted": r1 - r0 + 1,
                "by": by.to_uppercase(),
                "descending": descending,
            }),
        })
    }

    fn export_csv(&mut self, sheet: Option<&str>, range: Option<&str>) -> Result<String> {
        let default_idx = match sheet {
            Some(name) => self
                .sheets
                .iter()
                .position(|s| s.name == name)
                .with_context(|| format!("no sheet named '{name}'"))?,
            None => 0,
        };
        let (idx, (c0, r0), (c1, r1)) = match range {
            Some(spec) => self.parse_range_spec(default_idx, spec)?,
            None => match Self::used_range(&self.sheets[default_idx]) {
                Some((from, to)) => (default_idx, from, to),
                None => return Ok(String::new()),
            },
        };
        let mut out = String::new();
        for r in r0..=r1 {
            let mut fields = Vec::new();
            for col in c0..=c1 {
                fields.push(csv_field(&self.value_at(idx, col, r)));
            }
            out.push_str(&fields.join(","));
            out.push('\n');
        }
        Ok(out)
    }

    fn copy_el(&mut self, path_str: &str, pos: &Position) -> Result<Report> {
        let dpath = path::parse(path_str)?;
        if dpath.is_root() {
            bail!("cannot copy the workbook root");
        }
        let sheet_idx = self.sheet_index(&dpath.segments[0])?;

        // Copy a whole sheet.
        if dpath.segments.len() == 1 {
            let src_name = self.sheets[sheet_idx].name.clone();
            let src_part = self.sheets[sheet_idx].part.clone();
            let src_xml = self.sheets[sheet_idx].xml.clone();
            let mut n = 2;
            let mut name = format!("{src_name} ({n})");
            while self.sheets.iter().any(|s| s.name == name) {
                n += 1;
                name = format!("{src_name} ({n})");
            }
            let new_idx = self.create_sheet(&name)?;
            self.sheets[new_idx].xml = src_xml;
            // The sheet's rels (hyperlinks, drawings) must travel with it,
            // or its r:id references dangle.
            let src_rels_part =
                src_part.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels";
            if self.pkg.has_part(&src_rels_part) {
                let bytes = self.pkg.raw(&src_rels_part)?.to_vec();
                let new_rels_part = self.sheets[new_idx]
                    .part
                    .replace("xl/worksheets/", "xl/worksheets/_rels/")
                    + ".rels";
                self.pkg.put_raw(&new_rels_part, bytes);
            }
            let mut info = NodeInfo::new(format!("/{name}"), "sheet");
            info.attr("copied-from", src_name);
            return Ok(Report::Nodes(vec![info]));
        }

        // Copy a row: duplicate below the original (or at --index).
        let seg = &dpath.segments[1];
        if seg.name.eq_ignore_ascii_case("row") {
            let src_num = seg.index().context("row needs an index, e.g. row[5]")? as u32;
            let src_row = {
                let sd = self.sheets[sheet_idx]
                    .xml
                    .child("sheetData")
                    .context("worksheet has no <sheetData>")?;
                find_row(sd, src_num)
                    .with_context(|| format!("row {src_num} has no content"))?
                    .clone()
            };
            let dest_num: u32 = match pos {
                Position::Append => src_num + 1,
                Position::Index(n) => *n as u32,
                _ => bail!("copy row supports --index N (1-based destination row)"),
            };
            if dest_num == 0 {
                bail!("row numbers are 1-based");
            }
            // Make room, then materialize the clone at its new number.
            {
                let sd = self.sheets[sheet_idx].xml.child_mut("sheetData").unwrap();
                shift_rows_down(sd, dest_num);
            }
            self.rewrite_all_formulas(sheet_idx, &Shift::Rows { from: dest_num, delta: 1 });
            {
                let mut clone = src_row;
                clone.set_attr("r", &dest_num.to_string());
                for cnode in clone.children.iter_mut() {
                    if let XmlNode::Element(c) = cnode {
                        if c.local_name() == "c" {
                            if let Some((col, _)) = c.attr_local("r").and_then(parse_cell_ref_opt)
                            {
                                c.set_attr("r", &cell_name(col, dest_num - 1));
                            }
                        }
                    }
                }
                let sd = self.sheets[sheet_idx].xml.child_mut("sheetData").unwrap();
                // Insert keeping rows ordered.
                let mut at = sd.children.len();
                for (i, node) in sd.children.iter().enumerate() {
                    if let XmlNode::Element(e) = node {
                        if e.local_name() == "row" {
                            let r: u32 = e
                                .attr_local("r")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(u32::MAX);
                            if r > dest_num {
                                at = i;
                                break;
                            }
                        }
                    }
                }
                sd.children.insert(at, XmlNode::Element(clone));
            }
            let sheet_name = &self.sheets[sheet_idx].name;
            let mut info = NodeInfo::new(format!("/{sheet_name}/row[{dest_num}]"), "row");
            info.attr("copied-from", format!("/{sheet_name}/row[{src_num}]"));
            return Ok(Report::Nodes(vec![info]));
        }
        bail!("xlsx copy supports a sheet (/Sheet1) or a row (/Sheet1/row[5])")
    }

    fn screenshot(&mut self) -> Result<Vec<Vec<u8>>> {
        use crate::render::{Canvas, Color, BLACK, GRID, HEADER_BG, WHITE};
        let mut images = Vec::new();
        for sheet in &self.sheets {
            let ((c0, r0), (c1, r1)) = Self::used_range(sheet).unwrap_or(((0, 0), (7, 14)));
            // Charts anchored on the sheet extend the rendered area.
            let anchors = self.sheet_chart_anchors(sheet);
            let (mut c1, mut r1) = (c1, r1);
            for (_, _, tc, tr, _) in &anchors {
                c1 = c1.max(*tc);
                r1 = r1.max(*tr);
            }
            let font_px = 14.0;
            let row_h = 26.0f32;
            let pad = 6.0f32;
            let header_w = 44.0f32;

            // Column widths sized to content (measured with a probe canvas).
            let probe = Canvas::new(1, 1, WHITE)?;
            let mut col_w: Vec<f32> = vec![64.0; (c1 - c0 + 1) as usize];
            if let Some(sd) = sheet.xml.child("sheetData") {
                for row in sd.children_named("row") {
                    for c in row.children_named("c") {
                        if let Some((col, _)) = c.attr_local("r").and_then(parse_cell_ref_opt) {
                            if col >= c0 && col <= c1 {
                                let text = self.cell_display(c).0;
                                let w = probe.text_width(&text, font_px, false) + 2.0 * pad;
                                let slot = &mut col_w[(col - c0) as usize];
                                *slot = slot.max(w).min(280.0);
                            }
                        }
                    }
                }
            }
            let grid_w: f32 = col_w.iter().sum();
            let width = ((header_w + grid_w + 2.0).ceil() as u32).min(4000);
            let height = ((row_h * (r1 - r0 + 2) as f32 + 2.0).ceil() as u32).min(8000);
            let mut canvas = Canvas::new(width, height, WHITE)?;

            // Headers.
            canvas.fill_rect(0.0, 0.0, width as f32, row_h, HEADER_BG);
            canvas.fill_rect(0.0, 0.0, header_w, height as f32, HEADER_BG);
            let mut x = header_w;
            for (i, w) in col_w.iter().enumerate() {
                let letters = col_letters(c0 + i as u32);
                let tw = canvas.text_width(&letters, font_px, true);
                canvas.draw_text(&letters, x + (w - tw) / 2.0, row_h - 8.0, font_px, BLACK, true);
                x += w;
            }
            for r in r0..=r1 {
                let label = (r + 1).to_string();
                let y = row_h * (r - r0 + 1) as f32;
                let tw = canvas.text_width(&label, font_px, false);
                canvas.draw_text(&label, (header_w - tw) / 2.0, y + row_h - 8.0, font_px, BLACK, false);
            }

            // Cells.
            if let Some(sd) = sheet.xml.child("sheetData") {
                for row in sd.children_named("row") {
                    for c in row.children_named("c") {
                        let Some((col, r)) = c.attr_local("r").and_then(parse_cell_ref_opt) else {
                            continue;
                        };
                        if col < c0 || col > c1 || r < r0 || r > r1 {
                            continue;
                        }
                        let cx: f32 =
                            header_w + col_w[..(col - c0) as usize].iter().sum::<f32>();
                        let cy = row_h * (r - r0 + 1) as f32;
                        let cw = col_w[(col - c0) as usize];
                        let (mut bold, mut italic, mut color, mut fill) =
                            (false, false, BLACK, None::<Color>);
                        if let Some(s) = c.attr_local("s").and_then(|v| v.parse::<usize>().ok()) {
                            for (k, v) in style_props(&self.styles, s) {
                                match k.as_str() {
                                    "bold" => bold = v == "true",
                                    "italic" => italic = v == "true",
                                    "color" => color = Color::from_hex(&v).unwrap_or(BLACK),
                                    "fill" => fill = Color::from_hex(&v),
                                    _ => {}
                                }
                            }
                        }
                        if let Some(fill) = fill {
                            canvas.fill_rect(cx, cy, cw, row_h, fill);
                        }
                        let text = self.cell_display(c).0;
                        let numeric = text.parse::<f64>().is_ok() && !text.is_empty();
                        let tw = canvas.text_width(&text, font_px, bold);
                        let tx = if numeric { cx + cw - pad - tw } else { cx + pad };
                        canvas.draw_text_styled(
                            &text,
                            tx,
                            cy + row_h - 8.0,
                            font_px,
                            color,
                            bold,
                            italic,
                        );
                    }
                }
            }

            // Gridlines on top.
            let mut x = header_w;
            canvas.line(0.0, 0.0, 0.0, height as f32, GRID);
            canvas.line(x, 0.0, x, height as f32, GRID);
            for w in &col_w {
                x += w;
                canvas.line(x, 0.0, x, height as f32, GRID);
            }
            for r in 0..=(r1 - r0 + 2) {
                let y = row_h * r as f32;
                canvas.line(0.0, y, width as f32, y, GRID);
            }

            // Charts drawn over the grid at their anchor rectangles.
            for (fc, fr, tc, tr, space) in &anchors {
                let col_x = |col: u32| -> f32 {
                    let idx = (col.saturating_sub(c0) as usize).min(col_w.len());
                    header_w + col_w[..idx].iter().sum::<f32>()
                };
                let x0 = col_x(*fc);
                let x1 = col_x(*tc).max(x0 + 60.0);
                let y0 = row_h * (fr.saturating_sub(r0) + 1) as f32;
                let y1 = (row_h * (tr.saturating_sub(r0) + 1) as f32).max(y0 + 50.0);
                crate::chartdraw::draw_chart(&mut canvas, space, x0, y0, x1 - x0, y1 - y0);
            }
            images.push(canvas.png()?);
        }
        Ok(images)
    }

    fn dump(&mut self) -> Result<serde_json::Value> {
        let mut ops = Vec::new();
        for (i, sheet) in self.sheets.iter().enumerate() {
            if i == 0 {
                if sheet.name != "Sheet1" {
                    ops.push(json!({
                        "command": "set", "path": "/Sheet1",
                        "props": { "name": sheet.name },
                    }));
                }
            } else {
                ops.push(json!({
                    "command": "add", "parent": "/", "type": "sheet",
                    "props": { "name": sheet.name },
                }));
            }
            for (cell, author, text) in self.sheet_comments(sheet) {
                ops.push(json!({
                    "command": "set", "path": format!("/{}/{}", sheet.name, cell),
                    "props": { "comment": text, "author": author },
                }));
            }
            let Some(sd) = sheet.xml.child("sheetData") else { continue };
            for row in sd.children_named("row") {
                for c in row.children_named("c") {
                    let Some(r) = c.attr_local("r") else { continue };
                    let mut props = serde_json::Map::new();
                    if let Some(f) = c.child("f") {
                        props.insert("value".into(), json!(format!("={}", f.text_content())));
                    } else {
                        let (value, typ) = self.cell_display(c);
                        if value.is_empty() {
                            continue;
                        }
                        props.insert("value".into(), json!(value));
                        if typ == "string" && value.parse::<f64>().is_ok() {
                            // Preserve string-typed numerics.
                            props.insert("type".into(), json!("string"));
                        }
                    }
                    if let Some(s) = c.attr_local("s").and_then(|v| v.parse::<usize>().ok()) {
                        if s != 0 {
                            for (k, v) in style_props(&self.styles, s) {
                                props.insert(k, json!(v));
                            }
                        }
                    }
                    ops.push(json!({
                        "command": "set",
                        "path": format!("/{}/{}", sheet.name, r),
                        "props": props,
                    }));
                }
            }
        }
        Ok(serde_json::Value::Array(ops))
    }
}

// ------------------------------------------------- formula ref rewrite ----

/// A structural shift on one sheet that formula references must follow.
#[derive(Debug, Clone, Copy)]
enum Shift {
    /// Rows >= `from` (1-based row number) moved by `delta`.
    Rows { from: u32, delta: i64 },
    /// Columns >= `from` (0-based column index) moved by `delta`.
    Cols { from: u32, delta: i64 },
}

/// Rewrite A1-style references in `formula` when rows/columns of
/// `target_sheet` shift. `own_sheet` says the formula lives on the shifted
/// sheet (so unqualified refs shift too).
fn rewrite_formula_refs(
    formula: &str,
    target_sheet: &str,
    own_sheet: bool,
    shift: &Shift,
) -> String {
    // Never touch string literals: split on '"' and only process even
    // segments ("" escaping just yields extra empty segments, still even).
    let mut out = String::new();
    for (i, segment) in formula.split('"').enumerate() {
        if i > 0 {
            out.push('"');
        }
        if i % 2 == 0 {
            out.push_str(&rewrite_segment(segment, target_sheet, own_sheet, shift));
        } else {
            out.push_str(segment);
        }
    }
    out
}

/// Parse column letters into a 0-based index (inverse of `col_letters`).
fn letters_to_col(letters: &str) -> Option<u32> {
    let mut col: u64 = 0;
    for c in letters.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        col = col * 26 + (c.to_ascii_uppercase() as u64 - 'A' as u64 + 1);
    }
    if col == 0 || col > 16384 {
        None
    } else {
        Some(col as u32 - 1)
    }
}

fn rewrite_segment(
    segment: &str,
    target_sheet: &str,
    own_sheet: bool,
    shift: &Shift,
) -> String {
    // (Sheet! | 'Sheet Name'!)? $?COL $?ROW
    let re = Regex::new(r"(?:('[^']+'|[A-Za-z0-9_.]+)!)?(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)")
        .unwrap();
    let bytes = segment.as_bytes();
    let mut out = String::new();
    let mut last = 0;
    for caps in re.captures_iter(segment) {
        let m = caps.get(0).unwrap();
        out.push_str(&segment[last..m.start()]);
        last = m.end();
        let matched = m.as_str();

        // Guards: not part of a longer identifier (e.g. LOG10(...)) and not
        // a function call like A1(...).
        let prev_ok = m.start() == 0
            || !(bytes[m.start() - 1].is_ascii_alphanumeric()
                || bytes[m.start() - 1] == b'_'
                || bytes[m.start() - 1] == b'$');
        let next_ok = m.end() >= bytes.len()
            || !(bytes[m.end()].is_ascii_alphanumeric()
                || bytes[m.end()] == b'('
                || bytes[m.end()] == b'_');
        let sheet_ref = caps.get(1).map(|s| s.as_str().trim_matches('\''));
        let applies = match sheet_ref {
            Some(s) => s.eq_ignore_ascii_case(target_sheet),
            None => own_sheet,
        };
        let row: u32 = caps[5].parse().unwrap_or(0);
        let col = letters_to_col(&caps[3]);
        if !prev_ok || !next_ok || !applies || row == 0 || col.is_none() {
            out.push_str(matched);
            continue;
        }
        let col = col.unwrap();
        let (new_col, new_row) = match shift {
            Shift::Rows { from, delta } if row >= *from => {
                (col, ((row as i64) + delta).max(1) as u32)
            }
            Shift::Cols { from, delta } if col >= *from => {
                ((col as i64 + delta).max(0) as u32, row)
            }
            _ => {
                out.push_str(matched);
                continue;
            }
        };
        let sheet_part = caps
            .get(1)
            .map(|s| format!("{}!", s.as_str()))
            .unwrap_or_default();
        out.push_str(&format!(
            "{}{}{}{}{}",
            sheet_part,
            &caps[2],
            col_letters(new_col),
            &caps[4],
            new_row
        ));
    }
    out.push_str(&segment[last..]);
    out
}

#[cfg(test)]
mod tests {
    use super::{rewrite_formula_refs, Shift};

    #[test]
    fn shifts_own_sheet_refs() {
        assert_eq!(
            rewrite_formula_refs("SUM(B2:B10)+A1", "Sheet1", true, &Shift::Rows { from: 3, delta: 1 }),
            "SUM(B2:B11)+A1"
        );
        assert_eq!(
            rewrite_formula_refs("$B$5", "Sheet1", true, &Shift::Rows { from: 3, delta: -1 }),
            "$B$4"
        );
    }

    #[test]
    fn respects_sheet_qualifiers_and_literals() {
        assert_eq!(
            rewrite_formula_refs("Sheet2!A5+A5", "Sheet2", false, &Shift::Rows { from: 1, delta: 1 }),
            "Sheet2!A6+A5"
        );
        assert_eq!(
            rewrite_formula_refs("'My Sheet'!A5", "My Sheet", false, &Shift::Rows { from: 1, delta: 2 }),
            "'My Sheet'!A7"
        );
        assert_eq!(
            rewrite_formula_refs("IF(A5>0,\"A5\",LOG10(A5))", "Sheet1", true, &Shift::Rows { from: 1, delta: 1 }),
            "IF(A6>0,\"A5\",LOG10(A6))"
        );
    }

    #[test]
    fn date_serials_roundtrip() {
        use super::{parse_iso_datetime, serial_to_iso};
        let (s, d, t) = parse_iso_datetime("2026-07-10").unwrap();
        assert_eq!(s, 46213.0);
        assert!(d && !t);
        assert_eq!(serial_to_iso(s, true, false), "2026-07-10");
        let (s, _, _) = parse_iso_datetime("2026-07-10 14:30:00").unwrap();
        assert_eq!(serial_to_iso(s, true, true), "2026-07-10 14:30:00");
        let (s, d, t) = parse_iso_datetime("14:30").unwrap();
        assert!(!d && t);
        assert_eq!(serial_to_iso(s, false, true), "14:30:00");
        assert!(parse_iso_datetime("not a date").is_none());
        assert!(parse_iso_datetime("2026-13-01").is_none());
    }

    #[test]
    fn csv_quoting() {
        use super::{csv_field, parse_csv};
        let rows = parse_csv("a,\"b,c\",\"say \"\"hi\"\"\"\r\nd,,f\n");
        assert_eq!(rows, vec![
            vec!["a".to_string(), "b,c".into(), "say \"hi\"".into()],
            vec!["d".to_string(), "".into(), "f".into()],
        ]);
        assert_eq!(csv_field("plain"), "plain");
        assert_eq!(csv_field("a,b"), "\"a,b\"");
        assert_eq!(csv_field("say \"hi\""), "\"say \"\"hi\"\"\"");
    }

    #[test]
    fn shifts_columns() {
        // Insert a column before B (0-based col 1): B→C, A stays.
        assert_eq!(
            rewrite_formula_refs("SUM(B2:B10)+A1", "Sheet1", true, &Shift::Cols { from: 1, delta: 1 }),
            "SUM(C2:C10)+A1"
        );
        // Remove column B: C→B and beyond.
        assert_eq!(
            rewrite_formula_refs("C1+$D$2+Sheet2!C1", "Sheet1", true, &Shift::Cols { from: 2, delta: -1 }),
            "B1+$C$2+Sheet2!C1"
        );
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

/// Re-letter cells in columns >= `from` (0-based) by `delta` in every row.
fn shift_cols(sheet_data: &mut XmlElement, from: u32, delta: i64) {
    for node in sheet_data.children.iter_mut() {
        let XmlNode::Element(row) = node else { continue };
        if row.local_name() != "row" {
            continue;
        }
        for cnode in row.children.iter_mut() {
            let XmlNode::Element(c) = cnode else { continue };
            if c.local_name() != "c" {
                continue;
            }
            if let Some((col, r)) = c.attr_local("r").and_then(parse_cell_ref_opt) {
                if col >= from {
                    let new_col = (col as i64 + delta).max(0) as u32;
                    c.set_attr("r", &cell_name(new_col, r));
                }
            }
        }
    }
}

/// Drop every cell in the given 0-based column. Returns how many were removed.
fn remove_col_cells(sheet_data: &mut XmlElement, col: u32) -> usize {
    let mut removed = 0;
    for node in sheet_data.children.iter_mut() {
        let XmlNode::Element(row) = node else { continue };
        if row.local_name() != "row" {
            continue;
        }
        let before = row.children.len();
        row.children.retain(|n| {
            !matches!(n, XmlNode::Element(e)
                if e.local_name() == "c"
                    && e.attr_local("r").and_then(parse_cell_ref_opt).map(|(c, _)| c) == Some(col))
        });
        removed += before - row.children.len();
    }
    removed
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
