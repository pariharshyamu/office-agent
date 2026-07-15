//! DrawingML chart (chartSpace) builder shared by the xlsx and pptx
//! handlers. Series carry cached values plus optional `Sheet!$A$1:$A$5`
//! references; xlsx charts reference live cells, pptx charts are
//! cache-only (they render everywhere; "Edit Data" needs an embedded
//! workbook, which we do not create).

use anyhow::{bail, Result};

use crate::xml::{el, XmlElement};

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ChartKind {
    /// Vertical bars.
    Column,
    /// Horizontal bars.
    Bar,
    Line,
    Pie,
    /// XY scatter: each point is an (x, y) number pair on two value axes.
    Scatter,
}

pub fn parse_kind(s: &str) -> Result<ChartKind> {
    match s.to_ascii_lowercase().as_str() {
        "column" | "col" | "bar-vertical" => Ok(ChartKind::Column),
        "bar" | "bar-horizontal" => Ok(ChartKind::Bar),
        "line" => Ok(ChartKind::Line),
        "pie" => Ok(ChartKind::Pie),
        "scatter" | "xy" => Ok(ChartKind::Scatter),
        other => bail!("unknown chart kind '{other}' (column/bar/line/pie/scatter)"),
    }
}

#[derive(Debug, Clone, Default)]
pub struct Series {
    pub name: String,
    /// Cell reference for the series name, e.g. `Sheet1!$B$1`.
    pub name_ref: Option<String>,
    pub cats: Vec<String>,
    pub cats_ref: Option<String>,
    pub vals: Vec<f64>,
    pub vals_ref: Option<String>,
    /// Scatter-only x values (paired with `vals` as y). Empty for other
    /// kinds; scatter falls back to 1..n when missing.
    pub xs: Vec<f64>,
    pub xs_ref: Option<String>,
}

fn c_val(name: &str, val: &str) -> XmlElement {
    el(name, &[("val", val)])
}

fn str_cache(values: &[String]) -> XmlElement {
    let mut cache = XmlElement::new("c:strCache");
    cache.push(c_val("c:ptCount", &values.len().to_string()));
    for (i, v) in values.iter().enumerate() {
        let mut pt = el("c:pt", &[("idx", i.to_string().as_str())]);
        let mut cv = XmlElement::new("c:v");
        cv.push_text(v);
        pt.push(cv);
        cache.push(pt);
    }
    cache
}

fn num_cache(values: &[f64]) -> XmlElement {
    let mut cache = XmlElement::new("c:numCache");
    let mut fc = XmlElement::new("c:formatCode");
    fc.push_text("General");
    cache.push(fc);
    cache.push(c_val("c:ptCount", &values.len().to_string()));
    for (i, v) in values.iter().enumerate() {
        let mut pt = el("c:pt", &[("idx", i.to_string().as_str())]);
        let mut cv = XmlElement::new("c:v");
        cv.push_text(&format_f64(*v));
        pt.push(cv);
        cache.push(pt);
    }
    cache
}

fn format_f64(v: f64) -> String {
    if v.fract() == 0.0 && v.abs() < 1e15 {
        format!("{}", v as i64)
    } else {
        format!("{v}")
    }
}

/// `<c:strRef><c:f>REF</c:f><cache/></c:strRef>` (or bare cache when no ref).
fn str_source(reference: &Option<String>, values: &[String]) -> XmlElement {
    match reference {
        Some(r) => {
            let mut sref = XmlElement::new("c:strRef");
            let mut f = XmlElement::new("c:f");
            f.push_text(r);
            sref.push(f);
            sref.push(str_cache(values));
            sref
        }
        None => {
            // Literal data: same cache shape under strLit.
            let mut lit = str_cache(values);
            lit.name = "c:strLit".to_string();
            lit
        }
    }
}

fn num_source(reference: &Option<String>, values: &[f64]) -> XmlElement {
    match reference {
        Some(r) => {
            let mut nref = XmlElement::new("c:numRef");
            let mut f = XmlElement::new("c:f");
            f.push_text(r);
            nref.push(f);
            nref.push(num_cache(values));
            nref
        }
        None => {
            let mut lit = num_cache(values);
            lit.name = "c:numLit".to_string();
            // numLit has no formatCode child in the schema.
            lit.children.retain(|n| {
                !matches!(n, crate::xml::XmlNode::Element(e) if e.local_name() == "formatCode")
            });
            lit
        }
    }
}

/// `c:ser` with idx/order/tx filled in — the part every chart kind shares.
fn ser_header(idx: usize, series: &Series) -> XmlElement {
    let mut ser = XmlElement::new("c:ser");
    ser.push(c_val("c:idx", &idx.to_string()));
    ser.push(c_val("c:order", &idx.to_string()));
    let mut tx = XmlElement::new("c:tx");
    match &series.name_ref {
        Some(r) => {
            let mut sref = XmlElement::new("c:strRef");
            let mut f = XmlElement::new("c:f");
            f.push_text(r);
            sref.push(f);
            sref.push(str_cache(std::slice::from_ref(&series.name)));
            tx.push(sref);
        }
        None => {
            let mut v = XmlElement::new("c:v");
            v.push_text(&series.name);
            tx.push(v);
        }
    }
    ser.push(tx);
    ser
}

fn build_ser(idx: usize, series: &Series, line: bool) -> XmlElement {
    let mut ser = ser_header(idx, series);
    if !series.cats.is_empty() {
        let mut cat = XmlElement::new("c:cat");
        cat.push(str_source(&series.cats_ref, &series.cats));
        ser.push(cat);
    }
    let mut val = XmlElement::new("c:val");
    val.push(num_source(&series.vals_ref, &series.vals));
    ser.push(val);
    if line {
        ser.push(c_val("c:smooth", "0"));
    }
    ser
}

/// Scatter series: marker-only points (no connecting line) with paired
/// xVal/yVal number sources.
fn build_scatter_ser(idx: usize, series: &Series) -> XmlElement {
    let mut ser = ser_header(idx, series);
    let mut sp = XmlElement::new("c:spPr");
    let mut ln = XmlElement::new("a:ln");
    ln.push(XmlElement::new("a:noFill"));
    sp.push(ln);
    ser.push(sp);
    let xs: Vec<f64> = if series.xs.len() == series.vals.len() && !series.xs.is_empty() {
        series.xs.clone()
    } else {
        (1..=series.vals.len()).map(|i| i as f64).collect()
    };
    let mut xv = XmlElement::new("c:xVal");
    xv.push(num_source(&series.xs_ref, &xs));
    ser.push(xv);
    let mut yv = XmlElement::new("c:yVal");
    yv.push(num_source(&series.vals_ref, &series.vals));
    ser.push(yv);
    ser.push(c_val("c:smooth", "0"));
    ser
}

const AX_CAT: &str = "111111111";
const AX_VAL: &str = "222222222";

/// Scatter plots put value axes on both sides.
fn build_scatter_axes() -> (XmlElement, XmlElement) {
    let mk = |id: &str, pos: &str, cross: &str| {
        let mut ax = XmlElement::new("c:valAx");
        ax.push(c_val("c:axId", id));
        let mut scaling = XmlElement::new("c:scaling");
        scaling.push(c_val("c:orientation", "minMax"));
        ax.push(scaling);
        ax.push(c_val("c:delete", "0"));
        ax.push(c_val("c:axPos", pos));
        ax.push(c_val("c:crossAx", cross));
        ax
    };
    (mk(AX_CAT, "b", AX_VAL), mk(AX_VAL, "l", AX_CAT))
}

fn build_axes(kind: ChartKind) -> (XmlElement, XmlElement) {
    // Horizontal bar charts flip the axis positions.
    let (cat_pos, val_pos) = if kind == ChartKind::Bar { ("l", "b") } else { ("b", "l") };
    let mut cat_ax = XmlElement::new("c:catAx");
    cat_ax.push(c_val("c:axId", AX_CAT));
    let mut scaling = XmlElement::new("c:scaling");
    scaling.push(c_val("c:orientation", "minMax"));
    cat_ax.push(scaling);
    cat_ax.push(c_val("c:delete", "0"));
    cat_ax.push(c_val("c:axPos", cat_pos));
    cat_ax.push(c_val("c:crossAx", AX_VAL));

    let mut val_ax = XmlElement::new("c:valAx");
    val_ax.push(c_val("c:axId", AX_VAL));
    let mut scaling = XmlElement::new("c:scaling");
    scaling.push(c_val("c:orientation", "minMax"));
    val_ax.push(scaling);
    val_ax.push(c_val("c:delete", "0"));
    val_ax.push(c_val("c:axPos", val_pos));
    val_ax.push(c_val("c:crossAx", AX_CAT));
    (cat_ax, val_ax)
}

/// Build a complete `c:chartSpace` root.
pub fn build_chart_space(
    kind: ChartKind,
    title: Option<&str>,
    series: &[Series],
) -> Result<XmlElement> {
    if series.is_empty() {
        bail!("a chart needs at least one data series");
    }
    let mut space = el(
        "c:chartSpace",
        &[
            ("xmlns:c", "http://schemas.openxmlformats.org/drawingml/2006/chart"),
            ("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
            (
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            ),
        ],
    );
    let mut chart = XmlElement::new("c:chart");

    if let Some(title) = title {
        let mut t = XmlElement::new("c:title");
        let mut tx = XmlElement::new("c:tx");
        let mut rich = XmlElement::new("c:rich");
        rich.push(XmlElement::new("a:bodyPr"));
        rich.push(XmlElement::new("a:lstStyle"));
        let mut p = XmlElement::new("a:p");
        let mut r = XmlElement::new("a:r");
        let mut at = XmlElement::new("a:t");
        at.push_text(title);
        r.push(at);
        p.push(r);
        rich.push(p);
        tx.push(rich);
        t.push(tx);
        t.push(c_val("c:overlay", "0"));
        chart.push(t);
        chart.push(c_val("c:autoTitleDeleted", "0"));
    }

    let mut plot_area = XmlElement::new("c:plotArea");
    plot_area.push(XmlElement::new("c:layout"));

    match kind {
        ChartKind::Column | ChartKind::Bar => {
            let mut bar = XmlElement::new("c:barChart");
            bar.push(c_val("c:barDir", if kind == ChartKind::Bar { "bar" } else { "col" }));
            bar.push(c_val("c:grouping", "clustered"));
            bar.push(c_val("c:varyColors", "0"));
            for (i, s) in series.iter().enumerate() {
                bar.push(build_ser(i, s, false));
            }
            bar.push(c_val("c:gapWidth", "150"));
            bar.push(c_val("c:axId", AX_CAT));
            bar.push(c_val("c:axId", AX_VAL));
            plot_area.push(bar);
            let (cat_ax, val_ax) = build_axes(kind);
            plot_area.push(cat_ax);
            plot_area.push(val_ax);
        }
        ChartKind::Line => {
            let mut line = XmlElement::new("c:lineChart");
            line.push(c_val("c:grouping", "standard"));
            line.push(c_val("c:varyColors", "0"));
            for (i, s) in series.iter().enumerate() {
                line.push(build_ser(i, s, true));
            }
            line.push(c_val("c:marker", "1"));
            line.push(c_val("c:axId", AX_CAT));
            line.push(c_val("c:axId", AX_VAL));
            plot_area.push(line);
            let (cat_ax, val_ax) = build_axes(kind);
            plot_area.push(cat_ax);
            plot_area.push(val_ax);
        }
        ChartKind::Pie => {
            let mut pie = XmlElement::new("c:pieChart");
            pie.push(c_val("c:varyColors", "1"));
            for (i, s) in series.iter().enumerate().take(1) {
                pie.push(build_ser(i, s, false));
            }
            pie.push(c_val("c:firstSliceAng", "0"));
            plot_area.push(pie);
        }
        ChartKind::Scatter => {
            let mut sc = XmlElement::new("c:scatterChart");
            sc.push(c_val("c:scatterStyle", "lineMarker"));
            sc.push(c_val("c:varyColors", "0"));
            for (i, s) in series.iter().enumerate() {
                sc.push(build_scatter_ser(i, s));
            }
            sc.push(c_val("c:axId", AX_CAT));
            sc.push(c_val("c:axId", AX_VAL));
            plot_area.push(sc);
            let (x_ax, y_ax) = build_scatter_axes();
            plot_area.push(x_ax);
            plot_area.push(y_ax);
        }
    }
    chart.push(plot_area);

    if series.len() > 1 || kind == ChartKind::Pie {
        let mut legend = XmlElement::new("c:legend");
        legend.push(c_val("c:legendPos", "r"));
        legend.push(c_val("c:overlay", "0"));
        chart.push(legend);
    }
    chart.push(c_val("c:plotVisOnly", "1"));
    space.push(chart);
    Ok(space)
}

/// 0-based column index to letters (0 = A).
fn column_letters(mut c: u32) -> String {
    let mut s = String::new();
    loop {
        s.insert(0, (b'A' + (c % 26) as u8) as char);
        if c < 26 {
            break;
        }
        c = c / 26 - 1;
    }
    s
}

fn wb_cell_str(r: &str, text: &str) -> XmlElement {
    let mut c = el("c", &[("r", r), ("t", "inlineStr")]);
    let mut is = XmlElement::new("is");
    let mut t = XmlElement::new("t");
    t.push_text(text);
    is.push(t);
    c.push(is);
    c
}

fn wb_cell_num(r: &str, v: f64) -> XmlElement {
    let mut c = el("c", &[("r", r)]);
    let mut vv = XmlElement::new("v");
    vv.push_text(&format_f64(v));
    c.push(vv);
    c
}

/// Build a real workbook holding the chart data (categories/x in column A,
/// series in B..; names in row 1) and point every series' refs at it, so
/// Office's "Edit Data" opens a live sheet. Returns the xlsx zip bytes;
/// callers embed them and wire a `c:externalData` rel to the chart part.
pub fn embedded_workbook(series: &mut [Series]) -> Result<Vec<u8>> {
    let mut pkg = crate::templates::blank_package(crate::pkg::DocKind::Xlsx);
    let mut sheet = pkg.xml("xl/worksheets/sheet1.xml")?;
    let nrows = series.iter().map(|s| s.vals.len()).max().unwrap_or(0);
    let scatter = series.iter().any(|s| !s.xs.is_empty());
    {
        let Some(sd) = sheet.child_mut("sheetData") else {
            bail!("workbook template is missing sheetData");
        };
        let mut hdr = el("row", &[("r", "1")]);
        for (j, s) in series.iter().enumerate() {
            hdr.push(wb_cell_str(
                &format!("{}1", column_letters(j as u32 + 1)),
                &s.name,
            ));
        }
        sd.push(hdr);
        for i in 0..nrows {
            let rn = i + 2;
            let mut row = el("row", &[("r", rn.to_string().as_str())]);
            if scatter {
                if let Some(x) = series.first().and_then(|s| s.xs.get(i)) {
                    row.push(wb_cell_num(&format!("A{rn}"), *x));
                }
            } else if let Some(c) = series.first().and_then(|s| s.cats.get(i)) {
                if !c.is_empty() {
                    row.push(wb_cell_str(&format!("A{rn}"), c));
                }
            }
            for (j, s) in series.iter().enumerate() {
                if let Some(v) = s.vals.get(i) {
                    row.push(wb_cell_num(
                        &format!("{}{rn}", column_letters(j as u32 + 1)),
                        *v,
                    ));
                }
            }
            sd.push(row);
        }
    }
    pkg.put_xml("xl/worksheets/sheet1.xml", &sheet)?;

    let last = nrows + 1;
    for (j, s) in series.iter_mut().enumerate() {
        let col = column_letters(j as u32 + 1);
        s.name_ref = Some(format!("Sheet1!${col}$1"));
        s.vals_ref = Some(format!("Sheet1!${col}$2:${col}${last}"));
        if scatter {
            s.xs_ref = Some(format!("Sheet1!$A$2:$A${last}"));
        } else if !s.cats.is_empty() {
            s.cats_ref = Some(format!("Sheet1!$A$2:$A${last}"));
        }
    }
    pkg.to_zip_bytes()
}

/// Append `<c:externalData r:id/><c:autoUpdate 0/>` to a chartSpace,
/// linking it to an embedded workbook relationship.
pub fn attach_external_data(space: &mut XmlElement, rid: &str) {
    let mut ext = el("c:externalData", &[("r:id", rid)]);
    ext.push(c_val("c:autoUpdate", "0"));
    space.push(ext);
}

/// Rels part content for a chart that references an embedded workbook.
pub fn chart_rels_xml(embedding_target: &str) -> XmlElement {
    let mut rels = el(
        "Relationships",
        &[("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")],
    );
    rels.push(el(
        "Relationship",
        &[
            ("Id", "rId1"),
            (
                "Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
            ),
            ("Target", embedding_target),
        ],
    ));
    rels
}

pub const XLSX_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/// Parse inline chart data from props: `categories=A,B,C`, `values=1,2,3`,
/// `series=Name`, plus `values2=`/`series2=`... for more series, and
/// `xvalues=` (or numeric categories) for scatter. Shared by the pptx and
/// docx handlers, which have no cell range to reference.
pub fn inline_series(kind: ChartKind, props: &crate::props::Props) -> Result<Vec<Series>> {
    let cats: Vec<String> = props
        .get("categories")
        .or_else(|| props.get("cats"))
        .map(|c| c.split(',').map(|s| s.trim().to_string()).collect())
        .unwrap_or_default();
    let parse_vals = |raw: &str| -> Result<Vec<f64>> {
        raw.split(',')
            .map(|s| {
                s.trim()
                    .parse::<f64>()
                    .map_err(|_| anyhow::anyhow!("'{}' is not a number in values", s.trim()))
            })
            .collect()
    };
    // Scatter charts read categories as numeric x values (or --prop
    // xvalues=; plain x= stays the frame position).
    let scatter = kind == ChartKind::Scatter;
    let xs: Vec<f64> = if scatter {
        props
            .get("xvalues")
            .map(parse_vals)
            .transpose()?
            .unwrap_or_else(|| {
                cats.iter()
                    .enumerate()
                    .map(|(i, c)| c.parse::<f64>().unwrap_or((i + 1) as f64))
                    .collect()
            })
    } else {
        Vec::new()
    };
    let cats = if scatter { Vec::new() } else { cats };
    let mut series = Vec::new();
    let first = props
        .get("values")
        .ok_or_else(|| anyhow::anyhow!(
            "chart needs --prop values=10,20,30 (and usually --prop categories=A,B,C)"
        ))?;
    series.push(Series {
        name: props.get("series").unwrap_or("Series 1").to_string(),
        cats: cats.clone(),
        vals: parse_vals(first)?,
        xs: xs.clone(),
        ..Default::default()
    });
    let mut n = 2;
    while let Some(vals) = props.get(&format!("values{n}")) {
        series.push(Series {
            name: props
                .get(&format!("series{n}"))
                .map(|s| s.to_string())
                .unwrap_or_else(|| format!("Series {n}")),
            cats: cats.clone(),
            vals: parse_vals(vals)?,
            xs: xs.clone(),
            ..Default::default()
        });
        n += 1;
    }
    Ok(series)
}

pub const CHART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
pub const CHART_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
