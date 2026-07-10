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
}

pub fn parse_kind(s: &str) -> Result<ChartKind> {
    match s.to_ascii_lowercase().as_str() {
        "column" | "col" | "bar-vertical" => Ok(ChartKind::Column),
        "bar" | "bar-horizontal" => Ok(ChartKind::Bar),
        "line" => Ok(ChartKind::Line),
        "pie" => Ok(ChartKind::Pie),
        other => bail!("unknown chart kind '{other}' (column/bar/line/pie)"),
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

fn build_ser(idx: usize, series: &Series, line: bool) -> XmlElement {
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

const AX_CAT: &str = "111111111";
const AX_VAL: &str = "222222222";

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

pub const CHART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
pub const CHART_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
