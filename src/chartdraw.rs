//! Rasterizes a `c:chartSpace` onto a render Canvas for `view screenshot`.
//! Reads the cached data (numCache/strCache/numLit/strLit) so it works the
//! same for xlsx, docx, and pptx charts — no live cell evaluation needed.
//! Like the rest of the screenshot renderer, this is an approximation for
//! agent render-look-fix loops, not a print-accurate chart engine.

use crate::render::{Canvas, Color, BLACK, GRID, MUTED, WHITE};
use crate::xml::XmlElement;

/// Office default accent color cycle.
const PALETTE: [Color; 6] = [
    Color { r: 0x44, g: 0x72, b: 0xC4 },
    Color { r: 0xED, g: 0x7D, b: 0x31 },
    Color { r: 0xA5, g: 0xA5, b: 0xA5 },
    Color { r: 0xFF, g: 0xC0, b: 0x00 },
    Color { r: 0x5B, g: 0x9B, b: 0xD5 },
    Color { r: 0x70, g: 0xAD, b: 0x47 },
];

fn palette(i: usize) -> Color {
    PALETTE[i % PALETTE.len()]
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum Kind {
    Column,
    Bar,
    Line,
    Pie,
    Scatter,
}

struct Ser {
    name: String,
    /// Scatter x values (empty otherwise).
    xs: Vec<f64>,
    vals: Vec<f64>,
}

struct Data {
    kind: Kind,
    title: Option<String>,
    cats: Vec<String>,
    series: Vec<Ser>,
}

/// Point values from a data source holder (c:cat / c:val / c:xVal / c:yVal):
/// descend into the first *Cache or *Lit and read pt/v texts in idx order.
fn cache_values(holder: &XmlElement) -> Vec<String> {
    fn find_cache<'a>(e: &'a XmlElement) -> Option<&'a XmlElement> {
        let n = e.local_name();
        if n == "strCache" || n == "numCache" || n == "strLit" || n == "numLit" {
            return Some(e);
        }
        e.children
            .iter()
            .filter_map(|c| c.as_element())
            .find_map(find_cache)
    }
    let Some(cache) = find_cache(holder) else {
        return Vec::new();
    };
    let mut pts: Vec<(usize, String)> = cache
        .children_named("pt")
        .into_iter()
        .filter_map(|pt| {
            let idx = pt.attr_local("idx")?.parse::<usize>().ok()?;
            let v = pt.child("v").map(|v| v.text_content())?;
            Some((idx, v))
        })
        .collect();
    pts.sort_by_key(|(i, _)| *i);
    // Fill gaps (sparse caches) with empty strings.
    let max = pts.last().map(|(i, _)| *i + 1).unwrap_or(0);
    let mut out = vec![String::new(); max];
    for (i, v) in pts {
        out[i] = v;
    }
    out
}

fn nums(holder: &XmlElement) -> Vec<f64> {
    cache_values(holder)
        .iter()
        .map(|s| s.parse::<f64>().unwrap_or(0.0))
        .collect()
}

/// Series name: cached strRef value or literal c:v.
fn ser_name(ser: &XmlElement, idx: usize) -> String {
    let from_tx = ser.child("tx").and_then(|tx| {
        if tx.child("strRef").is_some() {
            cache_values(tx).into_iter().next()
        } else {
            tx.child("v").map(|v| v.text_content())
        }
    });
    from_tx
        .filter(|s| !s.is_empty())
        .unwrap_or_else(|| format!("Series {}", idx + 1))
}

fn extract(space: &XmlElement) -> Option<Data> {
    let chart = space.child("chart")?;
    let title = chart
        .child("title")
        .map(|t| t.child("tx").map(|tx| tx.text_content()).unwrap_or_default())
        .filter(|s| !s.trim().is_empty());
    let plot = chart.child("plotArea")?;
    let (pc, kind) = [
        ("barChart", Kind::Column),
        ("bar3DChart", Kind::Column),
        ("lineChart", Kind::Line),
        ("line3DChart", Kind::Line),
        ("areaChart", Kind::Line),
        ("pieChart", Kind::Pie),
        ("pie3DChart", Kind::Pie),
        ("doughnutChart", Kind::Pie),
        ("scatterChart", Kind::Scatter),
    ]
    .iter()
    .find_map(|(name, kind)| plot.child(name).map(|pc| (pc, *kind)))?;
    let kind = if pc.local_name().starts_with("bar")
        && pc.child("barDir").and_then(|d| d.attr_local("val")) == Some("bar")
    {
        Kind::Bar
    } else {
        kind
    };

    let mut cats: Vec<String> = Vec::new();
    let mut series = Vec::new();
    for (i, ser) in pc.children_named("ser").into_iter().enumerate() {
        let vals = ser
            .child("val")
            .or_else(|| ser.child("yVal"))
            .map(nums)
            .unwrap_or_default();
        let xs = ser.child("xVal").map(nums).unwrap_or_default();
        if let Some(cat) = ser.child("cat") {
            let c = cache_values(cat);
            if c.len() > cats.len() {
                cats = c;
            }
        }
        series.push(Ser { name: ser_name(ser, i), xs, vals });
    }
    if series.iter().all(|s| s.vals.is_empty()) {
        return None;
    }
    Some(Data { kind, title, cats, series })
}

fn fmt_num(v: f64) -> String {
    if v == 0.0 {
        return "0".into();
    }
    if v.fract() == 0.0 && v.abs() < 1e12 {
        format!("{}", v as i64)
    } else if v.abs() >= 10.0 {
        format!("{v:.1}")
    } else {
        format!("{v:.2}")
    }
}

/// Draw the chart into the given rectangle. Returns false when the
/// chartSpace has no drawable cached data (caller keeps its placeholder).
pub fn draw_chart(canvas: &mut Canvas, space: &XmlElement, x: f32, y: f32, w: f32, h: f32) -> bool {
    let Some(data) = extract(space) else {
        return false;
    };
    if w < 60.0 || h < 50.0 {
        return false;
    }
    canvas.fill_rect(x, y, w, h, WHITE);
    canvas.stroke_rect(x, y, w, h, GRID);

    let mut top = y + 6.0;
    let mut bottom = y + h - 6.0;
    if let Some(t) = &data.title {
        let size = (h * 0.07).clamp(10.0, 15.0);
        let tw = canvas.text_width(t, size, true);
        canvas.draw_text(t, x + (w - tw) / 2.0, top + canvas.ascent(size, true), size, BLACK, true);
        top += size * 1.6;
    }

    // Legend along the bottom: series names, or category names for pies.
    let legend: Vec<(String, Color)> = if data.kind == Kind::Pie {
        data.cats
            .iter()
            .enumerate()
            .map(|(i, c)| (c.clone(), palette(i)))
            .collect()
    } else if data.series.len() > 1 {
        data.series
            .iter()
            .enumerate()
            .map(|(i, s)| (s.name.clone(), palette(i)))
            .collect()
    } else {
        Vec::new()
    };
    if !legend.is_empty() {
        let size = 9.0;
        let total: f32 = legend
            .iter()
            .map(|(n, _)| 12.0 + canvas.text_width(n, size, false) + 12.0)
            .sum();
        let mut lx = (x + (w - total) / 2.0).max(x + 4.0);
        let ly = bottom - 9.0;
        for (name, color) in &legend {
            canvas.fill_rect(lx, ly, 8.0, 8.0, *color);
            canvas.draw_text(name, lx + 12.0, ly + 8.0, size, BLACK, false);
            lx += 12.0 + canvas.text_width(name, size, false) + 12.0;
        }
        bottom -= 18.0;
    }

    match data.kind {
        Kind::Pie => draw_pie(canvas, &data, x, top, w, bottom - top),
        Kind::Scatter => draw_scatter(canvas, &data, x, top, w, bottom - top),
        Kind::Bar => draw_bars(canvas, &data, x, top, w, bottom - top, true),
        Kind::Column | Kind::Line => {
            draw_cat_value(canvas, &data, x, top, w, bottom - top)
        }
    }
    true
}

/// Shared value-axis scale: min/max across all series, always spanning 0
/// for bars/columns/lines.
fn value_range(data: &Data) -> (f64, f64) {
    let mut lo = 0.0f64;
    let mut hi = 0.0f64;
    for s in &data.series {
        for v in &s.vals {
            lo = lo.min(*v);
            hi = hi.max(*v);
        }
    }
    if hi == lo {
        hi = lo + 1.0;
    }
    (lo, hi)
}

fn n_points(data: &Data) -> usize {
    data.series
        .iter()
        .map(|s| s.vals.len())
        .max()
        .unwrap_or(0)
        .max(data.cats.len())
        .max(1)
}

/// Category label, falling back to the 1-based index.
fn cat_label(data: &Data, i: usize) -> String {
    data.cats
        .get(i)
        .filter(|c| !c.is_empty())
        .cloned()
        .unwrap_or_else(|| (i + 1).to_string())
}

/// Plot rect with room for y-axis labels (left) and category labels
/// (bottom); draws the axes, gridlines, and value labels.
fn plot_frame(
    canvas: &mut Canvas,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    lo: f64,
    hi: f64,
) -> (f32, f32, f32, f32) {
    let label_w = canvas
        .text_width(&fmt_num(hi), 9.0, false)
        .max(canvas.text_width(&fmt_num(lo), 9.0, false));
    let px = x + 8.0 + label_w + 4.0;
    let py = y + 4.0;
    let pw = (x + w - 8.0) - px;
    let ph = (y + h - 14.0) - py;
    for step in 0..=4 {
        let v = lo + (hi - lo) * step as f64 / 4.0;
        let gy = py + ph - (ph * step as f32 / 4.0);
        canvas.line(px, gy, px + pw, gy, Color { r: 0xE0, g: 0xE0, b: 0xE0 });
        let label = fmt_num(v);
        let tw = canvas.text_width(&label, 9.0, false);
        canvas.draw_text(&label, px - 4.0 - tw, gy + 3.0, 9.0, MUTED, false);
    }
    canvas.line(px, py, px, py + ph, GRID);
    canvas.line(px, py + ph, px + pw, py + ph, GRID);
    (px, py, pw, ph)
}

/// Columns and lines share the category x-layout.
fn draw_cat_value(canvas: &mut Canvas, data: &Data, x: f32, y: f32, w: f32, h: f32) {
    let (lo, hi) = value_range(data);
    let (px, py, pw, ph) = plot_frame(canvas, x, y, w, h, lo, hi);
    let n = n_points(data);
    let group_w = pw / n as f32;
    let y_of = |v: f64| py + ph - ((v - lo) / (hi - lo)) as f32 * ph;
    let zero_y = y_of(0.0);

    let line_kind = data
        .series
        .iter()
        .any(|s| !s.xs.is_empty())
        .then_some(Kind::Scatter);
    let _ = line_kind;
    let is_line = data.kind == Kind::Line;
    if is_line {
        for (si, s) in data.series.iter().enumerate() {
            let color = palette(si);
            let mut prev: Option<(f32, f32)> = None;
            for (i, v) in s.vals.iter().enumerate() {
                let cx = px + group_w * (i as f32 + 0.5);
                let cy = y_of(*v);
                if let Some((lx, ly)) = prev {
                    canvas.thick_line(lx, ly, cx, cy, 2.0, color);
                }
                canvas.fill_circle(cx, cy, 2.5, color);
                prev = Some((cx, cy));
            }
        }
    } else {
        let ns = data.series.len().max(1);
        let bar_w = (group_w * 0.72 / ns as f32).max(1.0);
        for (si, s) in data.series.iter().enumerate() {
            let color = palette(si);
            for (i, v) in s.vals.iter().enumerate() {
                let bx = px + group_w * i as f32 + group_w * 0.14 + bar_w * si as f32;
                let vy = y_of(*v);
                let (top, bh) = if vy <= zero_y {
                    (vy, zero_y - vy)
                } else {
                    (zero_y, vy - zero_y)
                };
                canvas.fill_rect(bx, top, bar_w - 1.0, bh.max(1.0), color);
            }
        }
    }

    // Category labels: skip every k-th when they cannot fit.
    let every = (1..=n)
        .find(|k| {
            (0..n).step_by(*k).all(|i| {
                canvas.text_width(&cat_label(data, i), 9.0, false) < group_w * *k as f32 - 4.0
            })
        })
        .unwrap_or(n);
    for i in (0..n).step_by(every) {
        let label = cat_label(data, i);
        let tw = canvas.text_width(&label, 9.0, false);
        let cx = px + group_w * (i as f32 + 0.5);
        canvas.draw_text(&label, cx - tw / 2.0, py + ph + 11.0, 9.0, MUTED, false);
    }
}

/// Horizontal bars: categories run down the left side.
fn draw_bars(canvas: &mut Canvas, data: &Data, x: f32, y: f32, w: f32, h: f32, _horiz: bool) {
    let (lo, hi) = value_range(data);
    let n = n_points(data);
    let label_w = (0..n)
        .map(|i| canvas.text_width(&cat_label(data, i), 9.0, false))
        .fold(0.0f32, f32::max)
        .min(w * 0.3);
    let px = x + 8.0 + label_w + 4.0;
    let py = y + 4.0;
    let pw = (x + w - 8.0) - px;
    let ph = (y + h - 14.0) - py;
    canvas.line(px, py, px, py + ph, GRID);
    canvas.line(px, py + ph, px + pw, py + ph, GRID);
    let x_of = |v: f64| px + ((v - lo) / (hi - lo)) as f32 * pw;
    let zero_x = x_of(0.0);
    let group_h = ph / n as f32;
    let ns = data.series.len().max(1);
    let bar_h = (group_h * 0.72 / ns as f32).max(1.0);
    for (si, s) in data.series.iter().enumerate() {
        let color = palette(si);
        for (i, v) in s.vals.iter().enumerate() {
            let by = py + group_h * i as f32 + group_h * 0.14 + bar_h * si as f32;
            let vx = x_of(*v);
            let (left, bw) = if vx >= zero_x {
                (zero_x, vx - zero_x)
            } else {
                (vx, zero_x - vx)
            };
            canvas.fill_rect(left, by, bw.max(1.0), bar_h - 1.0, color);
        }
    }
    for i in 0..n {
        let label = cat_label(data, i);
        let tw = canvas.text_width(&label, 9.0, false).min(label_w);
        let cy = py + group_h * (i as f32 + 0.5);
        canvas.draw_text(&label, px - 4.0 - tw, cy + 3.0, 9.0, MUTED, false);
    }
    // Value labels on the x axis: min and max.
    for (v, align_right) in [(lo, false), (hi, true)] {
        let label = fmt_num(v);
        let tw = canvas.text_width(&label, 9.0, false);
        let lx = if align_right { px + pw - tw } else { px };
        canvas.draw_text(&label, lx, py + ph + 11.0, 9.0, MUTED, false);
    }
}

fn draw_pie(canvas: &mut Canvas, data: &Data, x: f32, y: f32, w: f32, h: f32) {
    let Some(s) = data.series.first() else { return };
    let total: f64 = s.vals.iter().filter(|v| **v > 0.0).sum();
    if total <= 0.0 {
        return;
    }
    let cx = x + w / 2.0;
    let cy = y + h / 2.0;
    let r = (w.min(h) / 2.0 - 8.0).max(8.0);
    let mut angle = -std::f32::consts::FRAC_PI_2;
    for (i, v) in s.vals.iter().enumerate() {
        if *v <= 0.0 {
            continue;
        }
        let sweep = (*v / total) as f32 * std::f32::consts::TAU;
        let steps = ((sweep / 0.08).ceil() as usize).max(2);
        let mut pts = vec![(cx, cy)];
        for k in 0..=steps {
            let a = angle + sweep * k as f32 / steps as f32;
            pts.push((cx + r * a.cos(), cy + r * a.sin()));
        }
        canvas.fill_polygon(&pts, palette(i));
        angle += sweep;
    }
}

fn draw_scatter(canvas: &mut Canvas, data: &Data, x: f32, y: f32, w: f32, h: f32) {
    let mut xlo = f64::INFINITY;
    let mut xhi = f64::NEG_INFINITY;
    let mut ylo = f64::INFINITY;
    let mut yhi = f64::NEG_INFINITY;
    for s in &data.series {
        for (i, v) in s.vals.iter().enumerate() {
            let sx = s.xs.get(i).copied().unwrap_or((i + 1) as f64);
            xlo = xlo.min(sx);
            xhi = xhi.max(sx);
            ylo = ylo.min(*v);
            yhi = yhi.max(*v);
        }
    }
    if !xlo.is_finite() {
        return;
    }
    if xhi == xlo {
        xhi = xlo + 1.0;
    }
    if yhi == ylo {
        yhi = ylo + 1.0;
    }
    // 5% padding so edge points stay visible.
    let (xpad, ypad) = ((xhi - xlo) * 0.05, (yhi - ylo) * 0.05);
    let (xlo, xhi, ylo, yhi) = (xlo - xpad, xhi + xpad, ylo - ypad, yhi + ypad);
    let (px, py, pw, ph) = plot_frame(canvas, x, y, w, h, ylo, yhi);
    for (si, s) in data.series.iter().enumerate() {
        let color = palette(si);
        for (i, v) in s.vals.iter().enumerate() {
            let sx = s.xs.get(i).copied().unwrap_or((i + 1) as f64);
            let gx = px + ((sx - xlo) / (xhi - xlo)) as f32 * pw;
            let gy = py + ph - ((v - ylo) / (yhi - ylo)) as f32 * ph;
            canvas.fill_circle(gx, gy, 3.0, color);
        }
    }
    // X-axis min/max labels.
    for (v, right) in [(xlo, false), (xhi, true)] {
        let label = fmt_num(v);
        let tw = canvas.text_width(&label, 9.0, false);
        let lx = if right { px + pw - tw } else { px };
        canvas.draw_text(&label, lx, py + ph + 11.0, 9.0, MUTED, false);
    }
}
