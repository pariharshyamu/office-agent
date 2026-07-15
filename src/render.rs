//! PNG rendering for `view screenshot`: a small canvas over tiny-skia with
//! text drawing via ab_glyph and the embedded DejaVu Sans fonts.
//!
//! The output is an approximation meant for agent render-look-fix loops:
//! layout, colors, and text are faithful enough to spot problems; it is not
//! a print-accurate Office renderer.

use ab_glyph::{Font, FontRef, PxScale, ScaleFont};
use anyhow::{Context, Result};
use tiny_skia::{Paint, PathBuilder, Pixmap, PixmapPaint, Rect, Stroke, Transform};

static FONT_REGULAR: &[u8] = include_bytes!("../assets/DejaVuSans.ttf");
static FONT_BOLD: &[u8] = include_bytes!("../assets/DejaVuSans-Bold.ttf");

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Color {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

pub const BLACK: Color = Color { r: 0x1A, g: 0x1A, b: 0x1A };
pub const WHITE: Color = Color { r: 0xFF, g: 0xFF, b: 0xFF };
pub const GRID: Color = Color { r: 0x99, g: 0x99, b: 0x99 };
pub const HEADER_BG: Color = Color { r: 0xF0, g: 0xF0, b: 0xF0 };
pub const MUTED: Color = Color { r: 0x88, g: 0x88, b: 0x88 };

impl Color {
    /// From a 6-hex-digit RRGGBB string (as stored in OOXML).
    pub fn from_hex(hex: &str) -> Option<Color> {
        let hex = hex.trim_start_matches('#');
        if hex.len() != 6 {
            return None;
        }
        let n = u32::from_str_radix(hex, 16).ok()?;
        Some(Color {
            r: (n >> 16) as u8,
            g: (n >> 8) as u8,
            b: n as u8,
        })
    }

    fn skia(&self) -> tiny_skia::Color {
        tiny_skia::Color::from_rgba8(self.r, self.g, self.b, 255)
    }
}

/// A run of formatted text for line layout.
#[derive(Debug, Clone)]
pub struct Span {
    pub text: String,
    /// Font size in pixels.
    pub size: f32,
    pub color: Color,
    pub bold: bool,
    /// Rendered as a synthetic oblique (sheared upright glyphs).
    pub italic: bool,
}

pub struct Canvas {
    pixmap: Pixmap,
    regular: FontRef<'static>,
    bold: FontRef<'static>,
}

impl Canvas {
    pub fn new(width: u32, height: u32, background: Color) -> Result<Canvas> {
        let mut pixmap = Pixmap::new(width.max(1), height.max(1))
            .context("cannot allocate render canvas")?;
        pixmap.fill(background.skia());
        Ok(Canvas {
            pixmap,
            regular: FontRef::try_from_slice(FONT_REGULAR).context("embedded font is invalid")?,
            bold: FontRef::try_from_slice(FONT_BOLD).context("embedded bold font is invalid")?,
        })
    }

    pub fn fill_rect(&mut self, x: f32, y: f32, w: f32, h: f32, color: Color) {
        let Some(rect) = Rect::from_xywh(x, y, w.max(0.1), h.max(0.1)) else {
            return;
        };
        let mut paint = Paint::default();
        paint.set_color(color.skia());
        paint.anti_alias = false;
        self.pixmap
            .fill_rect(rect, &paint, Transform::identity(), None);
    }

    pub fn stroke_rect(&mut self, x: f32, y: f32, w: f32, h: f32, color: Color) {
        let Some(rect) = Rect::from_xywh(x, y, w.max(0.1), h.max(0.1)) else {
            return;
        };
        let path = PathBuilder::from_rect(rect);
        let mut paint = Paint::default();
        paint.set_color(color.skia());
        paint.anti_alias = false;
        self.pixmap.stroke_path(
            &path,
            &paint,
            &Stroke { width: 1.0, ..Stroke::default() },
            Transform::identity(),
            None,
        );
    }

    pub fn line(&mut self, x0: f32, y0: f32, x1: f32, y1: f32, color: Color) {
        let mut pb = PathBuilder::new();
        pb.move_to(x0, y0);
        pb.line_to(x1, y1);
        let Some(path) = pb.finish() else { return };
        let mut paint = Paint::default();
        paint.set_color(color.skia());
        paint.anti_alias = false;
        self.pixmap.stroke_path(
            &path,
            &paint,
            &Stroke { width: 1.0, ..Stroke::default() },
            Transform::identity(),
            None,
        );
    }

    /// Filled polygon with anti-aliasing (chart slices, wedges).
    pub fn fill_polygon(&mut self, pts: &[(f32, f32)], color: Color) {
        if pts.len() < 3 {
            return;
        }
        let mut pb = PathBuilder::new();
        pb.move_to(pts[0].0, pts[0].1);
        for p in &pts[1..] {
            pb.line_to(p.0, p.1);
        }
        pb.close();
        let Some(path) = pb.finish() else { return };
        let mut paint = Paint::default();
        paint.set_color(color.skia());
        paint.anti_alias = true;
        self.pixmap.fill_path(
            &path,
            &paint,
            tiny_skia::FillRule::Winding,
            Transform::identity(),
            None,
        );
    }

    /// Filled circle (chart markers).
    pub fn fill_circle(&mut self, cx: f32, cy: f32, r: f32, color: Color) {
        let mut pb = PathBuilder::new();
        pb.push_circle(cx, cy, r.max(0.5));
        let Some(path) = pb.finish() else { return };
        let mut paint = Paint::default();
        paint.set_color(color.skia());
        paint.anti_alias = true;
        self.pixmap.fill_path(
            &path,
            &paint,
            tiny_skia::FillRule::Winding,
            Transform::identity(),
            None,
        );
    }

    /// Anti-aliased line with a stroke width (chart lines).
    pub fn thick_line(&mut self, x0: f32, y0: f32, x1: f32, y1: f32, width: f32, color: Color) {
        let mut pb = PathBuilder::new();
        pb.move_to(x0, y0);
        pb.line_to(x1, y1);
        let Some(path) = pb.finish() else { return };
        let mut paint = Paint::default();
        paint.set_color(color.skia());
        paint.anti_alias = true;
        self.pixmap.stroke_path(
            &path,
            &paint,
            &Stroke { width, ..Stroke::default() },
            Transform::identity(),
            None,
        );
    }

    fn font(&self, bold: bool) -> &FontRef<'static> {
        if bold {
            &self.bold
        } else {
            &self.regular
        }
    }

    pub fn text_width(&self, text: &str, size_px: f32, bold: bool) -> f32 {
        let scaled = self.font(bold).as_scaled(PxScale::from(size_px));
        let mut w = 0.0;
        let mut prev = None;
        for ch in text.chars() {
            let id = scaled.glyph_id(ch);
            if let Some(p) = prev {
                w += scaled.kern(p, id);
            }
            w += scaled.h_advance(id);
            prev = Some(id);
        }
        w
    }

    /// Draw one line of text with `y` as the baseline. Returns the advance.
    pub fn draw_text(&mut self, text: &str, x: f32, y: f32, size_px: f32, color: Color, bold: bool) -> f32 {
        self.draw_text_styled(text, x, y, size_px, color, bold, false)
    }

    /// `draw_text` with synthetic italics: an oblique shear of ~12° applied
    /// per pixel row (the embedded fonts have no italic faces).
    #[allow(clippy::too_many_arguments)]
    pub fn draw_text_styled(
        &mut self,
        text: &str,
        x: f32,
        y: f32,
        size_px: f32,
        color: Color,
        bold: bool,
        italic: bool,
    ) -> f32 {
        let font = self.font(bold).clone();
        let scaled = font.as_scaled(PxScale::from(size_px));
        let mut caret = x;
        let mut prev = None;
        let (width, height) = (self.pixmap.width() as i32, self.pixmap.height() as i32);
        for ch in text.chars() {
            let id = scaled.glyph_id(ch);
            if let Some(p) = prev {
                caret += scaled.kern(p, id);
            }
            let glyph = id.with_scale_and_position(PxScale::from(size_px), ab_glyph::point(caret, y));
            if let Some(outlined) = font.outline_glyph(glyph) {
                let bounds = outlined.px_bounds();
                let data = self.pixmap.pixels_mut();
                outlined.draw(|gx, gy, cov| {
                    let py = bounds.min.y as i32 + gy as i32;
                    let shear = if italic {
                        ((y - py as f32) * 0.21) as i32
                    } else {
                        0
                    };
                    let px = bounds.min.x as i32 + gx as i32 + shear;
                    if px < 0 || py < 0 || px >= width || py >= height || cov <= 0.0 {
                        return;
                    }
                    let idx = (py * width + px) as usize;
                    let dst = data[idx];
                    let a = cov.min(1.0);
                    let blend = |s: u8, d: u8| -> u8 {
                        (s as f32 * a + d as f32 * (1.0 - a)).round() as u8
                    };
                    data[idx] = tiny_skia::PremultipliedColorU8::from_rgba(
                        blend(color.r, dst.red()),
                        blend(color.g, dst.green()),
                        blend(color.b, dst.blue()),
                        blend(255, dst.alpha()),
                    )
                    .unwrap_or(dst);
                });
            }
            caret += scaled.h_advance(id);
            prev = Some(id);
        }
        caret - x
    }

    /// Ascent for baseline placement at a font size.
    pub fn ascent(&self, size_px: f32, bold: bool) -> f32 {
        self.font(bold).as_scaled(PxScale::from(size_px)).ascent()
    }

    /// Wrap formatted spans into lines of at most `max_w` pixels. `\n`
    /// inside a span forces a line break; adjacent same-format words merge.
    pub fn layout_spans(&self, spans: &[Span], max_w: f32) -> Vec<Vec<Span>> {
        let mut lines: Vec<Vec<Span>> = Vec::new();
        let mut cur: Vec<Span> = Vec::new();
        let mut cur_w = 0.0f32;
        for span in spans {
            for (i, seg) in span.text.split('\n').enumerate() {
                if i > 0 {
                    lines.push(std::mem::take(&mut cur));
                    cur_w = 0.0;
                }
                for word in seg.split(' ').filter(|w| !w.is_empty()) {
                    let word_w = self.text_width(word, span.size, span.bold);
                    let space_w = if cur.is_empty() {
                        0.0
                    } else {
                        self.text_width(" ", span.size, span.bold)
                    };
                    if !cur.is_empty() && cur_w + space_w + word_w > max_w {
                        lines.push(std::mem::take(&mut cur));
                        cur_w = 0.0;
                    }
                    let matches_last = cur
                        .last()
                        .map(|l| {
                            l.size == span.size
                                && l.color == span.color
                                && l.bold == span.bold
                                && l.italic == span.italic
                        })
                        .unwrap_or(false);
                    if matches_last {
                        let last = cur.last_mut().unwrap();
                        last.text.push(' ');
                        last.text.push_str(word);
                        cur_w += space_w + word_w;
                    } else {
                        if let Some(last) = cur.last_mut() {
                            last.text.push(' ');
                            cur_w += space_w;
                        }
                        cur.push(Span {
                            text: word.to_string(),
                            size: span.size,
                            color: span.color,
                            bold: span.bold,
                            italic: span.italic,
                        });
                        cur_w += word_w;
                    }
                }
            }
        }
        lines.push(cur);
        lines
    }

    pub fn spans_width(&self, line: &[Span]) -> f32 {
        line.iter()
            .map(|s| self.text_width(&s.text, s.size, s.bold))
            .sum()
    }

    /// Draw one laid-out line with `baseline` shared across spans.
    pub fn draw_spans_line(&mut self, line: &[Span], x: f32, baseline: f32) {
        let mut caret = x;
        for span in line {
            caret += self.draw_text_styled(
                &span.text,
                caret,
                baseline,
                span.size,
                span.color,
                span.bold,
                span.italic,
            );
        }
    }

    /// Draw a line stretched to `width` by widening word gaps (justified
    /// alignment). Falls back to normal drawing when the line has fewer
    /// than two words or already fills the width.
    pub fn draw_spans_justified(&mut self, line: &[Span], x: f32, baseline: f32, width: f32) {
        let mut words: Vec<(String, f32, Color, bool, bool)> = Vec::new();
        for span in line {
            for w in span.text.split(' ').filter(|w| !w.is_empty()) {
                words.push((w.to_string(), span.size, span.color, span.bold, span.italic));
            }
        }
        if words.len() < 2 {
            self.draw_spans_line(line, x, baseline);
            return;
        }
        let words_w: f32 = words
            .iter()
            .map(|(t, s, _, b, _)| self.text_width(t, *s, *b))
            .sum();
        let gap = (width - words_w) / (words.len() - 1) as f32;
        if gap <= 0.0 {
            self.draw_spans_line(line, x, baseline);
            return;
        }
        let mut caret = x;
        for (t, s, c, b, i) in words {
            caret += self.draw_text_styled(&t, caret, baseline, s, c, b, i);
            caret += gap;
        }
    }

    /// Composite image bytes (PNG, JPEG, or GIF first frame) into the given
    /// rectangle, scaled to fit exactly. Undecodable bytes draw a
    /// placeholder. Returns whether the image decoded.
    pub fn draw_image(&mut self, bytes: &[u8], x: f32, y: f32, w: f32, h: f32) -> bool {
        if let Some(img) = decode_image(bytes) {
            let sx = w / img.width() as f32;
            let sy = h / img.height() as f32;
            let transform = Transform::from_scale(sx, sy).post_translate(x, y);
            self.pixmap
                .draw_pixmap(0, 0, img.as_ref(), &PixmapPaint::default(), transform, None);
            return true;
        }
        // Placeholder: framed light box with a label.
        self.fill_rect(x, y, w, h, Color { r: 0xEE, g: 0xEE, b: 0xEE });
        self.stroke_rect(x, y, w, h, GRID);
        let label = "[image]";
        let size = 12.0f32.min(h * 0.5);
        let tw = self.text_width(label, size, false);
        self.draw_text(
            label,
            x + (w - tw) / 2.0,
            y + h / 2.0 + size / 3.0,
            size,
            MUTED,
            false,
        );
        false
    }

    pub fn png(&self) -> Result<Vec<u8>> {
        self.pixmap.encode_png().context("cannot encode PNG")
    }
}

/// Sniff and decode PNG, JPEG, or GIF (first frame) into a Pixmap.
fn decode_image(bytes: &[u8]) -> Option<Pixmap> {
    if bytes.starts_with(b"\x89PNG") {
        return Pixmap::decode_png(bytes).ok();
    }
    if bytes.starts_with(&[0xFF, 0xD8]) {
        return decode_jpeg(bytes);
    }
    if bytes.starts_with(b"GIF8") {
        return decode_gif(bytes);
    }
    Pixmap::decode_png(bytes).ok()
}

fn decode_jpeg(bytes: &[u8]) -> Option<Pixmap> {
    use jpeg_decoder::PixelFormat;
    let mut decoder = jpeg_decoder::Decoder::new(std::io::Cursor::new(bytes));
    let pixels = decoder.decode().ok()?;
    let info = decoder.info()?;
    let (w, h) = (info.width as u32, info.height as u32);
    let mut pm = Pixmap::new(w.max(1), h.max(1))?;
    let out = pm.pixels_mut();
    let n = (w * h) as usize;
    let px = |r: u8, g: u8, b: u8| tiny_skia::PremultipliedColorU8::from_rgba(r, g, b, 255);
    match info.pixel_format {
        PixelFormat::RGB24 => {
            for i in 0..n.min(pixels.len() / 3) {
                out[i] = px(pixels[i * 3], pixels[i * 3 + 1], pixels[i * 3 + 2])?;
            }
        }
        PixelFormat::L8 => {
            for i in 0..n.min(pixels.len()) {
                out[i] = px(pixels[i], pixels[i], pixels[i])?;
            }
        }
        PixelFormat::L16 => {
            for i in 0..n.min(pixels.len() / 2) {
                let v = pixels[i * 2]; // high byte of big-endian L16
                out[i] = px(v, v, v)?;
            }
        }
        PixelFormat::CMYK32 => {
            for i in 0..n.min(pixels.len() / 4) {
                let (c, m, yv, k) = (
                    pixels[i * 4] as u32,
                    pixels[i * 4 + 1] as u32,
                    pixels[i * 4 + 2] as u32,
                    pixels[i * 4 + 3] as u32,
                );
                // Adobe-style inverted CMYK as produced by jpeg-decoder.
                out[i] = px(
                    (c * k / 255) as u8,
                    (m * k / 255) as u8,
                    (yv * k / 255) as u8,
                )?;
            }
        }
    }
    Some(pm)
}

fn decode_gif(bytes: &[u8]) -> Option<Pixmap> {
    let mut opts = gif::DecodeOptions::new();
    opts.set_color_output(gif::ColorOutput::RGBA);
    let mut reader = opts.read_info(std::io::Cursor::new(bytes)).ok()?;
    let (gw, gh) = (reader.width() as u32, reader.height() as u32);
    let mut pm = Pixmap::new(gw.max(1), gh.max(1))?;
    let frame = reader.read_next_frame().ok()??;
    let out = pm.pixels_mut();
    for row in 0..frame.height as u32 {
        for col in 0..frame.width as u32 {
            let src = ((row * frame.width as u32 + col) * 4) as usize;
            if src + 3 >= frame.buffer.len() {
                continue;
            }
            let (tx, ty) = (frame.left as u32 + col, frame.top as u32 + row);
            if tx >= gw || ty >= gh {
                continue;
            }
            let c = tiny_skia::ColorU8::from_rgba(
                frame.buffer[src],
                frame.buffer[src + 1],
                frame.buffer[src + 2],
                frame.buffer[src + 3],
            );
            out[(ty * gw + tx) as usize] = c.premultiply();
        }
    }
    Some(pm)
}
