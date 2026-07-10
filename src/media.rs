//! Image loading for `add --type image`: format sniffing, intrinsic size,
//! and package/content-type bookkeeping shared by the docx/pptx handlers.

use anyhow::{Context, Result};
use std::path::Path;

use crate::pkg::Package;
use crate::xml::el;

pub struct Image {
    pub bytes: Vec<u8>,
    pub extension: &'static str,
    pub content_type: &'static str,
    /// Intrinsic size in EMU at 96 dpi.
    pub width_emu: i64,
    pub height_emu: i64,
}

const EMU_PER_PX: i64 = 9_525; // 96 dpi

pub fn load_image(path: &str) -> Result<Image> {
    let bytes = std::fs::read(Path::new(path))
        .with_context(|| format!("cannot read image file '{path}'"))?;
    load_image_bytes(bytes).with_context(|| format!("cannot load image '{path}'"))
}

pub fn load_image_bytes(bytes: Vec<u8>) -> Result<Image> {
    let (extension, content_type, dims) =
        sniff(&bytes).context("not a supported image (PNG/JPEG/GIF)")?;
    let (w, h) = dims.unwrap_or((300, 200));
    Ok(Image {
        bytes,
        extension,
        content_type,
        width_emu: w as i64 * EMU_PER_PX,
        height_emu: h as i64 * EMU_PER_PX,
    })
}

/// Resolve `src=path` or `srcdata=BASE64` props into an Image.
pub fn image_from_props(props: &crate::props::Props) -> Result<Image> {
    if let Some(data) = props.get("srcdata") {
        use base64::Engine;
        let cleaned: String = data.chars().filter(|c| !c.is_whitespace()).collect();
        let bytes = base64::engine::general_purpose::STANDARD
            .decode(cleaned.as_bytes())
            .context("srcdata is not valid base64")?;
        return load_image_bytes(bytes).context("srcdata does not decode to a PNG/JPEG/GIF");
    }
    let src = props
        .get("src")
        .context("image needs --prop src=path/to/file.png (or srcdata=BASE64)")?;
    load_image(src)
}

type Sniffed = (&'static str, &'static str, Option<(u32, u32)>);

fn sniff(bytes: &[u8]) -> Option<Sniffed> {
    if bytes.starts_with(&[0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]) {
        return Some(("png", "image/png", png_dims(bytes)));
    }
    if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
        return Some(("jpeg", "image/jpeg", jpeg_dims(bytes)));
    }
    if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        return Some(("gif", "image/gif", gif_dims(bytes)));
    }
    None
}

fn png_dims(bytes: &[u8]) -> Option<(u32, u32)> {
    // IHDR is the first chunk: width/height at offsets 16 and 20.
    if bytes.len() < 24 || &bytes[12..16] != b"IHDR" {
        return None;
    }
    let w = u32::from_be_bytes(bytes[16..20].try_into().ok()?);
    let h = u32::from_be_bytes(bytes[20..24].try_into().ok()?);
    Some((w, h))
}

fn jpeg_dims(bytes: &[u8]) -> Option<(u32, u32)> {
    // Scan markers for a start-of-frame segment.
    let mut i = 2;
    while i + 9 < bytes.len() {
        if bytes[i] != 0xFF {
            i += 1;
            continue;
        }
        let marker = bytes[i + 1];
        // SOF0..SOF15 except DHT(C4)/JPG(C8)/DAC(CC).
        if (0xC0..=0xCF).contains(&marker) && !matches!(marker, 0xC4 | 0xC8 | 0xCC) {
            let h = u16::from_be_bytes([bytes[i + 5], bytes[i + 6]]) as u32;
            let w = u16::from_be_bytes([bytes[i + 7], bytes[i + 8]]) as u32;
            return Some((w, h));
        }
        if marker == 0xD8 || (0xD0..=0xD9).contains(&marker) {
            i += 2;
            continue;
        }
        let len = u16::from_be_bytes([bytes[i + 2], bytes[i + 3]]) as usize;
        i += 2 + len;
    }
    None
}

fn gif_dims(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.len() < 10 {
        return None;
    }
    let w = u16::from_le_bytes([bytes[6], bytes[7]]) as u32;
    let h = u16::from_le_bytes([bytes[8], bytes[9]]) as u32;
    Some((w, h))
}

/// Store the image under `media_dir` (e.g. "word/media") with the next free
/// image number, ensuring a content-type Default for the extension exists.
/// Returns the created part name.
pub fn store_image(pkg: &mut Package, media_dir: &str, image: &Image) -> Result<String> {
    let prefix = format!("{media_dir}/image");
    let next = pkg
        .part_names()
        .filter_map(|p| {
            p.strip_prefix(&prefix)
                .and_then(|s| s.split('.').next())
                .and_then(|n| n.parse::<u32>().ok())
        })
        .max()
        .unwrap_or(0)
        + 1;
    let part = format!("{media_dir}/image{next}.{}", image.extension);

    // Ensure [Content_Types].xml has a Default for this extension.
    let mut ct = pkg.xml("[Content_Types].xml")?;
    let have = ct.children_named("Default").into_iter().any(|d| {
        d.attr_local("Extension")
            .map(|e| e.eq_ignore_ascii_case(image.extension))
            .unwrap_or(false)
    });
    if !have {
        // Defaults must appear before Overrides is not required by spec;
        // append is fine.
        ct.push(el(
            "Default",
            &[
                ("Extension", image.extension),
                ("ContentType", image.content_type),
            ],
        ));
        pkg.put_xml("[Content_Types].xml", &ct)?;
    }
    pkg.put_raw(&part, image.bytes.clone());
    Ok(part)
}

/// Add a relationship of the given type to a rels root, returning the new rId.
pub fn add_relationship(rels: &mut crate::xml::XmlElement, rel_type: &str, target: &str) -> String {
    let next = rels
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
    let rid = format!("rId{next}");
    rels.push(el(
        "Relationship",
        &[("Id", rid.as_str()), ("Type", rel_type), ("Target", target)],
    ));
    rid
}

pub const IMAGE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sniffs_png_dimensions() {
        // Minimal PNG header: signature + IHDR length/name + 4x3 size.
        let mut bytes = vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];
        bytes.extend_from_slice(&[0, 0, 0, 13]);
        bytes.extend_from_slice(b"IHDR");
        bytes.extend_from_slice(&4u32.to_be_bytes());
        bytes.extend_from_slice(&3u32.to_be_bytes());
        let (ext, _, dims) = sniff(&bytes).unwrap();
        assert_eq!(ext, "png");
        assert_eq!(dims, Some((4, 3)));
    }

    #[test]
    fn sniffs_gif_dimensions() {
        let mut bytes = b"GIF89a".to_vec();
        bytes.extend_from_slice(&[10, 0, 20, 0]);
        let (ext, _, dims) = sniff(&bytes).unwrap();
        assert_eq!(ext, "gif");
        assert_eq!(dims, Some((10, 20)));
    }
}
