//! `--prop key=value` parsing plus value conversions shared by all formats:
//! lengths (EMU), colors, font sizes, booleans, alignment.

use anyhow::{bail, Result};

#[derive(Debug, Clone, Default)]
pub struct Props {
    pairs: Vec<(String, String)>,
}

impl Props {
    pub fn from_args(args: &[String]) -> Result<Props> {
        let mut pairs = Vec::new();
        for arg in args {
            match arg.split_once('=') {
                Some((k, v)) => pairs.push((k.trim().to_string(), v.to_string())),
                None => bail!("--prop '{arg}' must be key=value"),
            }
        }
        Ok(Props { pairs })
    }

    pub fn from_pairs(pairs: Vec<(String, String)>) -> Props {
        Props { pairs }
    }

    pub fn get(&self, key: &str) -> Option<&str> {
        self.pairs
            .iter()
            .find(|(k, _)| k.eq_ignore_ascii_case(key))
            .map(|(_, v)| v.as_str())
    }

    pub fn has(&self, key: &str) -> bool {
        self.get(key).is_some()
    }

    pub fn is_empty(&self) -> bool {
        self.pairs.is_empty()
    }

    pub fn iter(&self) -> impl Iterator<Item = (&str, &str)> {
        self.pairs.iter().map(|(k, v)| (k.as_str(), v.as_str()))
    }

    pub fn get_bool(&self, key: &str) -> Result<Option<bool>> {
        match self.get(key) {
            None => Ok(None),
            Some(v) => Ok(Some(parse_bool(v)?)),
        }
    }
}

pub fn parse_bool(v: &str) -> Result<bool> {
    match v.to_ascii_lowercase().as_str() {
        "true" | "1" | "yes" | "on" => Ok(true),
        "false" | "0" | "no" | "off" => Ok(false),
        other => bail!("'{other}' is not a boolean (use true/false)"),
    }
}

/// Parse a length into EMU. Accepts `2cm`, `1in`, `72pt`, `96px`, `2.5cm`,
/// or a bare number which is taken as EMU directly.
pub fn parse_emu(v: &str) -> Result<i64> {
    let v = v.trim();
    let (num, unit) = split_unit(v);
    let n: f64 = num
        .parse()
        .map_err(|_| anyhow::anyhow!("'{v}' is not a valid length"))?;
    let emu = match unit.to_ascii_lowercase().as_str() {
        "" | "emu" => n,
        "cm" => n * 360_000.0,
        "mm" => n * 36_000.0,
        "in" => n * 914_400.0,
        "pt" => n * 12_700.0,
        "px" => n * 9_525.0,
        other => bail!("unknown length unit '{other}' in '{v}' (use cm/mm/in/pt/px/emu)"),
    };
    Ok(emu.round() as i64)
}

/// Parse a font size in points. Accepts `24`, `24pt`, `18.5pt`.
pub fn parse_pt(v: &str) -> Result<f64> {
    let (num, unit) = split_unit(v.trim());
    if !unit.is_empty() && !unit.eq_ignore_ascii_case("pt") {
        bail!("font sizes are in points; got unit '{unit}' in '{v}'");
    }
    num.parse()
        .map_err(|_| anyhow::anyhow!("'{v}' is not a valid point size"))
}

fn split_unit(v: &str) -> (&str, &str) {
    let idx = v
        .char_indices()
        .find(|(_, c)| c.is_ascii_alphabetic())
        .map(|(i, _)| i)
        .unwrap_or(v.len());
    (&v[..idx], &v[idx..])
}

/// Normalize a color to a 6-hex-digit uppercase RRGGBB string (no '#').
/// Accepts `FF0000`, `#FF0000`, `red`, `rgb(255,0,0)`.
pub fn parse_color(v: &str) -> Result<String> {
    let v = v.trim();
    if let Some(rgb) = v.strip_prefix("rgb(").and_then(|s| s.strip_suffix(')')) {
        let parts: Vec<&str> = rgb.split(',').map(|p| p.trim()).collect();
        if parts.len() != 3 {
            bail!("'{v}' is not a valid rgb() color");
        }
        let mut out = String::new();
        for p in parts {
            let n: u8 = p
                .parse()
                .map_err(|_| anyhow::anyhow!("'{v}' has a non-numeric rgb component"))?;
            out.push_str(&format!("{n:02X}"));
        }
        return Ok(out);
    }
    let hex = v.strip_prefix('#').unwrap_or(v);
    if hex.len() == 6 && hex.chars().all(|c| c.is_ascii_hexdigit()) {
        return Ok(hex.to_ascii_uppercase());
    }
    if hex.len() == 3 && hex.chars().all(|c| c.is_ascii_hexdigit()) {
        let mut out = String::new();
        for c in hex.chars() {
            out.push(c.to_ascii_uppercase());
            out.push(c.to_ascii_uppercase());
        }
        return Ok(out);
    }
    match v.to_ascii_lowercase().as_str() {
        "black" => Ok("000000".into()),
        "white" => Ok("FFFFFF".into()),
        "red" => Ok("FF0000".into()),
        "green" => Ok("008000".into()),
        "lime" => Ok("00FF00".into()),
        "blue" => Ok("0000FF".into()),
        "yellow" => Ok("FFFF00".into()),
        "orange" => Ok("FFA500".into()),
        "purple" => Ok("800080".into()),
        "gray" | "grey" => Ok("808080".into()),
        "silver" => Ok("C0C0C0".into()),
        "cyan" | "aqua" => Ok("00FFFF".into()),
        "magenta" | "fuchsia" => Ok("FF00FF".into()),
        "navy" => Ok("000080".into()),
        "teal" => Ok("008080".into()),
        "maroon" => Ok("800000".into()),
        "olive" => Ok("808000".into()),
        "pink" => Ok("FFC0CB".into()),
        "brown" => Ok("A52A2A".into()),
        "gold" => Ok("FFD700".into()),
        other => bail!("unknown color '{other}' (use hex like FF0000, rgb(255,0,0), or a basic color name)"),
    }
}

/// Normalize alignment names to OOXML `jc`/`algn` values.
/// Returns (docx_jc, pptx_algn).
pub fn parse_align(v: &str) -> Result<(&'static str, &'static str)> {
    match v.to_ascii_lowercase().as_str() {
        "left" | "start" => Ok(("left", "l")),
        "center" | "centre" => Ok(("center", "ctr")),
        "right" | "end" => Ok(("right", "r")),
        "justify" | "both" => Ok(("both", "just")),
        other => bail!("unknown alignment '{other}' (left/center/right/justify)"),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lengths() {
        assert_eq!(parse_emu("2cm").unwrap(), 720_000);
        assert_eq!(parse_emu("1in").unwrap(), 914_400);
        assert_eq!(parse_emu("72pt").unwrap(), 914_400);
        assert_eq!(parse_emu("914400").unwrap(), 914_400);
        assert!(parse_emu("2light-years").is_err());
    }

    #[test]
    fn colors() {
        assert_eq!(parse_color("#ff0000").unwrap(), "FF0000");
        assert_eq!(parse_color("red").unwrap(), "FF0000");
        assert_eq!(parse_color("rgb(0, 128, 255)").unwrap(), "0080FF");
        assert_eq!(parse_color("abc").unwrap(), "AABBCC");
        assert!(parse_color("chartreuse-ish").is_err());
    }
}
