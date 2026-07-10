//! OOXML package handling: an .docx/.xlsx/.pptx file is a zip archive of
//! "parts". The whole archive is held in memory; XML parts are parsed lazily
//! and written back only if the caller replaced them.

use anyhow::{bail, Context, Result};
use std::collections::BTreeMap;
use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::Path;
use zip::write::SimpleFileOptions;

use crate::xml::{self, XmlElement};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocKind {
    Docx,
    Xlsx,
    Pptx,
}

impl DocKind {
    pub fn from_path(path: &Path) -> Result<DocKind> {
        match path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_ascii_lowercase())
            .as_deref()
        {
            Some("docx") => Ok(DocKind::Docx),
            Some("xlsx") => Ok(DocKind::Xlsx),
            Some("pptx") => Ok(DocKind::Pptx),
            other => bail!(
                "unsupported file extension {:?}: expected .docx, .xlsx, or .pptx",
                other.unwrap_or("<none>")
            ),
        }
    }

    pub fn format_name(&self) -> &'static str {
        match self {
            DocKind::Docx => "docx",
            DocKind::Xlsx => "xlsx",
            DocKind::Pptx => "pptx",
        }
    }
}

pub struct Package {
    pub kind: DocKind,
    /// part name (zip entry path, no leading slash) -> raw bytes
    parts: BTreeMap<String, Vec<u8>>,
}

impl Package {
    pub fn open(path: &Path) -> Result<Package> {
        let kind = DocKind::from_path(path)?;
        let bytes = fs::read(path)
            .with_context(|| format!("cannot read {}", path.display()))?;
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
            .with_context(|| format!("{} is not a valid Office file (zip)", path.display()))?;
        let mut parts = BTreeMap::new();
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            if file.is_dir() {
                continue;
            }
            let mut buf = Vec::with_capacity(file.size() as usize);
            file.read_to_end(&mut buf)?;
            parts.insert(file.name().to_string(), buf);
        }
        Ok(Package { kind, parts })
    }

    pub fn from_parts(kind: DocKind, parts: Vec<(&str, Vec<u8>)>) -> Package {
        Package {
            kind,
            parts: parts
                .into_iter()
                .map(|(k, v)| (k.to_string(), v))
                .collect(),
        }
    }

    pub fn save(&self, path: &Path) -> Result<()> {
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut zip = zip::ZipWriter::new(&mut cursor);
            let options = SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            for (name, bytes) in &self.parts {
                zip.start_file(name, options)?;
                zip.write_all(bytes)?;
            }
            zip.finish()?;
        }
        fs::write(path, cursor.into_inner())
            .with_context(|| format!("cannot write {}", path.display()))?;
        Ok(())
    }

    pub fn has_part(&self, name: &str) -> bool {
        self.parts.contains_key(name)
    }

    pub fn part_names(&self) -> impl Iterator<Item = &str> {
        self.parts.keys().map(|s| s.as_str())
    }

    pub fn raw(&self, name: &str) -> Result<&[u8]> {
        self.parts
            .get(name)
            .map(|v| v.as_slice())
            .with_context(|| format!("package has no part '{name}'"))
    }

    pub fn xml(&self, name: &str) -> Result<XmlElement> {
        xml::parse(self.raw(name)?)
            .with_context(|| format!("cannot parse XML part '{name}'"))
    }

    pub fn put_xml(&mut self, name: &str, root: &XmlElement) -> Result<()> {
        self.parts.insert(name.to_string(), xml::serialize(root)?);
        Ok(())
    }

    pub fn put_raw(&mut self, name: &str, bytes: Vec<u8>) {
        self.parts.insert(name.to_string(), bytes);
    }

    pub fn remove_part(&mut self, name: &str) {
        self.parts.remove(name);
    }

    /// Ensure `[Content_Types].xml` has an Override for `part`.
    pub fn add_override(&mut self, part: &str, content_type: &str) -> Result<()> {
        let mut ct = self.xml("[Content_Types].xml")?;
        let part_name = format!("/{part}");
        let exists = ct
            .children_named("Override")
            .iter()
            .any(|o| o.attr_local("PartName") == Some(part_name.as_str()));
        if !exists {
            ct.push(xml::el(
                "Override",
                &[
                    ("PartName", part_name.as_str()),
                    ("ContentType", content_type),
                ],
            ));
            self.put_xml("[Content_Types].xml", &ct)?;
        }
        Ok(())
    }
}
