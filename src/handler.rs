//! Format-independent command interface implemented by the docx/xlsx/pptx
//! handlers.

use anyhow::Result;
use std::path::Path;

use crate::out::Report;
use crate::props::Props;

/// Where `add` inserts the new element among its siblings.
#[derive(Debug, Clone, Default)]
pub enum Position {
    #[default]
    Append,
    /// 0-based position (legacy `--index`).
    Index(usize),
    /// Insert before the element at this document path.
    Before(String),
    /// Insert after the element at this document path.
    After(String),
}

pub trait Handler {
    fn view(&mut self, mode: &str) -> Result<Report>;
    fn get(&mut self, path: &str, depth: usize) -> Result<Report>;
    fn add(&mut self, parent: &str, typ: &str, props: &Props, pos: &Position) -> Result<Report>;
    fn set(
        &mut self,
        path: &str,
        props: &Props,
        find: Option<&str>,
        replace: Option<&str>,
    ) -> Result<Report>;
    fn remove(&mut self, path: &str) -> Result<Report>;
    fn validate(&mut self) -> Result<Report>;
    /// Persist any pending changes back to the package and write it to disk.
    fn save(&mut self, path: &Path) -> Result<()>;

    /// Full document as a NodeInfo forest, used by `query`.
    fn tree(&mut self) -> Result<Vec<crate::out::NodeInfo>>;

    /// Move an element to a new position (optionally a new parent).
    fn move_el(&mut self, path: &str, to: Option<&str>, pos: &Position) -> Result<Report> {
        let _ = (path, to, pos);
        anyhow::bail!("move is not supported for this format")
    }

    /// Swap two elements.
    fn swap(&mut self, path1: &str, path2: &str) -> Result<Report> {
        let _ = (path1, path2);
        anyhow::bail!("swap is not supported for this format")
    }

    /// Emit a replayable batch-JSON representation of the document content.
    fn dump(&mut self) -> Result<serde_json::Value>;

    /// Render the document as PNGs, one per slide/sheet/page.
    fn screenshot(&mut self) -> Result<Vec<Vec<u8>>> {
        anyhow::bail!("screenshot is not supported for this format")
    }
}
