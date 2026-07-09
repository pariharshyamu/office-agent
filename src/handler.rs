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
}
