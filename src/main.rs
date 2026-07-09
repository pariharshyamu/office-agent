//! officecli (Rust port) — a command-line Office suite for AI agents.
//!
//! Reads, creates, and edits .docx / .xlsx / .pptx files with a DOM-like
//! path system, mirroring the command surface of iOfficeAI/OfficeCLI.

mod batch;
mod docx;
mod handler;
mod helptext;
mod out;
mod path;
mod pkg;
mod pptx;
mod props;
mod templates;
mod xlsx;
mod xml;

use anyhow::{bail, Context, Result};
use clap::{Parser, Subcommand};
use std::path::{Path, PathBuf};

use handler::{Handler, Position};
use out::Report;
use pkg::{DocKind, Package};
use props::Props;

#[derive(Parser)]
#[command(
    name = "officecli",
    version,
    about = "Office suite CLI for AI agents: read, edit, and automate .docx/.xlsx/.pptx (Rust port of OfficeCLI)",
    after_help = "Run `officecli help` for the agent-oriented command guide.",
    disable_help_subcommand = true
)]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// Create a blank document (type inferred from the extension)
    Create {
        file: PathBuf,
        /// Overwrite if the file already exists
        #[arg(long)]
        force: bool,
        #[arg(long)]
        json: bool,
    },
    /// Show a document as text, outline, or stats
    View {
        file: PathBuf,
        /// text | outline | stats
        #[arg(default_value = "outline")]
        mode: String,
        #[arg(long)]
        json: bool,
    },
    /// Get an element (and optionally its children) by path
    Get {
        file: PathBuf,
        /// Element path, e.g. '/body/p[3]', '/slide[1]/shape[2]', '/Sheet1/A1'
        path: String,
        /// Expand children N levels deep
        #[arg(long, default_value_t = 0)]
        depth: usize,
        #[arg(long)]
        json: bool,
    },
    /// Add an element under a parent path
    Add {
        file: PathBuf,
        /// Parent path, e.g. '/', '/body', '/slide[1]', '/Sheet1'
        parent: String,
        /// Element type: paragraph, run, table, row, slide, shape, sheet, ...
        #[arg(long = "type")]
        typ: String,
        /// Properties as key=value (repeatable)
        #[arg(long = "prop")]
        props: Vec<String>,
        /// 0-based insert position among siblings
        #[arg(long, conflicts_with_all = ["before", "after"])]
        index: Option<usize>,
        /// Insert before this sibling path
        #[arg(long, conflicts_with = "after")]
        before: Option<String>,
        /// Insert after this sibling path
        #[arg(long)]
        after: Option<String>,
        #[arg(long)]
        json: bool,
    },
    /// Set properties on an element, or find/replace text
    Set {
        file: PathBuf,
        /// Element path ('/' for whole-document find/replace)
        path: String,
        /// Properties as key=value (repeatable)
        #[arg(long = "prop")]
        props: Vec<String>,
        /// Text to find (literal; add --prop regex=true for regex)
        #[arg(long)]
        find: Option<String>,
        /// Replacement text (with --find)
        #[arg(long, requires = "find")]
        replace: Option<String>,
        #[arg(long)]
        json: bool,
    },
    /// Remove an element by path
    Remove {
        file: PathBuf,
        path: String,
        #[arg(long)]
        json: bool,
    },
    /// Run multiple operations from JSON in one save cycle
    Batch {
        file: PathBuf,
        /// Inline JSON array of operations
        #[arg(long)]
        commands: Option<String>,
        /// Read operations from a JSON file
        #[arg(long, conflicts_with = "commands")]
        input: Option<PathBuf>,
        /// Abort on the first failing operation
        #[arg(long)]
        stop_on_error: bool,
        #[arg(long)]
        json: bool,
    },
    /// Check package structure
    Validate {
        file: PathBuf,
        #[arg(long)]
        json: bool,
    },
    /// Show the agent-oriented command guide
    Help {
        /// Optional topic: docx | xlsx | pptx
        topic: Option<String>,
    },
}

fn open_handler(file: &Path) -> Result<Box<dyn Handler>> {
    if !file.exists() {
        bail!(
            "{} does not exist (use `officecli create {}` first)",
            file.display(),
            file.display()
        );
    }
    let package = Package::open(file)?;
    make_handler(package)
}

fn make_handler(package: Package) -> Result<Box<dyn Handler>> {
    Ok(match package.kind {
        DocKind::Docx => Box::new(docx::Docx::new(package)?),
        DocKind::Xlsx => Box::new(xlsx::Xlsx::new(package)?),
        DocKind::Pptx => Box::new(pptx::Pptx::new(package)?),
    })
}

fn position(index: Option<usize>, before: Option<String>, after: Option<String>) -> Position {
    if let Some(i) = index {
        Position::Index(i)
    } else if let Some(b) = before {
        Position::Before(b)
    } else if let Some(a) = after {
        Position::After(a)
    } else {
        Position::Append
    }
}

fn run() -> Result<(Report, bool)> {
    let cli = Cli::parse();
    match cli.command {
        Command::Create { file, force, json } => {
            let kind = DocKind::from_path(&file)?;
            if file.exists() && !force {
                bail!("{} already exists (use --force to overwrite)", file.display());
            }
            let package = templates::blank_package(kind);
            package.save(&file)?;
            Ok((
                Report::message(format!(
                    "created blank {} at {}",
                    kind.format_name(),
                    file.display()
                )),
                json,
            ))
        }
        Command::View { file, mode, json } => {
            let mut handler = open_handler(&file)?;
            Ok((handler.view(&mode)?, json))
        }
        Command::Get {
            file,
            path,
            depth,
            json,
        } => {
            let mut handler = open_handler(&file)?;
            Ok((handler.get(&path, depth)?, json))
        }
        Command::Add {
            file,
            parent,
            typ,
            props,
            index,
            before,
            after,
            json,
        } => {
            let props = Props::from_args(&props)?;
            let pos = position(index, before, after);
            let mut handler = open_handler(&file)?;
            let report = handler.add(&parent, &typ, &props, &pos)?;
            handler.save(&file)?;
            Ok((report, json))
        }
        Command::Set {
            file,
            path,
            props,
            find,
            replace,
            json,
        } => {
            let props = Props::from_args(&props)?;
            let mut handler = open_handler(&file)?;
            let report = handler.set(&path, &props, find.as_deref(), replace.as_deref())?;
            handler.save(&file)?;
            Ok((report, json))
        }
        Command::Remove { file, path, json } => {
            let mut handler = open_handler(&file)?;
            let report = handler.remove(&path)?;
            handler.save(&file)?;
            Ok((report, json))
        }
        Command::Batch {
            file,
            commands,
            input,
            stop_on_error,
            json,
        } => {
            let source = if let Some(inline) = commands {
                inline
            } else if let Some(path) = input {
                std::fs::read_to_string(&path)
                    .with_context(|| format!("cannot read {}", path.display()))?
            } else {
                use std::io::Read;
                let mut buf = String::new();
                std::io::stdin()
                    .read_to_string(&mut buf)
                    .context("cannot read batch JSON from stdin")?;
                buf
            };
            let mut handler = open_handler(&file)?;
            let (report, all_ok) = batch::run_batch(handler.as_mut(), &source, stop_on_error)?;
            handler.save(&file)?;
            if !all_ok {
                // Save partial progress but signal failure via exit code.
                println!("{}", report.render(json));
                std::process::exit(1);
            }
            Ok((report, json))
        }
        Command::Validate { file, json } => {
            let mut handler = open_handler(&file)?;
            Ok((handler.validate()?, json))
        }
        Command::Help { topic } => Ok((
            Report::Text(helptext::help_for(topic.as_deref())?),
            false,
        )),
    }
}

fn main() {
    // `--json` errors also go to stdout as a JSON envelope so agents can
    // parse failures uniformly.
    let wants_json = std::env::args().any(|a| a == "--json");
    match run() {
        Ok((report, json)) => println!("{}", report.render(json)),
        Err(err) => {
            let message = format!("{err:#}");
            if wants_json {
                println!("{}", out::error_json(&message));
            } else {
                eprintln!("error: {message}");
            }
            std::process::exit(1);
        }
    }
}
