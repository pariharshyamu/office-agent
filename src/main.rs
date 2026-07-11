//! officecli (Rust port) — a command-line Office suite for AI agents.
//!
//! Reads, creates, and edits .docx / .xlsx / .pptx files with a DOM-like
//! path system, mirroring the command surface of iOfficeAI/OfficeCLI.

mod batch;
mod chart;
mod docx;
mod handler;
mod helptext;
mod html;
mod mcp;
mod media;
mod out;
mod path;
mod pkg;
mod pptx;
mod props;
mod query;
mod render;
mod resident;
mod templates;
mod watch;
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
    /// Show a document as text, outline, stats, or html
    View {
        file: PathBuf,
        /// text | outline | stats | html
        #[arg(default_value = "outline")]
        mode: String,
        /// Write the output to a file instead of stdout
        #[arg(short, long)]
        output: Option<PathBuf>,
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
    /// Find elements with a CSS-like selector
    Query {
        file: PathBuf,
        /// e.g. 'paragraph[style=Normal] > run[font!=Arial]', 'cell[value>100]'
        selector: String,
        #[arg(long)]
        json: bool,
    },
    /// Add an element under a parent path
    Add {
        file: PathBuf,
        /// Parent path, e.g. '/', '/body', '/slide[1]', '/Sheet1'
        parent: String,
        /// Element type: paragraph, run, table, row, image, slide, shape, sheet, ...
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
    /// Move an element to a new position (and optionally a new parent)
    Move {
        file: PathBuf,
        path: String,
        /// New parent path (defaults to the current parent)
        #[arg(long)]
        to: Option<String>,
        /// 0-based target position among siblings
        #[arg(long, conflicts_with_all = ["before", "after"])]
        index: Option<usize>,
        #[arg(long, conflicts_with = "after")]
        before: Option<String>,
        #[arg(long)]
        after: Option<String>,
        #[arg(long)]
        json: bool,
    },
    /// Swap two elements
    Swap {
        file: PathBuf,
        path1: String,
        path2: String,
        #[arg(long)]
        json: bool,
    },
    /// Duplicate an element (slide, sheet, row, paragraph, table, shape)
    Copy {
        file: PathBuf,
        path: String,
        /// Target position (meaning depends on the element kind)
        #[arg(long, conflicts_with_all = ["before", "after"])]
        index: Option<usize>,
        #[arg(long, conflicts_with = "after")]
        before: Option<String>,
        #[arg(long)]
        after: Option<String>,
        #[arg(long)]
        json: bool,
    },
    /// Export a cell range as CSV (xlsx)
    Export {
        file: PathBuf,
        /// Range like 'Sheet1!A1:C9' or 'A1:C9' (default: the used range)
        #[arg(long)]
        range: Option<String>,
        /// Sheet name (default: the first sheet)
        #[arg(long)]
        sheet: Option<String>,
        /// Write to a file instead of stdout
        #[arg(short, long)]
        output: Option<PathBuf>,
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
    /// Emit replayable batch JSON recreating the document content
    Dump {
        file: PathBuf,
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
    /// Run as a Model Context Protocol server on stdio
    Mcp,
    /// Run a long-lived command loop: one command line in, one JSON line out
    Resident,
    /// Serve a live-reloading HTML preview of a document
    Watch {
        file: PathBuf,
        /// Port to listen on (0 picks a free port)
        #[arg(long, default_value_t = 8787)]
        port: u16,
    },
    /// Show the agent-oriented command guide
    Help {
        /// Optional topic: docx | xlsx | pptx
        topic: Option<String>,
    },
    /// Unknown subcommands dispatch to `officecli-<name>` plugins on PATH
    #[command(external_subcommand)]
    External(Vec<String>),
}

pub(crate) fn open_handler(file: &Path) -> Result<Box<dyn Handler>> {
    if !file.exists() {
        bail!(
            "{} does not exist (use `officecli create {}` first)",
            file.display(),
            file.display()
        );
    }
    let package = Package::open(file)?;
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

struct Outcome {
    report: Report,
    json: bool,
    ok: bool,
}

/// Parse and run one command from an argv vector. Used by the MCP server;
/// `allow_stdin` guards commands that would otherwise block on stdin.
pub(crate) fn execute_args(args: Vec<String>, allow_stdin: bool) -> Result<(Report, bool)> {
    let cli = Cli::try_parse_from(args).map_err(|e| anyhow::anyhow!("{e}"))?;
    let outcome = run_command(cli, allow_stdin)?;
    Ok((outcome.report, outcome.ok))
}

fn run_command(cli: Cli, allow_stdin: bool) -> Result<Outcome> {
    let done = |report: Report, json: bool| Ok(Outcome { report, json, ok: true });
    match cli.command {
        Command::Create { file, force, json } => {
            let kind = DocKind::from_path(&file)?;
            if file.exists() && !force {
                bail!("{} already exists (use --force to overwrite)", file.display());
            }
            let package = templates::blank_package(kind);
            package.save(&file)?;
            done(
                Report::message(format!(
                    "created blank {} at {}",
                    kind.format_name(),
                    file.display()
                )),
                json,
            )
        }
        Command::View {
            file,
            mode,
            output,
            json,
        } => {
            let mut handler = open_handler(&file)?;
            if mode == "screenshot" {
                let out_path =
                    output.context("view screenshot needs -o out.png (multi-page → out-1.png, ...)")?;
                let images = handler.screenshot()?;
                let mut written = Vec::new();
                if images.len() == 1 {
                    std::fs::write(&out_path, &images[0])
                        .with_context(|| format!("cannot write {}", out_path.display()))?;
                    written.push(out_path.display().to_string());
                } else {
                    let stem = out_path
                        .file_stem()
                        .and_then(|s| s.to_str())
                        .unwrap_or("screenshot");
                    let ext = out_path.extension().and_then(|s| s.to_str()).unwrap_or("png");
                    for (i, png) in images.iter().enumerate() {
                        let path = out_path.with_file_name(format!("{stem}-{}.{ext}", i + 1));
                        std::fs::write(&path, png)
                            .with_context(|| format!("cannot write {}", path.display()))?;
                        written.push(path.display().to_string());
                    }
                }
                return done(
                    Report::Data {
                        text: format!("wrote {}", written.join(", ")),
                        data: serde_json::json!({ "files": written }),
                    },
                    json,
                );
            }
            let report = handler.view(&mode)?;
            if let Some(out_path) = output {
                let rendered = report.render(false);
                std::fs::write(&out_path, rendered)
                    .with_context(|| format!("cannot write {}", out_path.display()))?;
                return done(
                    Report::message(format!("wrote {} view to {}", mode, out_path.display())),
                    json,
                );
            }
            done(report, json)
        }
        Command::Get {
            file,
            path,
            depth,
            json,
        } => {
            let mut handler = open_handler(&file)?;
            done(handler.get(&path, depth)?, json)
        }
        Command::Query {
            file,
            selector,
            json,
        } => {
            let sel = query::parse(&selector)?;
            let mut handler = open_handler(&file)?;
            let roots = handler.tree()?;
            let hits = query::run(&sel, &roots);
            done(Report::Nodes(hits), json)
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
            done(report, json)
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
            done(report, json)
        }
        Command::Move {
            file,
            path,
            to,
            index,
            before,
            after,
            json,
        } => {
            let pos = position(index, before, after);
            let mut handler = open_handler(&file)?;
            let report = handler.move_el(&path, to.as_deref(), &pos)?;
            handler.save(&file)?;
            done(report, json)
        }
        Command::Swap {
            file,
            path1,
            path2,
            json,
        } => {
            let mut handler = open_handler(&file)?;
            let report = handler.swap(&path1, &path2)?;
            handler.save(&file)?;
            done(report, json)
        }
        Command::Copy {
            file,
            path,
            index,
            before,
            after,
            json,
        } => {
            let pos = position(index, before, after);
            let mut handler = open_handler(&file)?;
            let report = handler.copy_el(&path, &pos)?;
            handler.save(&file)?;
            done(report, json)
        }
        Command::Export {
            file,
            range,
            sheet,
            output,
            json,
        } => {
            let mut handler = open_handler(&file)?;
            let csv = handler.export_csv(sheet.as_deref(), range.as_deref())?;
            if let Some(out_path) = output {
                std::fs::write(&out_path, &csv)
                    .with_context(|| format!("cannot write {}", out_path.display()))?;
                return done(
                    Report::message(format!("wrote CSV to {}", out_path.display())),
                    json,
                );
            }
            done(Report::Text(csv), json)
        }
        Command::Remove { file, path, json } => {
            let mut handler = open_handler(&file)?;
            let report = handler.remove(&path)?;
            handler.save(&file)?;
            done(report, json)
        }
        Command::Dump { file, json } => {
            let mut handler = open_handler(&file)?;
            let ops = handler.dump()?;
            done(
                Report::Data {
                    text: serde_json::to_string_pretty(&ops)?,
                    data: serde_json::json!({ "ops": ops }),
                },
                json,
            )
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
            } else if allow_stdin {
                use std::io::Read;
                let mut buf = String::new();
                std::io::stdin()
                    .read_to_string(&mut buf)
                    .context("cannot read batch JSON from stdin")?;
                buf
            } else {
                bail!("batch over MCP needs --commands or --input (stdin is the protocol channel)");
            };
            let mut handler = open_handler(&file)?;
            let (report, all_ok) = batch::run_batch(handler.as_mut(), &source, stop_on_error)?;
            // Save partial progress even when some items failed.
            handler.save(&file)?;
            Ok(Outcome {
                report,
                json,
                ok: all_ok,
            })
        }
        Command::Validate { file, json } => {
            let mut handler = open_handler(&file)?;
            done(handler.validate()?, json)
        }
        Command::Mcp => {
            if !allow_stdin {
                bail!("cannot start an MCP server from within the MCP server");
            }
            mcp::serve()?;
            // stdout is the protocol channel; end without printing anything.
            std::process::exit(0);
        }
        Command::Resident => {
            if !allow_stdin {
                bail!("cannot start resident mode from within MCP/resident");
            }
            resident::serve()?;
            std::process::exit(0);
        }
        Command::Watch { file, port } => {
            if !allow_stdin {
                bail!("watch runs as a foreground server; start it from a shell, not over MCP/resident");
            }
            watch::serve(&file, port)?;
            std::process::exit(0);
        }
        Command::Help { topic } => done(Report::Text(helptext::help_for(topic.as_deref())?), false),
        Command::External(args) => {
            let name = args.first().cloned().unwrap_or_default();
            if !allow_stdin {
                bail!("unknown command '{name}' (plugins are not available over MCP/resident)");
            }
            let exe = format!("officecli-{name}");
            match std::process::Command::new(&exe).args(&args[1..]).status() {
                Ok(status) => std::process::exit(status.code().unwrap_or(1)),
                Err(e) if e.kind() == std::io::ErrorKind::NotFound => bail!(
                    "unknown command '{name}': not a built-in and no '{exe}' plugin found on PATH (run 'officecli help')"
                ),
                Err(e) => Err(e).with_context(|| format!("cannot run plugin '{exe}'"))?,
            }
        }
    }
}

fn main() {
    // `--json` errors also go to stdout as a JSON envelope so agents can
    // parse failures uniformly.
    let wants_json = std::env::args().any(|a| a == "--json");
    match run_command(Cli::parse(), true) {
        Ok(outcome) => {
            println!("{}", outcome.report.render(outcome.json));
            if !outcome.ok {
                std::process::exit(1);
            }
        }
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
