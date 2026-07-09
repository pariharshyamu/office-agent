//! `officecli help` — the agent-oriented command guide, kept deliberately
//! close to the upstream OfficeCLI help system so agents trained on either
//! tool can drive this one.

use anyhow::{bail, Result};

const GENERAL: &str = r#"officecli (Rust port) — Office suite CLI for AI agents

Reads, creates, and edits .docx / .xlsx / .pptx with a DOM-like path system.
Paths are 1-based ('/body/p[3]' = third paragraph); --index is 0-based.
Add --json to any command for structured output ({"ok":true,"data":...}).

COMMANDS
  create <file> [--force]                       Create a blank document
  view <file> [outline|text|stats|html] [-o F]  Inspect (html = snapshot file)
  get <file> <path> [--depth N]                 Read an element
  query <file> <selector>                       CSS-like element search
  add <file> <parent> --type T [--prop k=v ...] Add an element
      [--index N | --before PATH | --after PATH]
  set <file> <path> [--prop k=v ...]            Modify properties
      [--find TEXT [--replace TEXT]]            Find/format or find/replace
  move <file> <path> [--to P] [--index N|--before P|--after P]
  swap <file> <path1> <path2>                   Exchange two elements
  remove <file> <path>                          Remove an element
  dump <file>                                   Replayable batch JSON
  batch <file> [--commands JSON|--input F|stdin] Many ops, one save
  validate <file>                               Check package structure
  mcp                                           MCP server on stdio
  help [docx|xlsx|pptx]                         Format-specific guide

QUERY SELECTORS
  paragraph[style=Normal] > run[font!=Arial]    direct-child chain
  cell[value>5000]   run[bold=true]   :contains("text")   :empty
  Operators: = != ~= (substring) >= <= > < (numeric when both numeric)

VALUE FORMATS
  Colors      FF0000, #FF0000, red, rgb(255,0,0)
  Lengths     2cm, 1in, 72pt, 96px, or raw EMU (914400 = 1 inch)
  Font sizes  24 or 24pt
  Booleans    true/false

QUICK START
  officecli create deck.pptx
  officecli add deck.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
  officecli add deck.pptx '/slide[1]' --type shape --prop text="Revenue grew 25%" \
      --prop x=2cm --prop y=5cm --prop size=24 --prop color=FFFFFF
  officecli view deck.pptx outline

  officecli create report.docx
  officecli add report.docx /body --type paragraph --prop text="Summary" --prop style=Heading1
  officecli set report.docx / --find draft --replace final

  officecli create data.xlsx
  officecli set data.xlsx /Sheet1/A1 --prop value=Name --prop bold=true
  officecli set data.xlsx /Sheet1/B2 --prop value="=SUM(B1:B1)"

Always quote paths with brackets ('/slide[1]') — shells glob-expand [].
Run 'officecli help <format>' for element types and properties."#;

const DOCX: &str = r#"officecli docx — Word documents

PATHS
  /body/p[3]           third paragraph        (aliases: paragraph, para)
  /body/p[3]/r[1]      first run in it        (alias: run)
  /body/tbl[1]         first table            (alias: table)
  /body/tbl[1]/tr[2]/tc[3]   row 2, cell 3    (aliases: row, cell/td)

ADD
  --type paragraph   props: text, style (Normal/Title/Heading1..3), align
                     (left/center/right/justify), bold, italic, underline,
                     size (pt), color, font, highlight
                     '\n' in text becomes a line break, '\t' a tab
  --type run         same text/format props; parent must be a paragraph
  --type table       props: rows, cols (default 2x2)
  --type row         parent must be a table; copies the column count
  --type break       page break
  --type image       props: src=file.png (PNG/JPEG/GIF), w/h (optional,
                     aspect kept), align; intrinsic size at 96 dpi

SET
  paragraph          text (replaces runs), style, align, plus run format
                     props applied to every run
  run                text + format props
  cell (tc)          text, plus run format props applied inside the cell

FIND / REPLACE
  officecli set doc.docx / --find draft --replace final          # whole doc
  officecli set doc.docx '/body/p[1]' --find weather --prop bold=true
                                                                 # format hits
  --prop regex=true makes --find a regular expression
  Matches work across run boundaries; case-sensitive (use '(?i)...' + regex).

VIEW  outline (headings + tables), text, stats"#;

const XLSX: &str = r#"officecli xlsx — Excel workbooks

PATHS
  /Sheet1/A1          cell (sheet by name)     also: $Sheet1:A1
  /sheet[2]/B5        cell (sheet by position)
  /Sheet1/A1:C10      range (get only)
  /Sheet1/row[5]      row 5

SET (cells are created on demand)
  value=Hello         inline string
  value=42            number (auto-detected)
  value="=SUM(A1:A9)" formula (auto-detected by leading '=')
  type=string|number|boolean|formula   forces interpretation
  bold, italic, color, size, font, fill   cell styling
  Formulas are stored uncalculated; Excel/LibreOffice recalculate on open.

SET (sheet)
  officecli set data.xlsx /Sheet1 --prop name=Budget    # rename

ADD
  --type sheet   props: name          (added at '/')
  --type row     props: values="a,b,c" and/or c1=..., c2=...
                 --index N is the 1-based row number; later rows shift down
                 and formula references are rewritten to follow

REMOVE
  /Sheet1        removes the sheet (refused for the last one)
  /Sheet1/row[5] removes the row and shifts rows up
  /Sheet1/B2     clears the cell

FIND / REPLACE
  officecli set data.xlsx / --find draft --replace final
  (string cells only; find+format is not supported for xlsx)

VIEW  outline (sheets + ranges), text (grid), stats"#;

const PPTX: &str = r#"officecli pptx — PowerPoint presentations

PATHS
  /slide[1]                     first slide
  /slide[1]/shape[2]            second shape (positional)
  /slide[1]/shape[@name=Title 1]  by name
  /slide[1]/shape[@id=5]          by stable id (preferred in workflows)

ADD
  --type slide   props: title, background (color)      (added at '/')
                 --index/--before/--after control position
  --type shape   textbox; props: text, x, y, w, h (lengths), size (pt),
                 color, bold, italic, font, align, fill, name
                 '\n' in text starts a new paragraph
  --type image   props: src=file.png, x, y, w, h (aspect kept)

SET
  slide          background=COLOR
  shape          text (replaces content), x/y/w/h, fill, name, and
                 size/color/bold/italic/font/align applied to all runs

FIND / REPLACE
  officecli set deck.pptx / --find draft --replace final
  (find+format is not supported for pptx)

VIEW  outline (slides + shapes), text, stats

TIP  shape[1] is usually the title textbox on slides created with
     --prop title=...; content shapes start at shape[2]."#;

pub fn help_for(topic: Option<&str>) -> Result<String> {
    match topic.map(|t| t.to_ascii_lowercase()) {
        None => Ok(GENERAL.to_string()),
        Some(t) => match t.as_str() {
            "docx" | "word" => Ok(DOCX.to_string()),
            "xlsx" | "excel" => Ok(XLSX.to_string()),
            "pptx" | "ppt" | "powerpoint" => Ok(PPTX.to_string()),
            other => bail!("unknown help topic '{other}' (docx/xlsx/pptx)"),
        },
    }
}
