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
  view <file> [outline|text|stats|html|screenshot|comments] [-o F]
                                                Inspect (screenshot = PNG per
                                                page/sheet/slide, needs -o)
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
  watch <file> [--port N]                       Live-reloading HTML preview
  resident                                      Command loop (line in, JSON out)
  mcp                                           MCP server on stdio
  help [docx|xlsx|pptx]                         Format-specific guide

Unknown subcommands run `officecli-<name>` plugins found on PATH.

QUERY SELECTORS
  paragraph[style=Normal] > run[font!=Arial]    direct-child chain
  cell[value>5000]   run[bold=true]   :contains("text")   :empty
  Operators: = != ~= (substring) >= <= > < (numeric when both numeric)

VALUE FORMATS
  Colors      FF0000, #FF0000, red, rgb(255,0,0)
  Lengths     2cm, 1in, 72pt, 96px, or raw EMU (914400 = 1 inch)
  Font sizes  24 or 24pt
  Durations   500ms, 1.5s, or milliseconds
  Booleans    true/false

QUICK START
  officecli create deck.pptx
  officecli add deck.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
  officecli add deck.pptx '/slide[1]' --type shape --prop text="Revenue grew 25%" \
      --prop x=2cm --prop y=5cm --prop size=24 --prop color=FFFFFF
  officecli add deck.pptx '/slide[1]' --type chart --prop kind=pie \
      --prop categories="East,West" --prop values="60,40"
  officecli set deck.pptx '/slide[1]' --prop transition=fade
  officecli view deck.pptx screenshot -o deck.png

  officecli create report.docx
  officecli add report.docx /body --type paragraph --prop text="Summary" --prop style=Heading1
  officecli add report.docx '/body/p[1]' --type comment --prop text="Check" --prop author=Bot
  officecli set report.docx / --find draft --replace final

  officecli create data.xlsx
  officecli set data.xlsx /Sheet1/A1 --prop value=Name --prop bold=true
  officecli add data.xlsx /Sheet1 --type chart --prop data=A1:B9 --prop kind=column
  officecli add data.xlsx / --type pivot --prop source=A1:C9 --prop rows=Region \
      --prop values=Sales --prop agg=sum

Always quote paths with brackets ('/slide[1]') — shells glob-expand [].
Run 'officecli help <format>' for element types and properties."#;

const DOCX: &str = r#"officecli docx — Word documents

PATHS
  /body/p[3]           third paragraph        (aliases: paragraph, para)
  /body/p[3]/r[1]      first run in it        (alias: run)
  /body/tbl[1]         first table            (alias: table)
  /body/tbl[1]/tr[2]/tc[3]   row 2, cell 3    (aliases: row, cell/td)
  /comment[1]          comment by id (remove only)

ADD
  --type paragraph   props: text, style (Normal/Title/Heading1..3), align
                     (left/center/right/justify), bold, italic, underline,
                     size (pt), color, font, highlight
                     '\n' in text becomes a line break, '\t' a tab
  --type run         same text/format props; parent must be a paragraph
  --type table       props: rows, cols (default 2x2)
  --type row         parent must be a table; copies the column count
  --type break       page break
  --type image       props: src=file.png (PNG/JPEG/GIF) or srcdata=BASE64,
                     w/h (optional, aspect kept), align
  --type toc         table of contents field (parent /body); props: levels
                     ("1-3"); Word populates it on open/update
  --type field       props: kind=page|numpages|date|time|filename|author
                     or code="..." (parent = a paragraph)
  --type comment     props: text, author; attaches to a paragraph
  --type footnote    props: text; adds superscript reference + note

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

VIEW  outline (headings + tables), text, stats, html, comments,
      screenshot (-o page.png; multi-page → page-1.png, ...)
REMOVE  /comment[N] deletes a comment and its body markers"#;

const XLSX: &str = r#"officecli xlsx — Excel workbooks

PATHS
  /Sheet1/A1          cell (sheet by name)     also: $Sheet1:A1
  /sheet[2]/B5        cell (sheet by position)
  /Sheet1/A1:C10      range (get only)
  /Sheet1/row[5]      row 5
  /Sheet1/B           column B  (also /Sheet1/col[2]; remove only)

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
  --type column  props: at=B (or --index N, 1-based), values="a,b,c" (fills
                 down from row 1), r1=..., r2=... ; cells shift right and
                 formula references follow
  --type chart   props: data=A1:C9 (first col = categories, first row =
                 headers), kind=column|bar|line|pie, title, at=E2 (anchor
                 cell), w/h (in cells); references live cells
  --type pivot   props: source=A1:C9, rows=Header, values=Header,
                 agg=sum|count|avg|min|max, name; writes a computed group-by
                 summary to a new sheet (static table, not an interactive
                 PivotTable)

REMOVE
  /Sheet1        removes the sheet (refused for the last one)
  /Sheet1/row[5] removes the row and shifts rows up
  /Sheet1/B      removes column B and shifts columns left (formulas follow)
  /Sheet1/B2     clears the cell

FIND / REPLACE
  officecli set data.xlsx / --find draft --replace final
  (string cells only; find+format is not supported for xlsx)

VIEW  outline (sheets + ranges), text (grid), stats, html,
      screenshot (-o grid.png; one PNG per sheet)"#;

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
  --type image   props: src=file.png or srcdata=BASE64, x, y, w, h
  --type chart   props: kind=column|bar|line|pie, categories="Q1,Q2",
                 values="10,20", series=Name, values2=/series2= for more
                 series, title, x/y/w/h; data is embedded in the chart
                 (PowerPoint renders it; Edit Data needs a linked workbook)

SET
  slide          background=COLOR
                 transition=fade|cut|push|wipe|dissolve|circle|diamond|
                 plus|wedge|wheel|zoom|cover|pull|split|blinds|checker|
                 comb|strips|newsflash|random|none
                 [direction=left|right|up|down|horizontal|vertical|in|out]
                 [speed=slow|medium|fast or a duration] [advance=5s]
  shape          text (replaces content), x/y/w/h, fill, name, and
                 size/color/bold/italic/font/align applied to all runs
                 animation=appear|fade|wipe (click-triggered entrance;
                 [duration=500ms] [delay=0ms])

FIND / REPLACE
  officecli set deck.pptx / --find draft --replace final
  (find+format is not supported for pptx)

VIEW  outline (slides + shapes), text, stats, html,
      screenshot (-o deck.png; one PNG per slide)

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
