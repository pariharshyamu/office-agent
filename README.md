# officecli (Rust)

A Rust implementation of [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) — a
command-line Office suite built for AI agents. It reads, creates, and edits
Word (`.docx`), Excel (`.xlsx`), and PowerPoint (`.pptx`) files natively:
single static binary, no Microsoft Office, no runtime dependencies.

The upstream project is ~260K lines of C#. This port reimplements its **core
L1/L2 command surface** (create / view / get / add / set / remove / batch /
validate, the DOM-path addressing system, `--prop` conventions, and JSON
output) directly on top of the OOXML package format, so an agent that knows
OfficeCLI can drive this tool with the same muscle memory.

## Build

```bash
cargo build --release
# binary at target/release/officecli
```

## Quick start

```bash
# PowerPoint
officecli create deck.pptx
officecli add deck.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
officecli add deck.pptx '/slide[1]' --type shape \
    --prop text="Revenue grew 25%" --prop x=2cm --prop y=5cm \
    --prop size=24 --prop color=FFFFFF --prop font=Arial
officecli view deck.pptx outline
# → Slide 1: Q4 Report
# →   Shape 1 [shape] Title 1: Q4 Report
# →   Shape 2 [shape] TextBox 3: Revenue grew 25%

# Word
officecli create report.docx
officecli add report.docx /body --type paragraph --prop text="Summary" --prop style=Heading1
officecli add report.docx /body --type table --prop rows=2 --prop cols=3
officecli set report.docx / --find draft --replace final
officecli set report.docx '/body/p[1]' --find Summary --prop bold=true --prop color=red

# Excel
officecli create data.xlsx
officecli set data.xlsx /Sheet1/A1 --prop value=Name --prop bold=true
officecli set data.xlsx /Sheet1/B2 --prop value=91
officecli set data.xlsx /Sheet1/B3 --prop value="=SUM(B2:B2)"
officecli view data.xlsx text
```

Add `--json` to any command for a structured envelope:

```json
{ "ok": true, "data": { "results": [ { "path": "/slide[1]/shape[2]",
  "type": "shape", "text": "Revenue grew 25%",
  "attributes": { "x": "720000", "y": "1800000", "id": "3" } } ] } }
```

Errors with `--json` also go to stdout as `{"ok":false,"error":"..."}` with
exit code 1, so agents can parse every outcome uniformly.

## The path system

Elements are addressed with 1-based, XPath-like paths:

| Format | Examples |
|--------|----------|
| docx | `/body/p[3]`, `/body/p[3]/r[1]`, `/body/tbl[1]/tr[2]/tc[3]` |
| xlsx | `/Sheet1/A1`, `$Sheet1:A1`, `/sheet[2]/B5`, `/Sheet1/A1:C10`, `/Sheet1/row[5]` |
| pptx | `/slide[1]`, `/slide[1]/shape[2]`, `/slide[1]/shape[@name=Title 1]`, `shape[@id=5]` |

`/` means the whole document (used for `add --type slide`, `add --type sheet`,
and whole-document find/replace). Friendly aliases work everywhere:
`paragraph`≡`p`, `run`≡`r`, `table`≡`tbl`, `row`≡`tr`, `cell`≡`tc`.

Always quote paths containing brackets (`'/slide[1]'`) — shells glob-expand `[]`.

## Commands

```
create <file> [--force]                          blank .docx/.xlsx/.pptx
view <file> [outline|text|stats|html] [-o FILE]  inspect (html = static snapshot)
get <file> <path> [--depth N] [--json]           read an element
query <file> <selector> [--json]                 CSS-like element search
add <file> <parent> --type T [--prop k=v ...]    add an element
    [--index N | --before PATH | --after PATH]
set <file> <path> [--prop k=v ...] [--json]      modify properties
    [--find TEXT [--replace TEXT]]               find/format or find/replace
move <file> <path> [--to P] [--index N|...]      reposition an element
swap <file> <path1> <path2>                      exchange two elements
remove <file> <path> [--json]                    delete an element
dump <file>                                      replayable batch JSON
batch <file> [--commands JSON | --input F]       many ops, one save cycle
validate <file> [--json]                         package sanity check
mcp                                              MCP server on stdio
help [docx|xlsx|pptx]                            agent-oriented guide
```

Paths are 1-based; `--index` is 0-based (array convention), except
`add --type row` on xlsx where `--index` is the 1-based row number
(matches the OOXML row index, as in upstream OfficeCLI).

### Value formats

| Type | Accepted forms |
|------|----------------|
| Colors | `FF0000`, `#FF0000`, `red`, `rgb(255,0,0)` |
| Lengths | `2cm`, `25mm`, `1in`, `72pt`, `96px`, raw EMU (`914400` = 1 inch) |
| Font sizes | `24` or `24pt` |
| Booleans | `true` / `false` |
| Text | `\n` becomes a line break (new paragraph in pptx), `\t` a tab |

### Find / replace

```bash
officecli set doc.docx / --find draft --replace final        # whole document
officecli set doc.docx '/body/p[1]' --find weather --prop bold=true   # format hits (docx)
officecli set doc.docx / --find '\d+%' --prop regex=true --replace "N%"
```

Matches work **across run boundaries** in docx and pptx; formatting matched
text (`--find` + format props without `--replace`) splits runs at match
boundaries and is docx-only. xlsx replaces within string cells.

### Query

CSS-like selectors matched against the document tree. Predicates test the
attributes shown by `get` (or the node text via `value`/`text`):

```bash
officecli query report.docx 'paragraph[style=Normal] > run[font!=Arial]'
officecli query report.docx 'run[bold=true]'
officecli query data.xlsx 'cell[value>5000]'
officecli query deck.pptx ':contains("Revenue")'
```

Operators: `=`, `!=`, `~=` (substring), `>=`, `<=`, `>`, `<` (numeric when
both sides are numbers); pseudo-classes `:contains("text")` and `:empty`;
`a > b` chains direct parent/child.

### Move / swap

```bash
officecli move report.docx '/body/p[5]' --index 0      # reorder paragraphs
officecli move deck.pptx '/slide[3]' --before '/slide[1]'
officecli move data.xlsx '/Summary' --index 0          # reorder sheets
officecli swap deck.pptx '/slide[1]' '/slide[2]'
officecli swap report.docx '/body/p[1]' '/body/p[2]'
```

### Images

```bash
officecli add report.docx /body --type image --prop src=chart.png --prop align=center
officecli add deck.pptx '/slide[1]' --type image \
    --prop src=logo.png --prop x=6in --prop y=0.5in --prop w=2in
```

PNG, JPEG, and GIF; intrinsic size is read from the file (96 dpi) and the
aspect ratio is preserved when only one of `w`/`h` is given.

### view html and dump

`view html` renders a static, self-contained HTML snapshot — styled
paragraphs and tables for docx, grids for xlsx, and absolutely-positioned
slide canvases (with backgrounds, fonts, colors) for pptx — giving agents a
render-look-fix loop without a browser engine:

```bash
officecli view deck.pptx html -o deck.html
```

`dump` emits a replayable batch-JSON approximation of the content
(paragraph/table text and styles, cell values and formulas, slides and
textboxes) that `batch` can apply to a fresh file:

```bash
officecli dump deck.pptx > ops.json
officecli create copy.pptx
officecli batch copy.pptx --input ops.json
```

### MCP server

`officecli mcp` runs a Model Context Protocol server on stdio, exposing a
single `officecli` tool whose one argument is the CLI command string —
the same design as upstream. Register it with any MCP client, e.g. for
Claude Code:

```bash
claude mcp add officecli -- /path/to/officecli mcp
```

```json
{ "mcpServers": { "officecli": { "command": "/path/to/officecli", "args": ["mcp"] } } }
```

### Batch

```bash
echo '[
  {"command":"set","path":"/Sheet1/A1","props":{"value":"Name","bold":"true"}},
  {"command":"set","path":"/Sheet1/B1","props":{"value":"Score","bold":"true"}}
]' | officecli batch data.xlsx --json
```

Items support `command` (or `op`), `path`/`parent`, `type`, `props`, `index`,
`before`, `after`, `find`, `replace`, `depth`, `mode`. Continues on error by
default (exit 1 if anything failed); `--stop-on-error` aborts at the first
failure.

## What's implemented per format

**Word (.docx)** — paragraphs (text, style Normal/Title/Heading1–3, align),
runs (bold/italic/underline/size/color/font/highlight), tables (create,
add row, set cell text + formatting), page breaks, insert before/after/at
index, whole-scope find/replace and find+format with run splitting,
outline/text/stats views.

**Excel (.xlsx)** — cells created on demand via `set` (strings, numbers,
booleans, formulas auto-detected by `=`), shared-string-aware reads,
cell styling (bold/italic/color/size/font/fill) through a styles-table
manager, sheets (add/rename/remove), rows (insert with shift, remove with
shift, `values=` / `cN=` fills), ranges on `get`, `$Sheet:A1` addressing,
find/replace over string cells, grid/outline/stats views. Formulas are
stored uncalculated with `fullCalcOnLoad` so Excel/LibreOffice recalculate
on open.

**PowerPoint (.pptx)** — slides (add with `title`/`background`, position
with `--index/--before/--after`, remove with full relationship cleanup),
textbox shapes (text, x/y/w/h in any length unit, size/color/bold/italic/
font/align/fill/name), shape addressing by position, `@name=`, or `@id=`,
slide background, find/replace across slides, outline/text/stats views.

Blank documents created by `create` open without repair prompts in Word,
Excel, PowerPoint, and LibreOffice, and parse with `python-docx`,
`openpyxl`, and `python-pptx` (used as reference validators in CI/tests).

## Known limitations vs upstream OfficeCLI

This is a focused port, not a feature-complete clone. Not (yet) implemented:
PNG rendering (`view screenshot`) and the live `watch` preview server
(`view html` provides static snapshots), resident mode, charts, pivot
tables, comments/footnotes/TOC/fields, animations/transitions, plugins,
column-shift formula rewriting (row shifts are rewritten), and `dump` is a
content-level approximation rather than a full-fidelity round-trip. The
architecture (lossless XML DOM over the zip package, one handler per
format behind a common trait) is designed so these can be added
incrementally.

## Architecture

```
src/
  main.rs       clap CLI, dispatch (shared with the MCP server)
  handler.rs    Handler trait implemented by the three formats
  pkg.rs        OOXML zip package (parts held in memory, saved atomically)
  xml.rs        minimal lossless XML DOM (quick-xml based, order-preserving)
  path.rs       /slide[1]/shape[@name=X] path parser + A1 cell refs
  props.rs      --prop parsing, lengths→EMU, colors, alignment
  out.rs        text + JSON report rendering
  query.rs      CSS-like selector parser + NodeInfo-tree matcher
  media.rs      image sniffing (PNG/JPEG/GIF) + package bookkeeping
  html.rs       escaping + page shell for `view html`
  mcp.rs        Model Context Protocol server (stdio JSON-RPC)
  templates.rs  minimal valid blank .docx/.xlsx/.pptx packages
  docx.rs       Word handler (incl. cross-run find/replace + run splitting)
  xlsx.rs       Excel handler (incl. shared strings, styles manager,
                formula-reference rewriting on row shifts)
  pptx.rs       PowerPoint handler (slides/shapes/images/relationships)
  batch.rs      JSON batch runner
  helptext.rs   `officecli help` agent guide
```

Only untouched parts are rewritten byte-identically: the DOM preserves
qualified names, attribute order, and text verbatim, so editing one
paragraph never reshuffles the rest of the document.

## License

Apache-2.0, same as the upstream project this ports.
