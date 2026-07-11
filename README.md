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

## Install

Linux / macOS:

```bash
curl -fsSL https://github.com/pariharshyamu/office-agent/releases/latest/download/install.sh | sh
```

Windows (PowerShell):

```powershell
irm https://github.com/pariharshyamu/office-agent/releases/latest/download/install.ps1 | iex
```

The installers pick the right prebuilt binary (Linux x64/arm64, macOS
Apple-silicon/Intel, Windows x64/ARM64 — both Windows builds use the MSVC
toolchain), install it to `~/.local/bin` (Windows:
`%LOCALAPPDATA%\Programs\officecli`), and print PATH instructions if
needed. Pin a version with `OFFICECLI_VERSION=v0.2.0`; change the
directory with `OFFICECLI_INSTALL=...`.

Or build from source:

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
view <file> [outline|text|stats|html|screenshot|comments] [-o FILE]
                                                 inspect (screenshot = PNGs)
get <file> <path> [--depth N] [--json]           read an element
query <file> <selector> [--json]                 CSS-like element search
add <file> <parent> --type T [--prop k=v ...]    add an element
    [--index N | --before PATH | --after PATH]
set <file> <path> [--prop k=v ...] [--json]      modify properties
    [--find TEXT [--replace TEXT]]               find/format or find/replace
move <file> <path> [--to P] [--index N|...]      reposition an element
swap <file> <path1> <path2>                      exchange two elements
copy <file> <path> [--index N]                   duplicate slide/sheet/row/
                                                 paragraph/table/shape
remove <file> <path> [--json]                    delete an element
export <file> [--range R] [--sheet S] [-o F]     cell range as CSV (xlsx)
dump <file>                                      replayable batch JSON
batch <file> [--commands JSON | --input F]       many ops, one save cycle
validate <file> [--json]                         package sanity check
watch <file> [--port N]                          live-reloading HTML preview
resident                                         command loop (line in, JSON line out)
mcp                                              MCP server on stdio
help [docx|xlsx|pptx]                            agent-oriented guide
<anything else>                                  officecli-<name> plugin from PATH
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
| Dates | `2026-07-10`, `2026-07-10 14:30`, `14:30:00` (xlsx: real date cells) |
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
aspect ratio is preserved when only one of `w`/`h` is given. Bytes can be
passed inline with `--prop srcdata=BASE64` (used by `dump` for round-trips).

### Lists

```bash
officecli add report.docx /body --type list --prop items='First\nSecond\n\tNested detail'
officecli add report.docx /body --type list --prop kind=number --prop items='Step 1\nStep 2'
officecli set report.docx '/body/p[4]' --prop list=bullet          # or list=none
officecli add deck.pptx '/slide[1]' --type shape --prop list=bullet \
    --prop text='Point one\n\tSub-point\nPoint two'
```

`\n` separates items; leading `\t`s select deeper levels (docx: real
`numbering.xml` definitions; pptx: `buChar`/`buAutoNum` with hanging
indents). Views, screenshots, and `dump` render and replay markers.

### Dates, number formats, CSV (xlsx)

```bash
officecli set data.xlsx /Sheet1/A1 --prop value=2026-07-10        # real date cell
officecli set data.xlsx /Sheet1/B1 --prop value=0.185 --prop format=0.00%
officecli set data.xlsx /Sheet1/C1 --prop value=1234.5 --prop format=currency
officecli add data.xlsx /Sheet1 --type csv --prop src=input.csv --prop at=A1
officecli export data.xlsx --range 'Sheet1!A1:C99' -o out.csv
```

ISO dates/times auto-detect into serial numbers with date formats (openpyxl
reads them back as `datetime` objects); `type=string` opts out. `format=`
accepts `date`/`datetime`/`time`/`percent`/`currency`/`integer`/`0.00`/raw
codes. Date cells read back as ISO strings with `type=date`. CSV import
auto-types numbers and dates; export quotes RFC 4180-style.

### Hyperlinks

```bash
officecli add report.docx '/body/p[1]' --type hyperlink --prop url=https://example.com --prop text=docs
officecli set data.xlsx /Sheet1/A1 --prop value=Home --prop url=https://example.com
officecli add deck.pptx '/slide[1]' --type shape --prop text="Visit us" --prop url=https://example.com
```

### Speaker notes (pptx)

```bash
officecli add deck.pptx / --type slide --prop title=Intro --prop notes="Keep it under 2 minutes"
officecli set deck.pptx '/slide[1]' --prop notes="Revised talk track"
officecli view deck.pptx notes
```

Notes survive `copy`, replay through `dump`, and read back via python-pptx.

### Copy

```bash
officecli copy deck.pptx '/slide[1]'          # duplicate (content, background, notes)
officecli copy deck.pptx '/slide[1]/shape[2]'
officecli copy data.xlsx /Sheet1              # → "Sheet1 (2)"
officecli copy data.xlsx '/Sheet1/row[2]'     # duplicate below, shifts rows
officecli copy report.docx '/body/p[1]'
```

### Theme colors and Word content controls

Documents authored in Office use theme references (`schemeClr accent1`,
tints, `lumMod`/`lumOff`) rather than literal colors; officecli resolves
them through `theme1.xml` in `get`/`view html`/`screenshot`/`dump`, so
corporate decks keep their palette. Word content controls (`w:sdt`) are
transparent: paragraphs inside them address as plain `/body/p[N]` paths and
the wrappers are preserved on save.

### Charts

```bash
# xlsx: chart over live cells (first column = categories, first row = headers)
officecli add data.xlsx /Sheet1 --type chart \
    --prop data=A1:C9 --prop kind=column --prop title="Sales" --prop at=E2

# pptx: chart with embedded data
officecli add deck.pptx '/slide[1]' --type chart --prop kind=pie \
    --prop categories="East,West,North" --prop values="40,35,25" \
    --prop series="Share" --prop title="Regional share"
```

Kinds: `column`, `bar`, `line`, `pie`. xlsx charts reference the cells (they
update when the data changes); pptx charts carry their data as cached
literals (add `values2=`/`series2=`, ... for more series). Charts parse
cleanly in openpyxl/python-pptx and render in Office and LibreOffice.

### Pivot summaries (xlsx)

```bash
officecli add data.xlsx / --type pivot --prop source=A1:C99 \
    --prop rows=Region --prop values=Sales --prop agg=sum
```

Writes a computed group-by summary (`sum`/`count`/`avg`/`min`/`max` plus a
Grand Total) to a new sheet. It is a static table, not a native interactive
PivotTable — agents usually want the numbers, not the UI widget.

### Columns (xlsx)

```bash
officecli add data.xlsx /Sheet1 --type column --prop at=B --prop r1=Header
officecli remove data.xlsx /Sheet1/B
```

Cells shift and **formula references are rewritten** on both row and column
inserts/removals, across all sheets (`Sheet2!B5`-style references included).

### Comments, footnotes, TOC, fields (docx)

```bash
officecli add report.docx '/body/p[2]' --type comment --prop text="Check this" --prop author=Reviewer
officecli view report.docx comments
officecli remove report.docx '/comment[1]'
officecli add report.docx '/body/p[2]' --type footnote --prop text="Source: ..."
officecli add report.docx /body --type toc --index 0
officecli add report.docx '/body/p[5]' --type field --prop kind=page
```

The TOC is inserted as a dirty field with `updateFields` set, so Word
populates it on open. Fields support `page`, `numpages`, `date`, `time`,
`filename`, `author`, or raw `code="..."`.

### Transitions and animations (pptx)

```bash
officecli set deck.pptx '/slide[1]' --prop transition=push --prop direction=left \
    --prop speed=fast --prop advance=5s
officecli set deck.pptx '/slide[1]/shape[2]' --prop animation=fade --prop duration=750ms
```

Transitions: `fade`, `cut`, `push`, `wipe`, `dissolve`, `circle`, `diamond`,
`plus`, `wedge`, `wheel`, `zoom`, `cover`, `pull`, `split`, `blinds`,
`checker`, `comb`, `strips`, `newsflash`, `random`, `none`. Animations are
click-triggered entrance effects (`appear`, `fade`, `wipe`) built as a
standard `p:timing` tree.

### view html, view screenshot, watch, and dump

`view html` renders a static, self-contained HTML snapshot — styled
paragraphs and tables for docx, grids for xlsx, and absolutely-positioned
slide canvases (with backgrounds, fonts, colors) for pptx:

```bash
officecli view deck.pptx html -o deck.html
```

`view screenshot` renders real PNGs — no browser engine involved; text is
rasterized with an embedded DejaVu Sans (regular + bold). One PNG per
slide/sheet/page (`out.png`, or `out-1.png`, `out-2.png`, ... for multiple),
giving agents a genuine render-look-fix loop:

```bash
officecli view deck.pptx screenshot -o deck.png
```

The rendering is an approximation (positions, fills, text size/color/bold,
PNG images; JPEG/GIF and charts appear as placeholders), designed to make
layout problems visible rather than to be print-accurate.

`watch` serves the HTML view on localhost and auto-reloads the browser
whenever the file changes on disk — edit with officecli in one terminal,
see the result live:

```bash
officecli watch deck.pptx --port 8787
```

`dump` emits replayable batch-JSON (paragraph/run text and formatting,
images with base64 bytes, cell values/formulas/styles, slides, textboxes,
pictures, transitions) that `batch` can apply to a fresh file:

```bash
officecli dump deck.pptx > ops.json
officecli create copy.pptx
officecli batch copy.pptx --input ops.json
```

### Resident mode and plugins

`officecli resident` keeps one process alive for agents issuing many
commands: each stdin line is a command string, each response is one compact
JSON line (`{"ok":true,"data":...}`); `exit` or EOF ends the loop.

Unknown subcommands dispatch to plugins: `officecli foo args...` runs
`officecli-foo args...` if such an executable exists on PATH.

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
runs (bold/italic/underline/size/color/font/highlight), bulleted/numbered
lists with nesting, hyperlinks, tables (create, add row, set cell text +
formatting), page breaks, images, comments (add/list/remove), footnotes,
TOC and fields, copy, transparent `w:sdt` content controls, insert
before/after/at index, whole-scope find/replace and find+format with run
splitting, outline/text/stats/html/comments/screenshot views.

**Excel (.xlsx)** — cells created on demand via `set` (strings, numbers,
booleans, real dates/times, formulas auto-detected by `=`), number formats
(percent/currency/custom codes), hyperlinks, CSV import/export,
shared-string-aware reads, cell styling (bold/italic/underline/color/size/
font/fill) through a styles-table manager, sheets (add/rename/remove/copy),
rows and columns (insert/remove/copy with shifts and workbook-wide
formula-reference rewriting), charts over live ranges, computed pivot
summaries, ranges on `get`, `$Sheet:A1` addressing, find/replace over
string cells, grid/outline/stats/html/screenshot views. Formulas are stored
uncalculated with `fullCalcOnLoad` so Excel/LibreOffice recalculate on open.

**PowerPoint (.pptx)** — slides (add with `title`/`background`/`notes`,
position with `--index/--before/--after`, copy with notes, remove with full
relationship cleanup), textbox shapes (text, x/y/w/h in any length unit,
size/color/bold/italic/font/align/fill/name/url, bullet/numbered lists with
levels), speaker notes, images, charts, slide transitions, entrance
animations, theme-color resolution (schemeClr + lumMod/lumOff), shape
addressing by position, `@name=`, or `@id=`, slide background, find/replace
across slides, outline/text/stats/html/notes/screenshot views.

Blank documents created by `create` open without repair prompts in Word,
Excel, PowerPoint, and LibreOffice, and parse with `python-docx`,
`openpyxl`, and `python-pptx` (used as reference validators in CI/tests).

## Known limitations vs upstream OfficeCLI

This is a focused port, not a feature-complete clone. Honest edges of the
implemented features: `view screenshot` is an approximation (PNG images
composite; JPEG/GIF and charts render as placeholders; no italics or
justified text); pivots are computed summary tables, not native interactive
PivotTables; pptx charts embed cached data without a linked workbook, so
PowerPoint's "Edit Data" won't open a sheet; animations cover click-triggered
entrance effects only; theme-color transforms (lumMod/lumOff/tint/shade) are
per-channel approximations of Office's HSL math; formula evaluation is not
performed (values compute on open); `dump` is a high-fidelity content
replay, not a byte-identical round-trip. Not implemented: scatter charts,
docx charts, xlsx cell comments, headers/footers, formula evaluation, and
PowerPoint motion-path animations. The architecture (lossless XML DOM over
the zip package, one handler per format behind a common trait) is designed
so these can be added incrementally.

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
  render.rs     PNG canvas (tiny-skia + ab_glyph + embedded DejaVu Sans)
  chart.rs      DrawingML chartSpace builder (xlsx + pptx charts)
  mcp.rs        Model Context Protocol server (stdio JSON-RPC)
  resident.rs   long-lived command loop (line in, JSON line out)
  watch.rs      live-preview HTTP server (std-only, mtime polling)
  templates.rs  minimal valid blank .docx/.xlsx/.pptx packages
  docx.rs       Word handler (incl. cross-run find/replace + run splitting,
                comments/footnotes/TOC/fields, page renderer)
  xlsx.rs       Excel handler (incl. shared strings, styles manager,
                row/column shifts with formula-reference rewriting,
                charts, pivot summaries, grid renderer)
  pptx.rs       PowerPoint handler (slides/shapes/images/charts/
                transitions/animations, slide renderer)
  batch.rs      JSON batch runner
  helptext.rs   `officecli help` agent guide
```

The embedded DejaVu fonts are under the Bitstream Vera license
(`assets/DEJAVU-LICENSE`).

Only untouched parts are rewritten byte-identically: the DOM preserves
qualified names, attribute order, and text verbatim, so editing one
paragraph never reshuffles the rest of the document.

## License

Apache-2.0, same as the upstream project this ports.
