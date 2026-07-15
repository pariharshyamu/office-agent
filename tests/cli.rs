//! End-to-end tests: drive the compiled binary the way an AI agent would.

use std::path::{Path, PathBuf};
use std::process::{Command, Output};

fn bin() -> &'static str {
    env!("CARGO_BIN_EXE_officecli")
}

fn run(dir: &Path, args: &[&str]) -> Output {
    Command::new(bin())
        .current_dir(dir)
        .args(args)
        .output()
        .expect("failed to run officecli")
}

fn ok(dir: &Path, args: &[&str]) -> String {
    let out = run(dir, args);
    assert!(
        out.status.success(),
        "command {:?} failed:\nstdout: {}\nstderr: {}",
        args,
        String::from_utf8_lossy(&out.stdout),
        String::from_utf8_lossy(&out.stderr)
    );
    String::from_utf8_lossy(&out.stdout).into_owned()
}

fn fails(dir: &Path, args: &[&str]) -> String {
    let out = run(dir, args);
    assert!(
        !out.status.success(),
        "command {:?} unexpectedly succeeded: {}",
        args,
        String::from_utf8_lossy(&out.stdout)
    );
    format!(
        "{}{}",
        String::from_utf8_lossy(&out.stdout),
        String::from_utf8_lossy(&out.stderr)
    )
}

fn temp_dir(name: &str) -> PathBuf {
    let dir = std::env::temp_dir().join(format!("officecli-test-{name}-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    dir
}

#[test]
fn docx_full_workflow() {
    let dir = temp_dir("docx");
    ok(&dir, &["create", "report.docx"]);
    ok(&dir, &[
        "add", "report.docx", "/body", "--type", "paragraph",
        "--prop", "text=Executive Summary", "--prop", "style=Heading1",
    ]);
    ok(&dir, &[
        "add", "report.docx", "/body", "--type", "paragraph",
        "--prop", "text=This draft grew 25% and covers Q4.",
    ]);
    ok(&dir, &[
        "add", "report.docx", "/body", "--type", "table",
        "--prop", "rows=2", "--prop", "cols=2",
    ]);
    ok(&dir, &[
        "set", "report.docx", "/body/tbl[1]/tr[1]/tc[1]",
        "--prop", "text=Region", "--prop", "bold=true",
    ]);

    // Whole-document replace.
    let replaced = ok(&dir, &["set", "report.docx", "/", "--find", "draft", "--replace", "final"]);
    assert!(replaced.contains("matched: 1"), "{replaced}");

    // Find + format splits runs.
    ok(&dir, &[
        "set", "report.docx", "/body/p[2]", "--find", "25%",
        "--prop", "bold=true", "--prop", "color=red",
    ]);
    let json = ok(&dir, &["get", "report.docx", "/body/p[2]", "--depth", "1", "--json"]);
    let v: serde_json::Value = serde_json::from_str(&json).unwrap();
    let runs = &v["data"]["results"][0]["children"];
    assert_eq!(runs.as_array().unwrap().len(), 3, "{json}");
    assert_eq!(runs[1]["text"], "25%");
    assert_eq!(runs[1]["attributes"]["bold"], "true");
    assert_eq!(runs[1]["attributes"]["color"], "FF0000");

    let text = ok(&dir, &["view", "report.docx", "text"]);
    assert!(text.contains("This final grew 25% and covers Q4."), "{text}");
    assert!(text.contains("Region"), "{text}");

    let outline = ok(&dir, &["view", "report.docx", "outline"]);
    assert!(outline.contains("Heading1: Executive Summary"), "{outline}");
    assert!(outline.contains("[table 2x2]"), "{outline}");

    // Insert before an anchor.
    ok(&dir, &[
        "add", "report.docx", "/body", "--type", "paragraph",
        "--prop", "text=Preamble", "--before", "/body/p[1]",
    ]);
    let first = ok(&dir, &["get", "report.docx", "/body/p[1]"]);
    assert!(first.contains("Preamble"), "{first}");

    // Remove it again.
    ok(&dir, &["remove", "report.docx", "/body/p[1]"]);
    let first = ok(&dir, &["get", "report.docx", "/body/p[1]"]);
    assert!(first.contains("Executive Summary"), "{first}");

    assert!(ok(&dir, &["validate", "report.docx"]).contains("valid"));

    // Refuse to overwrite without --force.
    fails(&dir, &["create", "report.docx"]);
    ok(&dir, &["create", "report.docx", "--force"]);
}

#[test]
fn xlsx_full_workflow() {
    let dir = temp_dir("xlsx");
    ok(&dir, &["create", "data.xlsx"]);
    ok(&dir, &["set", "data.xlsx", "/Sheet1/A1", "--prop", "value=Name", "--prop", "bold=true"]);
    ok(&dir, &["set", "data.xlsx", "/Sheet1/B1", "--prop", "value=Score"]);
    ok(&dir, &["set", "data.xlsx", "/Sheet1/A2", "--prop", "value=Alice"]);
    ok(&dir, &["set", "data.xlsx", "/Sheet1/B2", "--prop", "value=91"]);
    ok(&dir, &["set", "data.xlsx", "/Sheet1/B3", "--prop", "value==SUM(B2:B2)"]);

    let cell = ok(&dir, &["get", "data.xlsx", "/Sheet1/B2", "--json"]);
    let v: serde_json::Value = serde_json::from_str(&cell).unwrap();
    assert_eq!(v["data"]["results"][0]["text"], "91");
    assert_eq!(v["data"]["results"][0]["attributes"]["type"], "number");

    let formula = ok(&dir, &["get", "data.xlsx", "/Sheet1/B3", "--json"]);
    let v: serde_json::Value = serde_json::from_str(&formula).unwrap();
    assert_eq!(v["data"]["results"][0]["attributes"]["formula"], "=SUM(B2:B2)");

    // $Sheet:A1 addressing.
    let dollar = ok(&dir, &["get", "data.xlsx", "$Sheet1:A2"]);
    assert!(dollar.contains("Alice"), "{dollar}");

    // Sheets: add, rename, remove.
    ok(&dir, &["add", "data.xlsx", "/", "--type", "sheet", "--prop", "name=Summary"]);
    ok(&dir, &["set", "data.xlsx", "/Summary/A1", "--prop", "value=hi"]);
    ok(&dir, &["set", "data.xlsx", "/Summary", "--prop", "name=Overview"]);
    let outline = ok(&dir, &["view", "data.xlsx", "outline"]);
    assert!(outline.contains("Overview"), "{outline}");
    ok(&dir, &["remove", "data.xlsx", "/Overview"]);
    fails(&dir, &["get", "data.xlsx", "/Overview/A1"]);

    // Row insert with shift.
    ok(&dir, &[
        "add", "data.xlsx", "/Sheet1", "--type", "row",
        "--index", "2", "--prop", "values=Zed,99",
    ]);
    let grid = ok(&dir, &["view", "data.xlsx", "text"]);
    assert!(grid.contains("2\tZed\t99"), "{grid}");
    assert!(grid.contains("3\tAlice\t91"), "{grid}");

    // Row remove shifts back up.
    ok(&dir, &["remove", "data.xlsx", "/Sheet1/row[2]"]);
    let grid = ok(&dir, &["view", "data.xlsx", "text"]);
    assert!(grid.contains("2\tAlice\t91"), "{grid}");

    // Find/replace.
    let replaced = ok(&dir, &["set", "data.xlsx", "/", "--find", "Alice", "--replace", "Alicia"]);
    assert!(replaced.contains("matched: 1"), "{replaced}");
    assert!(ok(&dir, &["get", "data.xlsx", "/Sheet1/A2"]).contains("Alicia"));

    // Cannot remove the last sheet.
    fails(&dir, &["remove", "data.xlsx", "/Sheet1"]);
}

#[test]
fn pptx_full_workflow() {
    let dir = temp_dir("pptx");
    ok(&dir, &["create", "deck.pptx"]);
    ok(&dir, &[
        "add", "deck.pptx", "/", "--type", "slide",
        "--prop", "title=Q4 Report", "--prop", "background=1A1A2E",
    ]);
    ok(&dir, &[
        "add", "deck.pptx", "/slide[1]", "--type", "shape",
        "--prop", "text=Revenue grew 25%", "--prop", "x=2cm", "--prop", "y=5cm",
        "--prop", "size=24", "--prop", "color=FFFFFF", "--prop", "font=Arial",
    ]);
    ok(&dir, &["add", "deck.pptx", "/", "--type", "slide", "--prop", "title=Roadmap draft"]);

    let outline = ok(&dir, &["view", "deck.pptx", "outline"]);
    assert!(outline.contains("Slide 1: Q4 Report"), "{outline}");
    assert!(outline.contains("Slide 2: Roadmap draft"), "{outline}");

    // Shape lookup by name; geometry in EMU (2cm = 720000).
    let shape = ok(&dir, &["get", "deck.pptx", "/slide[1]/shape[2]", "--json"]);
    let v: serde_json::Value = serde_json::from_str(&shape).unwrap();
    assert_eq!(v["data"]["results"][0]["attributes"]["x"], "720000");
    assert_eq!(v["data"]["results"][0]["text"], "Revenue grew 25%");

    let by_name = ok(&dir, &["get", "deck.pptx", "/slide[1]/shape[@name=Title 1]"]);
    assert!(by_name.contains("Q4 Report"), "{by_name}");

    // Replace across all slides.
    let replaced = ok(&dir, &["set", "deck.pptx", "/", "--find", "draft", "--replace", "final"]);
    assert!(replaced.contains("matched: 1"), "{replaced}");
    assert!(ok(&dir, &["view", "deck.pptx", "text"]).contains("Roadmap final"));

    // Move/resize + retext.
    ok(&dir, &[
        "set", "deck.pptx", "/slide[1]/shape[2]",
        "--prop", "x=1in", "--prop", "text=Revenue grew 30%",
    ]);
    let shape = ok(&dir, &["get", "deck.pptx", "/slide[1]/shape[2]", "--json"]);
    let v: serde_json::Value = serde_json::from_str(&shape).unwrap();
    assert_eq!(v["data"]["results"][0]["attributes"]["x"], "914400");
    assert_eq!(v["data"]["results"][0]["text"], "Revenue grew 30%");

    // Slide insertion order + removal.
    ok(&dir, &[
        "add", "deck.pptx", "/", "--type", "slide",
        "--prop", "title=Agenda", "--before", "/slide[1]",
    ]);
    let outline = ok(&dir, &["view", "deck.pptx", "outline"]);
    assert!(outline.contains("Slide 1: Agenda"), "{outline}");
    ok(&dir, &["remove", "deck.pptx", "/slide[1]"]);
    let outline = ok(&dir, &["view", "deck.pptx", "outline"]);
    assert!(outline.contains("Slide 1: Q4 Report"), "{outline}");

    // Remove a shape.
    ok(&dir, &["remove", "deck.pptx", "/slide[1]/shape[2]"]);
    let info = ok(&dir, &["get", "deck.pptx", "/slide[1]"]);
    assert!(info.contains("shapes=1"), "{info}");

    assert!(ok(&dir, &["validate", "deck.pptx"]).contains("valid"));
}

#[test]
fn batch_and_error_envelope() {
    let dir = temp_dir("batch");
    ok(&dir, &["create", "data.xlsx"]);
    let out = ok(&dir, &[
        "batch", "data.xlsx", "--json", "--commands",
        r#"[{"command":"set","path":"/Sheet1/A1","props":{"value":"Name","bold":"true"}},
            {"op":"set","path":"/Sheet1/A2","props":{"value":42}},
            {"command":"get","path":"/Sheet1/A2"}]"#,
    ]);
    let v: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!(v["data"]["succeeded"], 3);

    // Continue-on-error: exit 1 but later ops applied.
    let out = run(&dir, &[
        "batch", "data.xlsx", "--commands",
        r#"[{"op":"set","path":"/Nope/A1","props":{"value":"x"}},
            {"op":"set","path":"/Sheet1/B1","props":{"value":"ok"}}]"#,
    ]);
    assert!(!out.status.success());
    assert!(ok(&dir, &["get", "data.xlsx", "/Sheet1/B1"]).contains("ok"));

    // JSON error envelope on stdout.
    let err = run(&dir, &["get", "data.xlsx", "/Missing/A1", "--json"]);
    assert!(!err.status.success());
    let v: serde_json::Value =
        serde_json::from_str(&String::from_utf8_lossy(&err.stdout)).unwrap();
    assert_eq!(v["ok"], false);
    assert!(v["error"].as_str().unwrap().contains("Missing"));
}

#[test]
fn files_are_valid_zip_packages() {
    let dir = temp_dir("pkgcheck");
    for name in ["a.docx", "a.xlsx", "a.pptx"] {
        ok(&dir, &["create", name]);
        let bytes = std::fs::read(dir.join(name)).unwrap();
        // Zip local-file-header magic.
        assert_eq!(&bytes[..4], b"PK\x03\x04", "{name} is not a zip");
    }
    fails(&dir, &["create", "a.txt"]);
}

#[test]
fn query_selectors() {
    let dir = temp_dir("query");
    ok(&dir, &["create", "doc.docx"]);
    ok(&dir, &[
        "add", "doc.docx", "/body", "--type", "paragraph",
        "--prop", "text=Bold intro", "--prop", "bold=true",
    ]);
    ok(&dir, &["add", "doc.docx", "/body", "--type", "paragraph", "--prop", "text=Plain body"]);

    let hits = ok(&dir, &["query", "doc.docx", "run[bold=true]"]);
    assert!(hits.contains("Bold intro"), "{hits}");
    assert!(!hits.contains("Plain body"), "{hits}");

    let hits = ok(&dir, &["query", "doc.docx", r#":contains("Plain")"#]);
    assert!(hits.contains("Plain body"), "{hits}");

    ok(&dir, &["create", "d.xlsx"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/A1", "--prop", "value=150"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/A2", "--prop", "value=50"]);
    let hits = ok(&dir, &["query", "d.xlsx", "cell[value>100]"]);
    assert!(hits.contains("150") && !hits.contains("\"50\""), "{hits}");

    // Chained parent > child.
    let hits = ok(&dir, &["query", "doc.docx", "paragraph > run[bold=true]"]);
    assert!(hits.contains("Bold intro"), "{hits}");
}

#[test]
fn move_and_swap() {
    let dir = temp_dir("moveswap");
    ok(&dir, &["create", "doc.docx"]);
    for text in ["One", "Two", "Three"] {
        ok(&dir, &[
            "add", "doc.docx", "/body", "--type", "paragraph",
            "--prop", &format!("text={text}"),
        ]);
    }
    ok(&dir, &["move", "doc.docx", "/body/p[3]", "--index", "0"]);
    let text = ok(&dir, &["view", "doc.docx", "text"]);
    assert!(text.starts_with("Three"), "{text}");
    ok(&dir, &["swap", "doc.docx", "/body/p[1]", "/body/p[2]"]);
    let text = ok(&dir, &["view", "doc.docx", "text"]);
    assert!(text.starts_with("One\nThree"), "{text}");

    // pptx slide reorder.
    ok(&dir, &["create", "deck.pptx"]);
    for title in ["A", "B", "C"] {
        ok(&dir, &[
            "add", "deck.pptx", "/", "--type", "slide",
            "--prop", &format!("title={title}"),
        ]);
    }
    ok(&dir, &["move", "deck.pptx", "/slide[3]", "--index", "0"]);
    let outline = ok(&dir, &["view", "deck.pptx", "outline"]);
    assert!(outline.contains("Slide 1: C"), "{outline}");
    ok(&dir, &["swap", "deck.pptx", "/slide[2]", "/slide[3]"]);
    let outline = ok(&dir, &["view", "deck.pptx", "outline"]);
    assert!(outline.contains("Slide 2: B"), "{outline}");

    // xlsx sheet reorder.
    ok(&dir, &["create", "wb.xlsx"]);
    ok(&dir, &["add", "wb.xlsx", "/", "--type", "sheet", "--prop", "name=Alpha"]);
    ok(&dir, &["move", "wb.xlsx", "/Alpha", "--index", "0"]);
    let outline = ok(&dir, &["view", "wb.xlsx", "outline"]);
    assert!(outline.contains("Sheet 1: Alpha"), "{outline}");
}

#[test]
fn images_embed() {
    let dir = temp_dir("images");
    // Minimal 4x3 PNG (header only is enough for sniffing, but write a
    // complete file so downstream tools can parse it if they want).
    let mut png = vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];
    png.extend_from_slice(&13u32.to_be_bytes());
    png.extend_from_slice(b"IHDR");
    png.extend_from_slice(&4u32.to_be_bytes());
    png.extend_from_slice(&3u32.to_be_bytes());
    png.extend_from_slice(&[8, 2, 0, 0, 0, 0, 0, 0, 0]);
    std::fs::write(dir.join("img.png"), &png).unwrap();

    ok(&dir, &["create", "doc.docx"]);
    ok(&dir, &["add", "doc.docx", "/body", "--type", "image", "--prop", "src=img.png"]);
    ok(&dir, &["create", "deck.pptx"]);
    ok(&dir, &["add", "deck.pptx", "/", "--type", "slide"]);
    let out = ok(&dir, &[
        "add", "deck.pptx", "/slide[1]", "--type", "image",
        "--prop", "src=img.png", "--prop", "w=2in",
    ]);
    // 2in wide, 4:3 intrinsic → h = 1.5in = 1371600 EMU.
    assert!(out.contains("w=1828800"), "{out}");
    assert!(out.contains("h=1371600"), "{out}");

    // Media parts + relationships present.
    let bytes = std::fs::read(dir.join("deck.pptx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    assert!(zip.by_name("ppt/media/image1.png").is_ok());
    let bytes = std::fs::read(dir.join("doc.docx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    assert!(zip.by_name("word/media/image1.png").is_ok());
}

#[test]
fn formula_refs_follow_row_shifts() {
    let dir = temp_dir("formulas");
    ok(&dir, &["create", "c.xlsx"]);
    ok(&dir, &["set", "c.xlsx", "/Sheet1/A1", "--prop", "value=10"]);
    ok(&dir, &["set", "c.xlsx", "/Sheet1/A2", "--prop", "value=20"]);
    ok(&dir, &["set", "c.xlsx", "/Sheet1/A3", "--prop", "value==SUM(A1:A2)"]);

    ok(&dir, &["add", "c.xlsx", "/Sheet1", "--type", "row", "--index", "2", "--prop", "values=5"]);
    let cell = ok(&dir, &["get", "c.xlsx", "/Sheet1/A4", "--json"]);
    let v: serde_json::Value = serde_json::from_str(&cell).unwrap();
    assert_eq!(v["data"]["results"][0]["attributes"]["formula"], "=SUM(A1:A3)");

    ok(&dir, &["remove", "c.xlsx", "/Sheet1/row[2]"]);
    let cell = ok(&dir, &["get", "c.xlsx", "/Sheet1/A3", "--json"]);
    let v: serde_json::Value = serde_json::from_str(&cell).unwrap();
    assert_eq!(v["data"]["results"][0]["attributes"]["formula"], "=SUM(A1:A2)");
}

#[test]
fn html_views_and_dump_roundtrip() {
    let dir = temp_dir("htmldump");
    ok(&dir, &["create", "deck.pptx"]);
    ok(&dir, &[
        "add", "deck.pptx", "/", "--type", "slide",
        "--prop", "title=Round Trip", "--prop", "background=222222",
    ]);
    let html = ok(&dir, &["view", "deck.pptx", "html"]);
    assert!(html.contains("<!DOCTYPE html>") && html.contains("Round Trip"), "{html}");
    assert!(html.contains("background:#222222"), "{html}");

    ok(&dir, &["view", "deck.pptx", "html", "-o", "deck.html"]);
    assert!(dir.join("deck.html").exists());

    // dump → batch into a fresh file reproduces the outline.
    let ops = ok(&dir, &["dump", "deck.pptx"]);
    std::fs::write(dir.join("ops.json"), &ops).unwrap();
    ok(&dir, &["create", "copy.pptx"]);
    ok(&dir, &["batch", "copy.pptx", "--input", "ops.json"]);
    let outline = ok(&dir, &["view", "copy.pptx", "outline"]);
    assert!(outline.contains("Slide 1: Round Trip"), "{outline}");

    // docx html renders headings and formatting.
    ok(&dir, &["create", "doc.docx"]);
    ok(&dir, &[
        "add", "doc.docx", "/body", "--type", "paragraph",
        "--prop", "text=Head", "--prop", "style=Heading1",
    ]);
    let html = ok(&dir, &["view", "doc.docx", "html"]);
    assert!(html.contains("<h1>Head</h1>"), "{html}");
}

#[test]
fn mcp_server_protocol() {
    use std::io::Write;
    use std::process::Stdio;
    let dir = temp_dir("mcp");
    let mut child = Command::new(bin())
        .current_dir(&dir)
        .arg("mcp")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .unwrap();
    let mut stdin = child.stdin.take().unwrap();
    writeln!(stdin, r#"{{"jsonrpc":"2.0","id":1,"method":"initialize","params":{{}}}}"#).unwrap();
    writeln!(stdin, r#"{{"jsonrpc":"2.0","method":"notifications/initialized"}}"#).unwrap();
    writeln!(stdin, r#"{{"jsonrpc":"2.0","id":2,"method":"tools/list"}}"#).unwrap();
    writeln!(
        stdin,
        r#"{{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{{"name":"officecli","arguments":{{"command":"create t.docx"}}}}}}"#
    )
    .unwrap();
    writeln!(
        stdin,
        r#"{{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{{"name":"officecli","arguments":{{"command":"get t.docx /body/p[9]"}}}}}}"#
    )
    .unwrap();
    drop(stdin);
    let out = child.wait_with_output().unwrap();
    let lines: Vec<serde_json::Value> = String::from_utf8_lossy(&out.stdout)
        .lines()
        .map(|l| serde_json::from_str(l).unwrap())
        .collect();
    assert_eq!(lines.len(), 4, "one response per request (not notification)");
    assert_eq!(lines[0]["result"]["serverInfo"]["name"], "officecli");
    assert_eq!(lines[1]["result"]["tools"][0]["name"], "officecli");
    assert_eq!(lines[2]["result"]["isError"], false);
    assert_eq!(lines[3]["result"]["isError"], true);
    assert!(dir.join("t.docx").exists());
}

#[test]
fn xlsx_columns_and_pivot() {
    let dir = temp_dir("cols");
    ok(&dir, &["create", "d.xlsx"]);
    ok(&dir, &[
        "batch", "d.xlsx", "--commands",
        r#"[
          {"command":"set","path":"/Sheet1/A1","props":{"value":"Region"}},
          {"command":"set","path":"/Sheet1/B1","props":{"value":"Sales"}},
          {"command":"set","path":"/Sheet1/A2","props":{"value":"East"}},
          {"command":"set","path":"/Sheet1/B2","props":{"value":"100"}},
          {"command":"set","path":"/Sheet1/A3","props":{"value":"West"}},
          {"command":"set","path":"/Sheet1/B3","props":{"value":"250"}},
          {"command":"set","path":"/Sheet1/A4","props":{"value":"East"}},
          {"command":"set","path":"/Sheet1/B4","props":{"value":"50"}},
          {"command":"set","path":"/Sheet1/C2","props":{"value":"=SUM(B2:B4)"}}
        ]"#,
    ]);
    // Insert a column before B: refs shift B→C.
    ok(&dir, &["add", "d.xlsx", "/Sheet1", "--type", "column", "--prop", "at=B", "--prop", "r1=Manager"]);
    let json = ok(&dir, &["get", "d.xlsx", "/Sheet1/D2", "--json"]);
    assert!(json.contains("=SUM(C2:C4)"), "{json}");
    // Remove it again: refs shift back.
    ok(&dir, &["remove", "d.xlsx", "/Sheet1/B"]);
    let json = ok(&dir, &["get", "d.xlsx", "/Sheet1/C2", "--json"]);
    assert!(json.contains("=SUM(B2:B4)"), "{json}");
    // Computed pivot.
    ok(&dir, &[
        "add", "d.xlsx", "/", "--type", "pivot",
        "--prop", "source=A1:B4", "--prop", "rows=Region",
        "--prop", "values=Sales", "--prop", "agg=sum",
    ]);
    let text = ok(&dir, &["view", "d.xlsx", "text"]);
    assert!(text.contains("Pivot"), "{text}");
    assert!(text.contains("East\t150"), "{text}");
    assert!(text.contains("Grand Total\t400"), "{text}");
}

#[test]
fn charts_in_both_formats() {
    let dir = temp_dir("charts");
    ok(&dir, &["create", "d.xlsx"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/A1", "--prop", "value=Q"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/B1", "--prop", "value=Rev"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/A2", "--prop", "value=Q1"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/B2", "--prop", "value=10"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/A3", "--prop", "value=Q2"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/B3", "--prop", "value=20"]);
    let out = ok(&dir, &[
        "add", "d.xlsx", "/Sheet1", "--type", "chart",
        "--prop", "kind=column", "--prop", "data=A1:B3", "--prop", "title=Revenue",
    ]);
    assert!(out.contains("chart"), "{out}");
    let bytes = std::fs::read(dir.join("d.xlsx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    assert!(zip.by_name("xl/charts/chart1.xml").is_ok());
    assert!(zip.by_name("xl/drawings/drawing1.xml").is_ok());

    ok(&dir, &["create", "p.pptx"]);
    ok(&dir, &["add", "p.pptx", "/", "--type", "slide", "--prop", "title=T"]);
    ok(&dir, &[
        "add", "p.pptx", "/slide[1]", "--type", "chart",
        "--prop", "kind=pie", "--prop", "categories=A,B", "--prop", "values=60,40",
    ]);
    let bytes = std::fs::read(dir.join("p.pptx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    assert!(zip.by_name("ppt/charts/chart1.xml").is_ok());
    let outline = ok(&dir, &["view", "p.pptx", "outline"]);
    assert!(outline.contains("graphicFrame"), "{outline}");
}

#[test]
fn docx_annotations() {
    let dir = temp_dir("annot");
    ok(&dir, &["create", "r.docx"]);
    ok(&dir, &["add", "r.docx", "/body", "--type", "toc"]);
    ok(&dir, &["add", "r.docx", "/body", "--type", "paragraph", "--prop", "text=Body text."]);
    ok(&dir, &[
        "add", "r.docx", "/body/p[2]", "--type", "comment",
        "--prop", "text=Please verify", "--prop", "author=Reviewer",
    ]);
    ok(&dir, &["add", "r.docx", "/body/p[2]", "--type", "footnote", "--prop", "text=A source."]);
    ok(&dir, &["add", "r.docx", "/body/p[2]", "--type", "field", "--prop", "kind=page"]);
    let comments = ok(&dir, &["view", "r.docx", "comments"]);
    assert!(comments.contains("Please verify") && comments.contains("Reviewer"), "{comments}");
    let bytes = std::fs::read(dir.join("r.docx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    for part in ["word/comments.xml", "word/footnotes.xml", "word/settings.xml"] {
        assert!(zip.by_name(part).is_ok(), "missing {part}");
    }
    ok(&dir, &["remove", "r.docx", "/comment[1]"]);
    let comments = ok(&dir, &["view", "r.docx", "comments"]);
    assert!(!comments.contains("Please verify"), "{comments}");
}

#[test]
fn pptx_transitions_and_animations() {
    let dir = temp_dir("anim");
    ok(&dir, &["create", "p.pptx"]);
    ok(&dir, &["add", "p.pptx", "/", "--type", "slide", "--prop", "title=T"]);
    ok(&dir, &[
        "set", "p.pptx", "/slide[1]",
        "--prop", "transition=push", "--prop", "direction=left",
        "--prop", "speed=fast", "--prop", "advance=3s",
    ]);
    ok(&dir, &["set", "p.pptx", "/slide[1]/shape[1]", "--prop", "animation=fade"]);
    let bytes = std::fs::read(dir.join("p.pptx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut zip.by_name("ppt/slides/slide1.xml").unwrap(), &mut xml).unwrap();
    assert!(xml.contains("<p:push dir=\"l\"/>"), "{xml}");
    assert!(xml.contains("advTm=\"3000\""), "{xml}");
    assert!(xml.contains("<p:timing>") && xml.contains("filter=\"fade\""), "{xml}");
    // Transition survives a dump→replay cycle.
    let dump = ok(&dir, &["dump", "p.pptx"]);
    assert!(dump.contains("\"transition\": \"push\""), "{dump}");
}

#[test]
fn screenshots_render_all_formats() {
    let dir = temp_dir("shots");
    ok(&dir, &["create", "d.docx"]);
    ok(&dir, &["add", "d.docx", "/body", "--type", "paragraph", "--prop", "text=Hello", "--prop", "style=Heading1"]);
    ok(&dir, &["view", "d.docx", "screenshot", "-o", "d.png"]);
    ok(&dir, &["create", "s.xlsx"]);
    ok(&dir, &["set", "s.xlsx", "/Sheet1/A1", "--prop", "value=42"]);
    ok(&dir, &["view", "s.xlsx", "screenshot", "-o", "s.png"]);
    ok(&dir, &["create", "p.pptx"]);
    ok(&dir, &["add", "p.pptx", "/", "--type", "slide", "--prop", "title=T", "--prop", "background=1A1A2E"]);
    ok(&dir, &["view", "p.pptx", "screenshot", "-o", "p.png"]);
    for f in ["d.png", "s.png", "p.png"] {
        let bytes = std::fs::read(dir.join(f)).unwrap();
        assert!(bytes.starts_with(&[0x89, b'P', b'N', b'G']), "{f} is not a PNG");
        assert!(bytes.len() > 500, "{f} suspiciously small");
    }
    // Screenshot without -o is an error.
    fails(&dir, &["view", "p.pptx", "screenshot"]);
}

#[test]
fn resident_mode_loop() {
    use std::io::Write;
    use std::process::Stdio;
    let dir = temp_dir("resident");
    let mut child = Command::new(bin())
        .current_dir(&dir)
        .arg("resident")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .unwrap();
    let mut stdin = child.stdin.take().unwrap();
    writeln!(stdin, "create t.xlsx").unwrap();
    writeln!(stdin, "set t.xlsx /Sheet1/A1 --prop value=7").unwrap();
    writeln!(stdin, "get t.xlsx /Sheet1/A1").unwrap();
    writeln!(stdin, "explode t.xlsx").unwrap();
    writeln!(stdin, "exit").unwrap();
    drop(stdin);
    let out = child.wait_with_output().unwrap();
    assert!(out.status.success());
    let lines: Vec<serde_json::Value> = String::from_utf8_lossy(&out.stdout)
        .lines()
        .map(|l| serde_json::from_str(l).unwrap())
        .collect();
    assert_eq!(lines.len(), 4, "one JSON line per command");
    assert_eq!(lines[0]["ok"], true);
    assert_eq!(lines[2]["data"]["results"][0]["text"], "7");
    assert_eq!(lines[3]["ok"], false);
}

#[test]
fn image_srcdata_dump_roundtrip() {
    let dir = temp_dir("srcdata");
    // 1x1 red PNG.
    let png_b64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==";
    ok(&dir, &["create", "a.docx"]);
    ok(&dir, &["add", "a.docx", "/body", "--type", "image", "--prop", &format!("srcdata={png_b64}")]);
    let dump = ok(&dir, &["dump", "a.docx"]);
    assert!(dump.contains("srcdata"), "{dump}");
    std::fs::write(dir.join("ops.json"), &dump).unwrap();
    ok(&dir, &["create", "b.docx"]);
    ok(&dir, &["batch", "b.docx", "--input", "ops.json"]);
    let bytes = std::fs::read(dir.join("b.docx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    assert!(zip.by_name("word/media/image1.png").is_ok());
}

#[test]
fn dates_formats_and_csv() {
    let dir = temp_dir("datescsv");
    ok(&dir, &["create", "d.xlsx"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/A1", "--prop", "value=2026-07-10"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/A2", "--prop", "value=0.185", "--prop", "format=0.00%"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/A3", "--prop", "value=2026-07-10", "--prop", "type=string"]);
    let json = ok(&dir, &["get", "d.xlsx", "/Sheet1/A1", "--json"]);
    let v: serde_json::Value = serde_json::from_str(&json).unwrap();
    assert_eq!(v["data"]["results"][0]["text"], "2026-07-10");
    assert_eq!(v["data"]["results"][0]["attributes"]["type"], "date");
    let json = ok(&dir, &["get", "d.xlsx", "/Sheet1/A3", "--json"]);
    assert!(json.contains("\"type\": \"string\""), "{json}");
    // The stored value is a serial number, not a string.
    let bytes = std::fs::read(dir.join("d.xlsx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut sheet = String::new();
    std::io::Read::read_to_string(&mut zip.by_name("xl/worksheets/sheet1.xml").unwrap(), &mut sheet).unwrap();
    assert!(sheet.contains("<v>46213</v>"), "{sheet}");

    // CSV round trip.
    std::fs::write(dir.join("in.csv"), "Name,When\n\"A, B\",2026-01-02\n").unwrap();
    ok(&dir, &["create", "c.xlsx"]);
    ok(&dir, &["add", "c.xlsx", "/Sheet1", "--type", "csv", "--prop", "src=in.csv"]);
    let csv = ok(&dir, &["export", "c.xlsx"]);
    assert!(csv.contains("\"A, B\",2026-01-02"), "{csv}");
}

#[test]
fn hyperlinks_everywhere() {
    let dir = temp_dir("links");
    ok(&dir, &["create", "l.docx"]);
    ok(&dir, &["add", "l.docx", "/body", "--type", "paragraph", "--prop", "text=See "]);
    ok(&dir, &[
        "add", "l.docx", "/body/p[1]", "--type", "hyperlink",
        "--prop", "url=https://example.com", "--prop", "text=docs",
    ]);
    let json = ok(&dir, &["get", "l.docx", "/body/p[1]", "--depth", "1", "--json"]);
    assert!(json.contains("https://example.com"), "{json}");

    ok(&dir, &["create", "l.xlsx"]);
    ok(&dir, &["set", "l.xlsx", "/Sheet1/A1", "--prop", "value=Home", "--prop", "url=https://example.org"]);
    let bytes = std::fs::read(dir.join("l.xlsx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut sheet = String::new();
    std::io::Read::read_to_string(&mut zip.by_name("xl/worksheets/sheet1.xml").unwrap(), &mut sheet).unwrap();
    assert!(sheet.contains("<hyperlinks>"), "{sheet}");

    ok(&dir, &["create", "l.pptx"]);
    ok(&dir, &["add", "l.pptx", "/", "--type", "slide", "--prop", "title=T"]);
    ok(&dir, &[
        "add", "l.pptx", "/slide[1]", "--type", "shape",
        "--prop", "text=Visit", "--prop", "url=https://example.net", "--prop", "y=3in",
    ]);
    let bytes = std::fs::read(dir.join("l.pptx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut rels = String::new();
    std::io::Read::read_to_string(&mut zip.by_name("ppt/slides/_rels/slide1.xml.rels").unwrap(), &mut rels).unwrap();
    assert!(rels.contains("https://example.net") && rels.contains("External"), "{rels}");
}

#[test]
fn lists_in_docx_and_pptx() {
    let dir = temp_dir("lists");
    ok(&dir, &["create", "r.docx"]);
    ok(&dir, &[
        "add", "r.docx", "/body", "--type", "list",
        "--prop", r"items=One\nTwo\n\tNested",
    ]);
    ok(&dir, &[
        "add", "r.docx", "/body", "--type", "list", "--prop", "kind=number",
        "--prop", r"items=First\nSecond",
    ]);
    let text = ok(&dir, &["view", "r.docx", "text"]);
    assert!(text.contains("• One"), "{text}");
    assert!(text.contains("  • Nested"), "{text}");
    assert!(text.contains("2. Second"), "{text}");
    let json = ok(&dir, &["get", "r.docx", "/body/p[3]", "--json"]);
    assert!(json.contains("\"list\": \"bullet\"") && json.contains("\"level\": \"1\""), "{json}");
    // list=none strips the numbering.
    ok(&dir, &["set", "r.docx", "/body/p[1]", "--prop", "list=none"]);
    let text = ok(&dir, &["view", "r.docx", "text"]);
    assert!(text.contains("\nOne") || text.starts_with("One"), "{text}");

    ok(&dir, &["create", "d.pptx"]);
    ok(&dir, &["add", "d.pptx", "/", "--type", "slide", "--prop", "title=T"]);
    ok(&dir, &[
        "add", "d.pptx", "/slide[1]", "--type", "shape", "--prop", "list=bullet",
        "--prop", r"text=Alpha\n\tBeta", "--prop", "y=2in",
    ]);
    let bytes = std::fs::read(dir.join("d.pptx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut slide = String::new();
    std::io::Read::read_to_string(&mut zip.by_name("ppt/slides/slide1.xml").unwrap(), &mut slide).unwrap();
    assert!(slide.contains("buChar") && slide.contains("lvl=\"1\""), "{slide}");
}

#[test]
fn speaker_notes_lifecycle() {
    let dir = temp_dir("notes");
    ok(&dir, &["create", "n.pptx"]);
    ok(&dir, &["add", "n.pptx", "/", "--type", "slide", "--prop", "title=A", "--prop", "notes=First take"]);
    ok(&dir, &["set", "n.pptx", "/slide[1]", "--prop", "notes=Revised"]);
    let notes = ok(&dir, &["view", "n.pptx", "notes"]);
    assert!(notes.contains("Revised") && !notes.contains("First take"), "{notes}");
    // Dump replays notes.
    let dump = ok(&dir, &["dump", "n.pptx"]);
    assert!(dump.contains("\"notes\": \"Revised\""), "{dump}");
    let bytes = std::fs::read(dir.join("n.pptx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    assert!(zip.by_name("ppt/notesSlides/notesSlide1.xml").is_ok());
    assert!(zip.by_name("ppt/notesMasters/notesMaster1.xml").is_ok());
}

#[test]
fn copy_duplicates_elements() {
    let dir = temp_dir("copy");
    ok(&dir, &["create", "c.docx"]);
    ok(&dir, &["add", "c.docx", "/body", "--type", "paragraph", "--prop", "text=Row"]);
    ok(&dir, &["copy", "c.docx", "/body/p[1]"]);
    let text = ok(&dir, &["view", "c.docx", "text"]);
    assert_eq!(text.matches("Row").count(), 2, "{text}");

    ok(&dir, &["create", "c.pptx"]);
    ok(&dir, &["add", "c.pptx", "/", "--type", "slide", "--prop", "title=S", "--prop", "notes=talk"]);
    ok(&dir, &["copy", "c.pptx", "/slide[1]"]);
    let outline = ok(&dir, &["view", "c.pptx", "outline"]);
    assert_eq!(outline.matches("Slide ").count(), 2, "{outline}");
    let notes = ok(&dir, &["view", "c.pptx", "notes"]);
    assert_eq!(notes.matches("talk").count(), 2, "{notes}");

    ok(&dir, &["create", "c.xlsx"]);
    ok(&dir, &["set", "c.xlsx", "/Sheet1/A1", "--prop", "value=x", "--prop", "url=https://example.com"]);
    ok(&dir, &["copy", "c.xlsx", "/Sheet1"]);
    let outline = ok(&dir, &["view", "c.xlsx", "outline"]);
    assert!(outline.contains("Sheet1 (2)"), "{outline}");
    // The copy carries the sheet rels, so its hyperlink r:id resolves.
    let bytes = std::fs::read(dir.join("c.xlsx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    assert!(zip.by_name("xl/worksheets/_rels/sheet2.xml.rels").is_ok());
}

/// Rewrite one zip entry in place (simulating Word/PowerPoint-authored XML).
fn rewrite_zip_entry(path: &Path, entry: &str, from: &str, to: &str) {
    let bytes = std::fs::read(path).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut items: Vec<(String, Vec<u8>)> = Vec::new();
    for i in 0..zip.len() {
        let mut f = zip.by_index(i).unwrap();
        let mut buf = Vec::new();
        std::io::Read::read_to_end(&mut f, &mut buf).unwrap();
        items.push((f.name().to_string(), buf));
    }
    let item = items.iter_mut().find(|(n, _)| n == entry).unwrap();
    let text = String::from_utf8(item.1.clone()).unwrap();
    assert!(text.contains(from), "pattern not found in {entry}");
    item.1 = text.replace(from, to).into_bytes();
    let mut cursor = std::io::Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut cursor);
        let options = zip::write::SimpleFileOptions::default();
        for (name, data) in &items {
            writer.start_file(name, options).unwrap();
            std::io::Write::write_all(&mut writer, data).unwrap();
        }
        writer.finish().unwrap();
    }
    std::fs::write(path, cursor.into_inner()).unwrap();
}

#[test]
fn theme_colors_resolve() {
    let dir = temp_dir("theme");
    ok(&dir, &["create", "t.pptx"]);
    ok(&dir, &["add", "t.pptx", "/", "--type", "slide", "--prop", "title=T"]);
    ok(&dir, &["add", "t.pptx", "/slide[1]", "--type", "shape", "--prop", "text=Box", "--prop", "fill=FF0000", "--prop", "y=3in"]);
    rewrite_zip_entry(
        &dir.join("t.pptx"),
        "ppt/slides/slide1.xml",
        r#"<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>"#,
        r#"<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>"#,
    );
    let dump = ok(&dir, &["dump", "t.pptx"]);
    // accent1 in the default theme is 4472C4.
    assert!(dump.contains("\"fill\": \"4472C4\""), "{dump}");
}

#[test]
fn sdt_content_controls_are_transparent() {
    let dir = temp_dir("sdt");
    ok(&dir, &["create", "s.docx"]);
    ok(&dir, &["add", "s.docx", "/body", "--type", "paragraph", "--prop", "text=Before"]);
    ok(&dir, &["add", "s.docx", "/body", "--type", "paragraph", "--prop", "text=Wrapped"]);
    ok(&dir, &["add", "s.docx", "/body", "--type", "paragraph", "--prop", "text=After"]);
    let target = r#"<w:p><w:r><w:t xml:space="preserve">Wrapped</w:t></w:r></w:p>"#;
    rewrite_zip_entry(
        &dir.join("s.docx"),
        "word/document.xml",
        target,
        &format!("<w:sdt><w:sdtPr><w:id w:val=\"1\"/></w:sdtPr><w:sdtContent>{target}</w:sdtContent></w:sdt>"),
    );
    // Addressing, editing, and views all look through the control.
    let json = ok(&dir, &["get", "s.docx", "/body/p[2]", "--json"]);
    assert!(json.contains("Wrapped"), "{json}");
    ok(&dir, &["set", "s.docx", "/body/p[2]", "--prop", "text=Edited"]);
    let text = ok(&dir, &["view", "s.docx", "text"]);
    assert_eq!(text.trim(), "Before\nEdited\nAfter", "{text}");
    // The wrapper survives the edit.
    let bytes = std::fs::read(dir.join("s.docx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut doc = String::new();
    std::io::Read::read_to_string(&mut zip.by_name("word/document.xml").unwrap(), &mut doc).unwrap();
    assert!(doc.contains("<w:sdt>") && doc.contains("Edited"), "{doc}");
}

#[test]
fn formula_evaluation_and_calc() {
    let dir = temp_dir("formula");
    ok(&dir, &["create", "d.xlsx"]);
    for (path, val) in [
        ("A1", "10"), ("A2", "20"), ("A3", "30"),
        ("B1", "=SUM(A1:A3)"), ("B2", "=AVERAGE(A1:A3)"),
        ("B3", "=IF(B1>50,\"big\",\"small\")"),
        ("B4", "=VLOOKUP(20,A1:A3,1,FALSE)"),
    ] {
        ok(&dir, &["set", "d.xlsx", &format!("/Sheet1/{path}"), "--prop", &format!("value={val}")]);
    }
    let json = ok(&dir, &["get", "d.xlsx", "/Sheet1/B1", "--computed", "--json"]);
    let v: serde_json::Value = serde_json::from_str(&json).unwrap();
    assert_eq!(v["data"]["results"][0]["text"], "60");
    assert_eq!(v["data"]["results"][0]["attributes"]["computed"], "true");
    let out = ok(&dir, &["get", "d.xlsx", "/Sheet1/B3", "--computed"]);
    assert!(out.contains("\"small\"") || out.contains("small"), "{out}");
    // Ad-hoc calc against the workbook.
    let out = ok(&dir, &["calc", "d.xlsx", "=SUM(A1:A3)*2"]);
    assert_eq!(out.trim(), "120");
    let out = ok(&dir, &["calc", "d.xlsx", "=1/0"]);
    assert_eq!(out.trim(), "#DIV/0!");
}

#[test]
fn sort_merge_width_height() {
    let dir = temp_dir("layout");
    ok(&dir, &["create", "d.xlsx"]);
    for (path, val) in [
        ("A1", "Name"), ("B1", "Score"),
        ("A2", "C"), ("B2", "30"),
        ("A3", "A"), ("B3", "10"),
        ("A4", "B"), ("B4", "20"),
    ] {
        ok(&dir, &["set", "d.xlsx", &format!("/Sheet1/{path}"), "--prop", &format!("value={val}")]);
    }
    ok(&dir, &["sort", "d.xlsx", "Sheet1!A2:B4", "--by", "B", "--desc"]);
    let json = ok(&dir, &["get", "d.xlsx", "/Sheet1/A2", "--json"]);
    assert!(json.contains("\"text\": \"C\""), "{json}"); // C (30) first
    ok(&dir, &["set", "d.xlsx", "/Sheet1/A1:B1", "--prop", "merge=true"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/A", "--prop", "width=20"]);
    ok(&dir, &["set", "d.xlsx", "/Sheet1/row[1]", "--prop", "height=30"]);
    let bytes = std::fs::read(dir.join("d.xlsx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut sheet = String::new();
    std::io::Read::read_to_string(&mut zip.by_name("xl/worksheets/sheet1.xml").unwrap(), &mut sheet).unwrap();
    assert!(sheet.contains("<mergeCells"), "{sheet}");
    assert!(sheet.contains("customWidth"), "{sheet}");
    assert!(sheet.contains("customHeight"), "{sheet}");
}

#[test]
fn pptx_tables_roundtrip() {
    let dir = temp_dir("ptables");
    ok(&dir, &["create", "t.pptx"]);
    ok(&dir, &["add", "t.pptx", "/", "--type", "slide", "--prop", "title=T"]);
    ok(&dir, &[
        "add", "t.pptx", "/slide[1]", "--type", "table",
        "--prop", r"data=Region,Q1\nEast,100\nWest,250",
    ]);
    let text = ok(&dir, &["view", "t.pptx", "text"]);
    assert!(text.contains("Region | Q1"), "{text}");
    assert!(text.contains("East | 100"), "{text}");
    ok(&dir, &["set", "t.pptx", "/slide[1]/shape[2]", "--prop", r"data=Region,Q1\nEast,999\nWest,250"]);
    let text = ok(&dir, &["view", "t.pptx", "text"]);
    assert!(text.contains("East | 999"), "{text}");
    let bytes = std::fs::read(dir.join("t.pptx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut slide = String::new();
    std::io::Read::read_to_string(&mut zip.by_name("ppt/slides/slide1.xml").unwrap(), &mut slide).unwrap();
    assert!(slide.contains("<a:tbl>"), "{slide}");
}

#[test]
fn docx_headers_footers_page_setup() {
    let dir = temp_dir("hf");
    ok(&dir, &["create", "r.docx"]);
    ok(&dir, &["add", "r.docx", "/body", "--type", "paragraph", "--prop", "text=Body"]);
    ok(&dir, &["add", "r.docx", "/", "--type", "header", "--prop", "text=CONFIDENTIAL"]);
    ok(&dir, &["add", "r.docx", "/", "--type", "footer", "--prop", "page-numbers=true"]);
    ok(&dir, &["set", "r.docx", "/", "--prop", "orientation=landscape", "--prop", "margins=0.75in"]);
    let bytes = std::fs::read(dir.join("r.docx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    assert!(zip.by_name("word/header1.xml").is_ok());
    assert!(zip.by_name("word/footer1.xml").is_ok());
    let mut doc = String::new();
    std::io::Read::read_to_string(&mut zip.by_name("word/document.xml").unwrap(), &mut doc).unwrap();
    assert!(doc.contains("w:orient=\"landscape\""), "{doc}");
    assert!(doc.contains("<w:headerReference"), "{doc}");
    // Removal drops the reference and part.
    ok(&dir, &["remove", "r.docx", "/header"]);
    let bytes = std::fs::read(dir.join("r.docx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    assert!(zip.by_name("word/header1.xml").is_err());
}

#[test]
fn diff_reports_changes() {
    let dir = temp_dir("diff");
    ok(&dir, &["create", "a.docx"]);
    ok(&dir, &["add", "a.docx", "/body", "--type", "paragraph", "--prop", "text=Original"]);
    std::fs::copy(dir.join("a.docx"), dir.join("b.docx")).unwrap();
    ok(&dir, &["set", "b.docx", "/body/p[1]", "--prop", "text=Edited"]);
    ok(&dir, &["add", "b.docx", "/body", "--type", "paragraph", "--prop", "text=New"]);
    let out = ok(&dir, &["diff", "a.docx", "b.docx"]);
    assert!(out.contains("(changed)") && out.contains("from=Original"), "{out}");
    assert!(out.contains("(added)"), "{out}");
    let same = ok(&dir, &["diff", "a.docx", "a.docx"]);
    assert!(same.contains("identical"), "{same}");
}

#[test]
fn safety_rails() {
    let dir = temp_dir("safety");
    ok(&dir, &["create", "d.docx"]);
    ok(&dir, &["add", "d.docx", "/body", "--type", "paragraph", "--prop", "text=Keep me"]);
    // --dry-run does not persist.
    ok(&dir, &["set", "d.docx", "/body/p[1]", "--prop", "text=Changed", "--dry-run"]);
    let text = ok(&dir, &["view", "d.docx", "text"]);
    assert!(text.contains("Keep me") && !text.contains("Changed"), "{text}");
    // --backup writes a .bak of the pre-change file.
    ok(&dir, &["set", "d.docx", "/body/p[1]", "--prop", "text=Changed", "--backup"]);
    assert!(dir.join("d.docx.bak").exists());
    // --atomic batch: any failure discards everything.
    ok(&dir, &["create", "e.xlsx"]);
    let out = run(&dir, &[
        "batch", "e.xlsx", "--atomic", "--commands",
        r#"[{"command":"set","path":"/Sheet1/A1","props":{"value":"x"}},{"command":"set","path":"/Nope/A1","props":{"value":"y"}}]"#,
    ]);
    assert!(!out.status.success());
    let a1 = ok(&dir, &["get", "e.xlsx", "/Sheet1/A1", "--json"]);
    assert!(a1.contains("\"type\": \"empty\""), "atomic batch should have saved nothing: {a1}");
}

/// Read one zip entry of an OOXML package as text.
fn zip_text(path: &Path, entry: &str) -> String {
    use std::io::Read;
    let bytes = std::fs::read(path).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut s = String::new();
    zip.by_name(entry)
        .unwrap_or_else(|_| panic!("missing zip entry {entry}"))
        .read_to_string(&mut s)
        .unwrap();
    s
}

#[test]
fn scatter_charts_and_docx_charts() {
    let dir = temp_dir("v05charts");
    // xlsx scatter: first column = x values.
    ok(&dir, &["create", "sc.xlsx"]);
    for (i, (x, y)) in [(1, 1), (2, 4), (3, 9)].iter().enumerate() {
        ok(&dir, &["set", "sc.xlsx", &format!("/Sheet1/A{}", i + 1), "--prop", &format!("value={x}")]);
        ok(&dir, &["set", "sc.xlsx", &format!("/Sheet1/B{}", i + 1), "--prop", &format!("value={y}")]);
    }
    ok(&dir, &[
        "add", "sc.xlsx", "/Sheet1", "--type", "chart",
        "--prop", "kind=scatter", "--prop", "data=A1:B3",
    ]);
    let chart = zip_text(&dir.join("sc.xlsx"), "xl/charts/chart1.xml");
    assert!(chart.contains("<c:scatterChart>"), "{chart}");
    assert!(chart.contains("<c:xVal>") && chart.contains("<c:yVal>"), "{chart}");
    assert_eq!(chart.matches("<c:valAx>").count(), 2, "scatter needs two value axes");

    // pptx scatter with explicit x values + embedded workbook (Edit Data).
    ok(&dir, &["create", "p.pptx"]);
    ok(&dir, &["add", "p.pptx", "/", "--type", "slide"]);
    ok(&dir, &[
        "add", "p.pptx", "/slide[1]", "--type", "chart",
        "--prop", "kind=scatter", "--prop", "xvalues=1,2,4",
        "--prop", "values=3,5,4", "--prop", "series=Trials",
    ]);
    let chart = zip_text(&dir.join("p.pptx"), "ppt/charts/chart1.xml");
    assert!(chart.contains("<c:scatterChart>"), "{chart}");
    assert!(chart.contains("c:externalData"), "chart should link its embedded workbook");
    assert!(chart.contains("Sheet1!$B$2:$B$4"), "series should reference the workbook: {chart}");
    let rels = zip_text(&dir.join("p.pptx"), "ppt/charts/_rels/chart1.xml.rels");
    assert!(rels.contains("embeddings/Microsoft_Excel_Worksheet1.xlsx"), "{rels}");
    let bytes = std::fs::read(dir.join("p.pptx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    assert!(zip.by_name("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx").is_ok());

    // docx inline chart.
    ok(&dir, &["create", "c.docx"]);
    ok(&dir, &[
        "add", "c.docx", "/body", "--type", "chart",
        "--prop", "kind=column", "--prop", "categories=Q1,Q2",
        "--prop", "values=10,20", "--prop", "title=Revenue",
    ]);
    let doc = zip_text(&dir.join("c.docx"), "word/document.xml");
    assert!(doc.contains("drawingml/2006/chart"), "{doc}");
    let chart = zip_text(&dir.join("c.docx"), "word/charts/chart1.xml");
    assert!(chart.contains("<c:barChart>") && chart.contains("Revenue"), "{chart}");
    let bytes = std::fs::read(dir.join("c.docx")).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    assert!(zip.by_name("word/embeddings/Microsoft_Excel_Worksheet1.xlsx").is_ok());
}

#[test]
fn xlsx_cell_comments() {
    let dir = temp_dir("v05comments");
    ok(&dir, &["create", "c.xlsx"]);
    ok(&dir, &["set", "c.xlsx", "/Sheet1/A1", "--prop", "value=Total"]);
    let out = ok(&dir, &[
        "set", "c.xlsx", "/Sheet1/B2",
        "--prop", "comment=Check this figure", "--prop", "author=Reviewer",
    ]);
    assert!(out.contains("comment=Check this figure"), "{out}");
    ok(&dir, &["set", "c.xlsx", "/Sheet1/A1", "--prop", "comment=Header note"]);

    let view = ok(&dir, &["view", "c.xlsx", "comments"]);
    assert!(view.contains("Sheet1!B2 [Reviewer] Check this figure"), "{view}");
    assert!(view.contains("Sheet1!A1 [officecli] Header note"), "{view}");

    // The comments part, the VML note shapes, and the sheet hook all exist.
    let comments = zip_text(&dir.join("c.xlsx"), "xl/comments1.xml");
    assert!(comments.contains("Check this figure") && comments.contains("Reviewer"), "{comments}");
    let vml = zip_text(&dir.join("c.xlsx"), "xl/drawings/vmlDrawing1.vml");
    assert_eq!(vml.matches("ObjectType=\"Note\"").count(), 2, "{vml}");
    let sheet = zip_text(&dir.join("c.xlsx"), "xl/worksheets/sheet1.xml");
    assert!(sheet.contains("<legacyDrawing"), "{sheet}");

    // dump replays comments.
    let dump = ok(&dir, &["dump", "c.xlsx"]);
    assert!(dump.contains("Header note"), "{dump}");

    // Empty comment removes.
    ok(&dir, &["set", "c.xlsx", "/Sheet1/B2", "--prop", "comment="]);
    let view = ok(&dir, &["view", "c.xlsx", "comments"]);
    assert!(!view.contains("Check this figure"), "{view}");
    assert!(view.contains("Header note"), "{view}");
}

#[test]
fn screenshots_draw_charts_and_decode_jpeg_gif() {
    let dir = temp_dir("v05shots");
    // A 4x4 red JPEG and GIF (generated fixtures, embedded as base64).
    const JPEG_B64: &str = "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCAAEAAQDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDj6KKK+BP3M//Z";
    const GIF_B64: &str = "R0lGODdhBAAEAIEAAMg8HgAAAAAAAAAAACwAAAAABAAEAAAICQABCBxIsCCAgAA7";

    // pptx chart screenshot: chart data renders instead of a placeholder.
    ok(&dir, &["create", "p.pptx"]);
    ok(&dir, &["add", "p.pptx", "/", "--type", "slide", "--prop", "title=T"]);
    ok(&dir, &[
        "add", "p.pptx", "/slide[1]", "--type", "chart",
        "--prop", "kind=pie", "--prop", "categories=A,B", "--prop", "values=60,40",
    ]);
    ok(&dir, &["view", "p.pptx", "screenshot", "-o", "p.png"]);
    let png = std::fs::read(dir.join("p.png"))
        .or_else(|_| std::fs::read(dir.join("p-1.png")))
        .unwrap();
    assert!(png.starts_with(b"\x89PNG"));

    // xlsx chart anchored on the grid.
    ok(&dir, &["create", "g.xlsx"]);
    ok(&dir, &["set", "g.xlsx", "/Sheet1/A1", "--prop", "value=5"]);
    ok(&dir, &["set", "g.xlsx", "/Sheet1/A2", "--prop", "value=9"]);
    ok(&dir, &[
        "add", "g.xlsx", "/Sheet1", "--type", "chart",
        "--prop", "kind=column", "--prop", "data=A1:A2",
    ]);
    ok(&dir, &["view", "g.xlsx", "screenshot", "-o", "g.png"]);
    let png = std::fs::read(dir.join("g.png"))
        .or_else(|_| std::fs::read(dir.join("g-1.png")))
        .unwrap();
    assert!(png.len() > 4000, "chart area should add pixels");

    // docx: JPEG + GIF images and an italic run all render.
    use base64::Engine;
    let jpg = base64::engine::general_purpose::STANDARD.decode(JPEG_B64).unwrap();
    let gif = base64::engine::general_purpose::STANDARD.decode(GIF_B64).unwrap();
    std::fs::write(dir.join("pic.jpg"), &jpg).unwrap();
    std::fs::write(dir.join("pic.gif"), &gif).unwrap();
    ok(&dir, &["create", "m.docx"]);
    ok(&dir, &["add", "m.docx", "/body", "--type", "image", "--prop", "src=pic.jpg"]);
    ok(&dir, &["add", "m.docx", "/body", "--type", "image", "--prop", "src=pic.gif"]);
    ok(&dir, &[
        "add", "m.docx", "/body", "--type", "paragraph",
        "--prop", "text=slanted", "--prop", "italic=true",
    ]);
    ok(&dir, &["view", "m.docx", "screenshot", "-o", "m.png"]);
    let png = std::fs::read(dir.join("m.png"))
        .or_else(|_| std::fs::read(dir.join("m-1.png")))
        .unwrap();
    assert!(png.starts_with(b"\x89PNG"));
}

#[test]
fn fly_in_motion_animation() {
    let dir = temp_dir("v05flyin");
    ok(&dir, &["create", "a.pptx"]);
    ok(&dir, &["add", "a.pptx", "/", "--type", "slide", "--prop", "title=T"]);
    ok(&dir, &[
        "add", "a.pptx", "/slide[1]", "--type", "shape",
        "--prop", "text=Mover", "--prop", "x=1in", "--prop", "y=3in",
    ]);
    ok(&dir, &[
        "set", "a.pptx", "/slide[1]/shape[2]",
        "--prop", "animation=fly-in", "--prop", "direction=left",
    ]);
    let slide = zip_text(&dir.join("a.pptx"), "ppt/slides/slide1.xml");
    assert!(slide.contains("presetID=\"2\""), "{slide}");
    assert!(slide.contains("presetSubtype=\"8\""), "fromLeft subtype: {slide}");
    assert!(slide.contains("0-#ppt_w/2") && slide.contains("#ppt_x"), "{slide}");
    assert!(slide.contains("ppt_y"), "{slide}");
    // Unknown directions are rejected.
    let err = fails(&dir, &[
        "set", "a.pptx", "/slide[1]/shape[2]",
        "--prop", "animation=fly-in", "--prop", "direction=sideways",
    ]);
    assert!(err.contains("unknown direction"), "{err}");
}

#[test]
fn formula_v05_functions() {
    let dir = temp_dir("v05formula");
    ok(&dir, &["create", "f.xlsx"]);
    ok(&dir, &["add", "f.xlsx", "/Sheet1", "--type", "row", "--prop", "values=2,10"]);
    ok(&dir, &["add", "f.xlsx", "/Sheet1", "--type", "row", "--prop", "values=3,20"]);
    let out = ok(&dir, &["calc", "f.xlsx", "=SUMPRODUCT(A1:A2,B1:B2)"]);
    assert!(out.trim().ends_with("80"), "{out}");
    let out = ok(&dir, &["calc", "f.xlsx", "=TEXTJOIN(\"/\",TRUE,A1:A2)"]);
    assert!(out.contains("2/3"), "{out}");
    let out = ok(&dir, &["calc", "f.xlsx", "=CEILING(A2+0.1,0.5)"]);
    assert!(out.contains("3.5"), "{out}");
}
