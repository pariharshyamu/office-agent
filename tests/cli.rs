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
