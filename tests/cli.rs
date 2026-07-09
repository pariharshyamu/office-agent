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
