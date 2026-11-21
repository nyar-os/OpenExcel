//! End-to-end tests over the project model, the XLSX writer and the exports.

use std::{
    fs,
    path::PathBuf,
    sync::atomic::{AtomicU64, Ordering},
};

use serde_json::json;

use crate::{
    model::{CellRange, CellStyle, CellValue, Freeze, Project, SheetColumn},
    xlsx,
};

static COUNTER: AtomicU64 = AtomicU64::new(0);

fn temp_path(name: &str, extension: &str) -> PathBuf {
    let directory = std::env::temp_dir().join(format!(
        "open-excel-tests-{}-{}",
        std::process::id(),
        COUNTER.fetch_add(1, Ordering::Relaxed)
    ));
    fs::create_dir_all(&directory).expect("create temp directory");
    directory.join(format!("{name}.{extension}"))
}

fn project_of(handle: u64) -> Project {
    crate::store::with(handle, |entry| entry.project.clone()).expect("open handle")
}

#[test]
fn create_edit_save_open_round_trip() {
    let path = temp_path("ledger", "oexcel");
    let created = crate::create(&json!({ "title": "Ledger" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");

    let edited = crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 42 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "formula": "=A1*2" },
                { "op": "sheet.add", "name": "Notes" },
                { "op": "sheet.rename", "sheet": "Notes", "name": "Readme" },
            ]
        }),
    )
    .expect("edit");
    assert_eq!(edited["revision"], json!(1));

    let report = crate::inspect(handle, Some(vec!["formulas".to_owned(), "references".to_owned()])).expect("inspect");
    assert_eq!(report["checkerVersion"], json!("open-excel/checker/1"));

    let saved = crate::save(handle, &path, true).expect("save");
    assert_eq!(saved["sheets"], json!(2));
    assert!(path.is_file());
    crate::close(handle).expect("close");
    assert_eq!(crate::summary(handle).expect_err("closed handle").code, "workbook/not-found");

    let opened = crate::open(&path).expect("open");
    let reopened = opened["handle"].as_u64().expect("handle");
    let project = project_of(reopened);
    assert_eq!(project.title, "Ledger");
    assert_eq!(project.sheets.len(), 2);
    assert_eq!(project.sheets[0].name, "Sheet1");
    assert_eq!(project.sheets[1].name, "Readme");
    let sheet = project.sheet("Sheet1").expect("sheet");
    assert_eq!(sheet.cell(1, 1).expect("A1").value, CellValue::number(42.0));
    assert_eq!(sheet.cell(2, 1).expect("A2").formula.as_deref(), Some("=A1*2"));

    crate::close(reopened).expect("close");
    let _ = fs::remove_file(path);
}

#[test]
fn handwritten_xlsx_round_trips_through_the_project() {
    let mut project = Project::blank("Workbook");
    let sheet = &mut project.sheets[0];
    sheet.write_cell(1, 1, CellValue::text("hello"), None, None);
    sheet.write_cell(1, 2, CellValue::number(3.5), None, None);
    sheet.write_cell(1, 3, CellValue::boolean(true), None, None);
    sheet.write_cell(2, 1, CellValue::Empty, Some("=B1*2".to_owned()), Some(CellValue::number(7.0)));
    sheet.cells[0].style = Some(CellStyle { bold: Some(true), font_size: Some(12.0), ..CellStyle::default() });
    sheet.merged.push(CellRange { start_row: 3, end_row: 4, start_column: 1, end_column: 2 });
    sheet.columns.push(SheetColumn { min: 1, max: 1, width: 24.0, custom: true, hidden: false });
    sheet.freeze = Some(Freeze { columns: 1, rows: 1 });
    sheet.dimension = sheet.actual_range();

    let bytes = xlsx::write_package(&project).expect("write package");
    assert!(bytes.starts_with(b"PK\x03\x04"));

    let read = xlsx::read_project("book.xlsx", &bytes, "Workbook").expect("read package");
    let read_sheet = &read.sheets[0];
    assert_eq!(read_sheet.cell(1, 1).expect("A1").value, CellValue::text("hello"));
    assert_eq!(read_sheet.cell(1, 2).expect("B1").value, CellValue::number(3.5));
    assert_eq!(read_sheet.cell(1, 3).expect("C1").value, CellValue::boolean(true));
    assert_eq!(read_sheet.cell(2, 1).expect("A2").formula.as_deref(), Some("=B1*2"));
    assert_eq!(read_sheet.cell(2, 1).expect("A2").cached, Some(CellValue::number(7.0)));
    assert_eq!(read_sheet.cell(1, 1).expect("A1").style.as_ref().and_then(|style| style.bold), Some(true));
    assert_eq!(read_sheet.merged, vec![CellRange { start_row: 3, end_row: 4, start_column: 1, end_column: 2 }]);
    assert_eq!(read_sheet.freeze, Some(Freeze { columns: 1, rows: 1 }));
    assert!((read_sheet.column_width(1) - 24.0).abs() < 0.001);
}

#[test]
fn csv_import_routes_through_open() {
    let path = temp_path("data", "csv");
    fs::write(&path, "id,note\n1,\"Doe, Jane\"\n2,plain\n").expect("write csv");
    let opened = crate::open(&path).expect("open csv");
    let handle = opened["handle"].as_u64().expect("handle");
    let project = project_of(handle);
    let sheet = &project.sheets[0];
    assert_eq!(sheet.name, "Sheet1");
    assert_eq!(sheet.cell(2, 1).expect("id").value, CellValue::number(1.0));
    assert_eq!(sheet.cell(2, 2).expect("note").value, CellValue::text("Doe, Jane"));
    assert_eq!(sheet.cell(3, 2).expect("note").value, CellValue::text("plain"));
    assert_eq!(project.cell_count(), 6);
    crate::close(handle).expect("close");
    let _ = fs::remove_file(path);
}

#[test]
fn formula_cycle_detection_runs_on_the_editable_model() {
    let created = crate::create(&json!({ "title": "Cycle" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "formula": "=B1" },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "B1", "formula": "=A1" },
            ]
        }),
    )
    .expect("edit");
    let report = crate::inspect(handle, Some(vec!["cycles".to_owned()])).expect("inspect");
    assert_eq!(report["passed"], json!(false));
    assert_eq!(report["issues"][0]["code"], json!("inspect.cycles.cycle"));
    crate::close(handle).expect("close");
}

#[test]
fn revision_conflicts_are_reported() {
    let created = crate::create(&json!({ "title": "Conflict" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(handle, &json!({ "operations": [{ "op": "sheet.add", "name": "Second" }] })).expect("edit");
    let error = crate::edit(handle, &json!({ "expectedRevision": 0, "operations": [{ "op": "sheet.add", "name": "Third" }] }))
        .expect_err("conflict");
    assert_eq!(error.code, "document/revision-conflict");
    crate::close(handle).expect("close");
}

#[test]
fn exports_cover_csv_json_and_xlsx() {
    let created = crate::create(&json!({ "title": "Export" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": "name" },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "B1", "value": 2 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "value": "quote, \"and\" newline" },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "B2", "formula": "=B1*2", "value": 4 },
            ]
        }),
    )
    .expect("edit");

    let csv_path = temp_path("export", "csv");
    let report = crate::export(handle, "csv", &csv_path, true).expect("export csv");
    assert_eq!(report["editable"], json!(true));
    let text = fs::read_to_string(&csv_path).expect("read csv");
    assert!(text.starts_with("name,2\n"));
    assert!(text.contains("\"quote, \"\"and\"\" newline\",4\n"));

    let json_path = temp_path("export", "json");
    crate::export(handle, "json", &json_path, true).expect("export json");
    let projection: serde_json::Value = serde_json::from_slice(&fs::read(&json_path).expect("read json")).expect("parse json");
    assert_eq!(projection["sheets"][0]["name"], json!("Sheet1"));
    assert_eq!(projection["sheets"][0]["cells"][0]["address"], json!("A1"));

    let xlsx_path = temp_path("export", "xlsx");
    let report = crate::export(handle, "xlsx", &xlsx_path, true).expect("export xlsx");
    assert_eq!(report["editable"], json!(true));
    let opened = crate::open(&xlsx_path).expect("reopen xlsx");
    let reopened = opened["handle"].as_u64().expect("handle");
    let project = project_of(reopened);
    let sheet = &project.sheets[0];
    assert_eq!(sheet.cell(1, 1).expect("A1").value, CellValue::text("name"));
    assert_eq!(sheet.cell(1, 2).expect("B1").value, CellValue::number(2.0));
    assert_eq!(sheet.cell(2, 2).expect("B2").formula.as_deref(), Some("=B1*2"));
    crate::close(reopened).expect("close");
    crate::close(handle).expect("close");

    for path in [csv_path, json_path, xlsx_path] {
        let _ = fs::remove_file(path);
    }
}

#[test]
fn svg_render_exposes_the_grid() {
    let created = crate::create(&json!({ "title": "Render" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(handle, &json!({ "operations": [{ "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": "<&>" }] }))
        .expect("edit");
    let rendered = crate::render(handle, &json!({})).expect("render");
    assert_eq!(rendered["format"], json!("svg"));
    let svg = rendered["svg"].as_str().expect("svg");
    assert!(svg.starts_with("<svg"));
    assert!(svg.contains("&lt;&amp;&gt;"));
    let unsupported = crate::render(handle, &json!({ "format": "png" })).expect_err("png");
    assert_eq!(unsupported.code, "format/unsupported");
    crate::close(handle).expect("close");
}
