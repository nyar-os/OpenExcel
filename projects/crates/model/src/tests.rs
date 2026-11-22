//! End-to-end tests over the project model, the XLSX writer and the exports.

use std::{
    fs,
    path::PathBuf,
    sync::atomic::{AtomicU64, Ordering},
};

use serde_json::json;

use crate::{
    model::{CellRange, CellStyle, CellValue, Freeze, Project, SheetColumn, SheetVisibility},
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

#[test]
fn cell_and_range_reads_return_evaluated_values() {
    let created = crate::create(&json!({ "title": "Reads" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 2 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "value": 3 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A3", "formula": "=SUM(A1:A2)" },
                { "op": "sheet.add", "name": "Data" },
                { "op": "range.set", "sheet": "Data", "start": "A1", "rows": [[10], [20]] },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A4", "formula": "=SUM(Data!A1:A2)+A3" },
            ]
        }),
    )
    .expect("edit");

    let detail = crate::cell(handle, &json!({ "sheet": "Sheet1", "address": "A3" })).expect("cell");
    assert_eq!(detail["kind"], json!("formula"));
    assert_eq!(detail["formula"], json!("=SUM(A1:A2)"));
    assert_eq!(detail["value"], json!(5));
    assert_eq!(detail["cached"], json!(null));

    let cross = crate::cell(handle, &json!({ "address": "A4" })).expect("cell");
    assert_eq!(cross["value"], json!(35));

    let matrix = crate::range(handle, &json!({ "sheet": "Sheet1", "range": "A1:B2" })).expect("range");
    assert_eq!(matrix["range"], json!("A1:B2"));
    assert_eq!(matrix["rows"][0][0]["address"], json!("A1"));
    assert_eq!(matrix["rows"][0][0]["value"], json!(2));
    assert_eq!(matrix["rows"][1][0]["value"], json!(3));
    assert_eq!(matrix["rows"][0][1]["value"], json!(null));

    let recalc = crate::recalculate(handle).expect("recalc");
    assert_eq!(recalc["revision"], json!(1));
    assert_eq!(recalc["changed"], json!(0));
    crate::close(handle).expect("close");
}

#[test]
fn range_and_format_edits_apply() {
    let created = crate::create(&json!({ "title": "Edits" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let edited = crate::edit(
        handle,
        &json!({
            "operations": [
                {
                    "op": "range.set",
                    "sheet": "Sheet1",
                    "start": "B2",
                    "rows": [
                        [1, "two", true],
                        [null, { "formula": "=B2*2" }, 4],
                    ]
                },
                {
                    "op": "format.set",
                    "sheet": "Sheet1",
                    "range": "B2:D3",
                    "style": { "bold": true, "fill": "#ffeeaa", "align": "center", "numberFormat": "0.00" }
                },
            ]
        }),
    )
    .expect("edit");
    assert_eq!(edited["revision"], json!(1));

    let project = project_of(handle);
    let sheet = project.sheet("Sheet1").expect("sheet");
    assert_eq!(sheet.cell(2, 2).expect("B2").value, CellValue::number(1.0));
    assert_eq!(sheet.cell(2, 3).expect("C2").value, CellValue::text("two"));
    assert_eq!(sheet.cell(2, 4).expect("D2").value, CellValue::boolean(true));
    assert_eq!(sheet.cell(3, 2).map(|cell| cell.value.clone()), Some(CellValue::Empty), "a null input clears the value");
    assert_eq!(sheet.cell(3, 3).expect("C3").computed, Some(CellValue::number(2.0)));
    assert_eq!(sheet.cell(3, 4).expect("D3").value, CellValue::number(4.0));
    for (row, column) in [(2, 2), (3, 2), (3, 4)] {
        let style = sheet.cell(row, column).and_then(|cell| cell.style.as_ref()).expect("styled cell");
        assert_eq!(style.bold, Some(true));
        assert_eq!(style.fill.as_deref(), Some("#ffeeaa"));
        assert_eq!(style.align.as_deref(), Some("center"));
    }
    crate::close(handle).expect("close");
}

#[test]
fn duplicate_and_move_keep_sheet_names_unique() {
    let created = crate::create(&json!({ "title": "Sheets" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "sheet.add", "name": "Notes" },
                { "op": "cell.set", "sheet": "Notes", "reference": "A1", "value": "hi" },
                { "op": "sheet.duplicate", "sheet": "Notes" },
            ]
        }),
    )
    .expect("edit");
    let project = project_of(handle);
    assert_eq!(
        project.sheets.iter().map(|sheet| sheet.name.as_str()).collect::<Vec<_>>(),
        vec!["Sheet1", "Notes", "Notes Copy"]
    );
    let copy = project.sheet("Notes Copy").expect("copy");
    assert_ne!(copy.id, project.sheet("Notes").expect("source").id);
    assert_eq!(copy.cell(1, 1).expect("A1").value, CellValue::text("hi"));

    crate::edit(handle, &json!({ "operations": [{ "op": "sheet.move", "sheet": "Notes Copy", "index": 0 }] })).expect("move");
    let project = project_of(handle);
    assert_eq!(
        project.sheets.iter().map(|sheet| sheet.name.as_str()).collect::<Vec<_>>(),
        vec!["Notes Copy", "Sheet1", "Notes"]
    );
    assert_eq!(project.sheets.iter().map(|sheet| sheet.index).collect::<Vec<_>>(), vec![0, 1, 2]);

    let conflict =
        crate::edit(handle, &json!({ "operations": [{ "op": "sheet.duplicate", "sheet": "Notes", "name": "Sheet1" }] }))
            .expect_err("duplicate name");
    assert_eq!(conflict.code, "session/argument-invalid");
    crate::close(handle).expect("close");
}

#[test]
fn exports_write_evaluated_formula_values() {
    let created = crate::create(&json!({ "title": "Cached" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 2 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "formula": "=A1*3" },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "B1", "formula": "=1/0" },
                { "op": "format.set", "sheet": "Sheet1", "range": "D5", "style": { "bold": true } },
            ]
        }),
    )
    .expect("edit");

    let csv_path = temp_path("cached", "csv");
    crate::export(handle, "csv", &csv_path, true).expect("export csv");
    let text = fs::read_to_string(&csv_path).expect("read csv");
    assert!(text.starts_with("2,#DIV/0!\n6,\n"), "csv was `{text}`");
    let reopened = crate::open(&csv_path).expect("reopen csv");
    let csv_handle = reopened["handle"].as_u64().expect("handle");
    let csv_sheet = project_of(csv_handle);
    assert_eq!(csv_sheet.sheets[0].cell(2, 1).expect("A2").value, CellValue::number(6.0));
    assert_eq!(csv_sheet.sheets[0].cell(1, 2).expect("B1").value, CellValue::text("#DIV/0!"));
    crate::close(csv_handle).expect("close");

    let json_path = temp_path("cached", "json");
    crate::export(handle, "json", &json_path, true).expect("export json");
    let projection: serde_json::Value = serde_json::from_slice(&fs::read(&json_path).expect("read json")).expect("parse json");
    let cells = projection["sheets"][0]["cells"].as_array().expect("cells");
    let formula_cell = cells.iter().find(|cell| cell["address"] == json!("A2")).expect("A2");
    assert_eq!(formula_cell["formula"], json!("=A1*3"));
    assert_eq!(formula_cell["value"], json!(6));

    let xlsx_path = temp_path("cached", "xlsx");
    crate::export(handle, "xlsx", &xlsx_path, true).expect("export xlsx");
    let read = xlsx::read_project("cached.xlsx", &fs::read(&xlsx_path).expect("read xlsx"), "Cached").expect("read xlsx");
    let sheet = &read.sheets[0];
    assert_eq!(sheet.cell(2, 1).expect("A2").formula.as_deref(), Some("=A1*3"));
    assert_eq!(sheet.cell(2, 1).expect("A2").cached, Some(CellValue::number(6.0)));
    assert_eq!(sheet.cell(1, 2).expect("B1").cached, Some(CellValue::error("#DIV/0!")));
    assert_eq!(sheet.cell(5, 4).and_then(|cell| cell.style.as_ref()).and_then(|style| style.bold), Some(true));

    crate::close(handle).expect("close");
    for path in [csv_path, json_path, xlsx_path] {
        let _ = fs::remove_file(path);
    }
}

#[test]
fn old_project_json_without_computed_values_still_opens() {
    let path = temp_path("legacy", "oexcel");
    let legacy = json!({
        "format": "open-excel/project",
        "version": 0,
        "title": "Legacy",
        "sheets": [{
            "id": "sheet-1",
            "name": "Sheet1",
            "index": 0,
            "visibility": "visible",
            "cells": [
                { "reference": "A1", "row": 1, "column": 1, "value": { "kind": "number", "number": 4 } },
                { "reference": "B1", "row": 1, "column": 2, "value": { "kind": "empty" }, "formula": "=A1*2" }
            ]
        }]
    });
    fs::write(&path, serde_json::to_vec(&legacy).expect("serialize")).expect("write legacy");
    let opened = crate::open(&path).expect("open legacy");
    let handle = opened["handle"].as_u64().expect("handle");
    let detail = crate::cell(handle, &json!({ "address": "B1" })).expect("cell");
    assert_eq!(detail["formula"], json!("=A1*2"));
    assert_eq!(detail["value"], json!(8));
    crate::close(handle).expect("close");
    let _ = fs::remove_file(path);
}

#[test]
fn unsupported_formulas_keep_cached_values_through_inspect_and_exports() {
    let path = temp_path("unsupported", "oexcel");
    let project = json!({
        "format": "open-excel/project",
        "version": 0,
        "title": "Unsupported",
        "sheets": [{
            "id": "sheet-1",
            "name": "Sheet1",
            "index": 0,
            "visibility": "visible",
            "cells": [
                { "reference": "A1", "row": 1, "column": 1, "value": { "kind": "number", "number": 2 } },
                {
                    "reference": "B1",
                    "row": 1,
                    "column": 2,
                    "value": { "kind": "empty" },
                    "formula": "=FUTUREFN(A1)",
                    "cached": { "kind": "number", "number": 7 }
                }
            ]
        }]
    });
    fs::write(&path, serde_json::to_vec(&project).expect("serialize")).expect("write project");
    let opened = crate::open(&path).expect("open project");
    let handle = opened["handle"].as_u64().expect("handle");

    let report = crate::inspect(handle, Some(vec!["formulas".to_owned(), "types".to_owned()])).expect("inspect");
    assert_eq!(report["passed"], json!(true));
    let issues = report["issues"].as_array().expect("issues");
    let unsupported = issues.iter().find(|issue| issue["code"] == json!("formulas.unsupported-function")).expect("warning");
    assert_eq!(unsupported["severity"], json!("warning"));

    let detail = crate::cell(handle, &json!({ "address": "B1" })).expect("cell");
    assert_eq!(detail["value"], json!(7));
    assert_eq!(detail["computedState"], json!("unsupported"));

    let csv_path = temp_path("unsupported", "csv");
    crate::export(handle, "csv", &csv_path, true).expect("export csv");
    assert_eq!(fs::read_to_string(&csv_path).expect("read csv"), "2,7\n");

    let json_path = temp_path("unsupported", "json");
    crate::export(handle, "json", &json_path, true).expect("export json");
    let projection: serde_json::Value = serde_json::from_slice(&fs::read(&json_path).expect("read json")).expect("parse json");
    let cells = projection["sheets"][0]["cells"].as_array().expect("cells");
    let formula_cell = cells.iter().find(|cell| cell["address"] == json!("B1")).expect("B1");
    assert_eq!(formula_cell["value"], json!(7));

    let xlsx_path = temp_path("unsupported", "xlsx");
    crate::export(handle, "xlsx", &xlsx_path, true).expect("export xlsx");
    let read =
        xlsx::read_project("unsupported.xlsx", &fs::read(&xlsx_path).expect("read xlsx"), "Unsupported").expect("read xlsx");
    assert_eq!(read.sheets[0].cell(1, 2).expect("B1").cached, Some(CellValue::number(7.0)));

    crate::close(handle).expect("close");
    for path in [path, csv_path, json_path, xlsx_path] {
        let _ = fs::remove_file(path);
    }
}

#[test]
fn merge_freeze_visibility_and_replace_text_edits_apply() {
    let created = crate::create(&json!({ "title": "Structure" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let edited = crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "sheet.add", "name": "Data" },
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [["Alpha alpha ALPHA"], ["alpha ALPHA"]] },
                { "op": "merge.set", "sheet": "Sheet1", "start": "C1", "end": "D2", "merge": true },
                { "op": "sheet.freeze", "sheet": "Sheet1", "rows": 1, "columns": 1 },
                { "op": "sheet.setVisibility", "sheet": "Data", "visibility": "hidden" },
                { "op": "cell.replaceText", "sheet": "Sheet1", "range": "A1:A1", "find": "alpha", "replace": "x" },
                {
                    "op": "cell.replaceText",
                    "sheet": "Sheet1",
                    "range": "A2:A2",
                    "find": "alpha",
                    "replace": "y",
                    "matchCase": true
                }
            ]
        }),
    )
    .expect("edit");
    assert_eq!(edited["replacements"], json!(4));

    let project = project_of(handle);
    let sheet = project.sheet("Sheet1").expect("sheet");
    assert_eq!(sheet.cell(1, 1).expect("A1").value, CellValue::text("x x x"));
    assert_eq!(sheet.cell(2, 1).expect("A2").value, CellValue::text("y ALPHA"));
    assert_eq!(sheet.merged, vec![CellRange { start_row: 1, end_row: 2, start_column: 3, end_column: 4 }]);
    assert_eq!(sheet.freeze, Some(Freeze { columns: 1, rows: 1 }));
    assert_eq!(project.sheet("Data").expect("data").visibility, crate::model::SheetVisibility::Hidden);

    let overlap = crate::edit(
        handle,
        &json!({ "operations": [{ "op": "merge.set", "sheet": "Sheet1", "start": "C2", "end": "D3", "merge": true }] }),
    )
    .expect_err("overlap");
    assert_eq!(overlap.code, "session/argument-invalid");
    let outside = crate::edit(
        handle,
        &json!({ "operations": [{ "op": "merge.set", "sheet": "Sheet1", "start": "A1", "end": "XFE1", "merge": true }] }),
    )
    .expect_err("outside");
    assert_eq!(outside.code, "session/argument-invalid");

    crate::edit(
        handle,
        &json!({ "operations": [{ "op": "merge.set", "sheet": "Sheet1", "start": "C1", "end": "D2", "merge": false }] }),
    )
    .expect("unmerge");
    assert!(project_of(handle).sheet("Sheet1").expect("sheet").merged.is_empty());
    crate::close(handle).expect("close");
}

#[test]
fn find_reports_character_offsets_for_values_and_formulas() {
    let created = crate::create(&json!({ "title": "Search" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": "Find me find" },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "B1", "formula": "=LEFT(A1,4)" }
            ]
        }),
    )
    .expect("edit");

    let report = crate::find(handle, &json!({ "query": "find", "options": { "formulas": false } })).expect("find values");
    assert_eq!(report["truncated"], json!(false));
    let hits = report["hits"].as_array().expect("hits");
    assert_eq!(hits.len(), 3);
    assert_eq!(hits[0]["address"], json!("A1"));
    assert_eq!(hits[0]["field"], json!("value"));
    assert_eq!((hits[0]["start"].clone(), hits[0]["end"].clone()), (json!(0), json!(4)));
    assert_eq!(hits[0]["text"], json!("Find"));
    assert_eq!((hits[1]["start"].clone(), hits[1]["end"].clone()), (json!(8), json!(12)));
    assert_eq!(hits[1]["text"], json!("find"));
    assert_eq!(hits[2]["address"], json!("B1"));
    assert_eq!(hits[2]["text"], json!("Find"));

    let case_sensitive = crate::find(handle, &json!({ "query": "find", "options": { "matchCase": true } })).expect("case");
    let case_hits = case_sensitive["hits"].as_array().expect("hits");
    assert_eq!(case_hits.len(), 1);
    assert_eq!(case_hits[0]["start"], json!(8));

    let formulas = crate::find(handle, &json!({ "query": "LEFT", "options": { "values": false } })).expect("formulas");
    let formula_hits = formulas["hits"].as_array().expect("hits");
    assert_eq!(formula_hits.len(), 1);
    assert_eq!(formula_hits[0]["field"], json!("formula"));
    assert_eq!((formula_hits[0]["start"].clone(), formula_hits[0]["end"].clone()), (json!(1), json!(5)));

    let missing = crate::find(handle, &json!({ "query": "find", "sheet": "Nope" })).expect_err("missing sheet");
    assert_eq!(missing.code, "object/not-found");
    crate::close(handle).expect("close");
}

#[test]
fn find_reports_truncation_at_the_hit_cap() {
    let created = crate::create(&json!({ "title": "Truncate" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::store::with_mut(handle, |entry| {
        for row in 1..=620 {
            entry.project.sheets[0].write_cell(row, 1, CellValue::text("needle"), None, None);
        }
    })
    .expect("seed");
    let report = crate::find(handle, &json!({ "query": "needle" })).expect("find");
    assert_eq!(report["truncated"], json!(true));
    assert_eq!(report["hits"].as_array().expect("hits").len(), 500);
    crate::close(handle).expect("close");
}

#[test]
fn range_sort_orders_whole_rows_and_rejects_merges() {
    let created = crate::create(&json!({ "title": "Sort" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [
                    ["name", "qty", "double"],
                    ["b", 2, { "formula": "=B2*2" }],
                    ["a", 10, { "formula": "=B3*2" }],
                    ["c", 1, { "formula": "=B4*2" }],
                ] }
            ]
        }),
    )
    .expect("seed");

    let sorted = crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.sort", "sheet": "Sheet1", "range": "A1:C4", "column": 3, "ascending": true, "hasHeader": true }
            ]
        }),
    )
    .expect("sort by the formula column");
    assert_eq!(sorted["revision"], json!(2));

    let project = project_of(handle);
    let sheet = project.sheet("Sheet1").expect("sheet");
    assert_eq!(sheet.cell(1, 1).expect("A1").value, CellValue::text("name"));
    assert_eq!(sheet.cell(2, 1).expect("A2").value, CellValue::text("c"));
    assert_eq!(sheet.cell(3, 1).expect("A3").value, CellValue::text("b"));
    assert_eq!(sheet.cell(4, 1).expect("A4").value, CellValue::text("a"));
    // Formulas travel with their row, unchanged.
    assert_eq!(sheet.cell(2, 3).expect("C2").formula.as_deref(), Some("=B4*2"));
    assert_eq!(sheet.cell(3, 3).expect("C3").formula.as_deref(), Some("=B2*2"));
    assert_eq!(sheet.cell(4, 3).expect("C4").formula.as_deref(), Some("=B3*2"));
    assert_eq!(sheet.cell(2, 3).expect("C2").computed, Some(CellValue::number(20.0)));

    let outside = crate::edit(
        handle,
        &json!({ "operations": [{ "op": "range.sort", "sheet": "Sheet1", "range": "A1:C4", "column": 4 }] }),
    )
    .expect_err("column outside the range");
    assert_eq!(outside.code, "session/argument-invalid");

    crate::edit(
        handle,
        &json!({ "operations": [{ "op": "merge.set", "sheet": "Sheet1", "start": "E1", "end": "F2", "merge": true }] }),
    )
    .expect("merge");
    let merged = crate::edit(
        handle,
        &json!({ "operations": [{ "op": "range.sort", "sheet": "Sheet1", "range": "E1:F3", "column": 1 }] }),
    )
    .expect_err("merged range");
    assert_eq!(merged.code, "session/argument-invalid");
    crate::close(handle).expect("close");
}

#[test]
fn fill_series_writes_linear_numbers_and_rejects_bad_ranges() {
    let created = crate::create(&json!({ "title": "Series" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({ "operations": [{ "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [[2], [5]] }] }),
    )
    .expect("seed");
    crate::edit(
        handle,
        &json!({ "operations": [{ "op": "format.set", "sheet": "Sheet1", "range": "A3", "style": { "bold": true } }] }),
    )
    .expect("style");

    let edited = crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.fillSeries", "sheet": "Sheet1", "range": "A1:A5" },
                { "op": "range.fillSeries", "sheet": "Sheet1", "range": "C1:F1", "start": 10, "step": -2 },
            ]
        }),
    )
    .expect("fill");
    assert_eq!(edited["revision"], json!(3));
    assert_eq!(
        edited["affected"],
        json!([
            { "op": "range.fillSeries", "sheet": "Sheet1", "range": "A1:A5" },
            { "op": "range.fillSeries", "sheet": "Sheet1", "range": "C1:F1" },
        ])
    );

    let project = project_of(handle);
    let sheet = project.sheet("Sheet1").expect("sheet");
    let vertical = (1..=5)
        .map(|row| match sheet.cell(row, 1).map(|cell| cell.value.clone()) {
            Some(CellValue::Number { number }) => number,
            other => panic!("A{row} is {other:?}"),
        })
        .collect::<Vec<_>>();
    assert_eq!(vertical, vec![2.0, 5.0, 8.0, 11.0, 14.0]);
    let horizontal = (3..=6)
        .map(|column| match sheet.cell(1, column).map(|cell| cell.value.clone()) {
            Some(CellValue::Number { number }) => number,
            other => panic!("row 1 column {column} is {other:?}"),
        })
        .collect::<Vec<_>>();
    assert_eq!(horizontal, vec![10.0, 8.0, 6.0, 4.0]);
    assert_eq!(sheet.cell(3, 1).and_then(|cell| cell.style.as_ref()).and_then(|style| style.bold), Some(true));

    for request in [
        json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Sheet1", "range": "D1:E2" }] }),
        json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Sheet1", "range": "H1" }] }),
        json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Sheet1", "range": "H1:H3" }] }),
        json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Sheet1", "range": "A1:A3", "start": "x" }] }),
        json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Sheet1", "range": "A1:A3", "step": "two" }] }),
    ] {
        let error = crate::edit(handle, &request).expect_err("invalid series");
        assert_eq!(error.code, "session/argument-invalid");
    }
    let missing = crate::edit(
        handle,
        &json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Nope", "range": "A1:A3", "start": 1 }] }),
    )
    .expect_err("missing sheet");
    assert_eq!(missing.code, "object/not-found");

    crate::edit(
        handle,
        &json!({ "operations": [{ "op": "merge.set", "sheet": "Sheet1", "start": "J1", "end": "J2", "merge": true }] }),
    )
    .expect("merge");
    let merged = crate::edit(
        handle,
        &json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Sheet1", "range": "J1:J5", "start": 1 }] }),
    )
    .expect_err("merged range");
    assert_eq!(merged.code, "session/argument-invalid");
    crate::close(handle).expect("close");
}

#[test]
fn transpose_moves_cells_into_the_target_rectangle() {
    let created = crate::create(&json!({ "title": "Transpose" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [
                    ["a", "b", "c"],
                    [1, 2, { "formula": "=1+2" }],
                ] },
                { "op": "format.set", "sheet": "Sheet1", "range": "A1", "style": { "bold": true, "fill": "#ffeeaa" } },
            ]
        }),
    )
    .expect("seed");

    let edited = crate::edit(
        handle,
        &json!({
            "operations": [{ "op": "range.transpose", "sheet": "Sheet1", "range": "A1:C2", "target": "E1" }]
        }),
    )
    .expect("transpose");
    assert_eq!(edited["revision"], json!(2));
    assert_eq!(edited["affected"], json!([{ "op": "range.transpose", "sheet": "Sheet1", "range": "E1:F3" }]));

    let project = project_of(handle);
    let sheet = project.sheet("Sheet1").expect("sheet");
    assert!(sheet.cell(1, 1).is_none(), "the source rectangle is cleared");
    assert_eq!(sheet.cell(1, 5).expect("E1").value, CellValue::text("a"));
    assert_eq!(sheet.cell(1, 6).expect("F1").value, CellValue::number(1.0));
    assert_eq!(sheet.cell(2, 5).expect("E2").value, CellValue::text("b"));
    assert_eq!(sheet.cell(2, 6).expect("F2").value, CellValue::number(2.0));
    assert_eq!(sheet.cell(3, 5).expect("E3").value, CellValue::text("c"));
    let moved_formula = sheet.cell(3, 6).expect("F3");
    assert_eq!(moved_formula.formula.as_deref(), Some("=1+2"));
    assert_eq!(moved_formula.computed, Some(CellValue::number(3.0)));
    let style = sheet.cell(1, 5).and_then(|cell| cell.style.as_ref()).expect("E1 style");
    assert_eq!(style.bold, Some(true));
    assert_eq!(style.fill.as_deref(), Some("#ffeeaa"));

    let missing =
        crate::edit(handle, &json!({ "operations": [{ "op": "range.transpose", "sheet": "Sheet1", "range": "A1:B2" }] }))
            .expect_err("missing target");
    assert_eq!(missing.code, "session/argument-invalid");
    let overlap = crate::edit(
        handle,
        &json!({ "operations": [{ "op": "range.transpose", "sheet": "Sheet1", "range": "A1:C2", "target": "B1" }] }),
    )
    .expect_err("overlapping target");
    assert_eq!(overlap.code, "session/argument-invalid");
    let outside = crate::edit(
        handle,
        &json!({ "operations": [{ "op": "range.transpose", "sheet": "Sheet1", "range": "A1:B1", "target": "A1048576" }] }),
    )
    .expect_err("grid");
    assert_eq!(outside.code, "session/argument-invalid");

    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "H1", "rows": [[1, 2], [3, 4]] },
                { "op": "range.transpose", "sheet": "Sheet1", "range": "H1:I2", "target": "H1" },
            ]
        }),
    )
    .expect("square transpose in place");
    let project = project_of(handle);
    let sheet = project.sheet("Sheet1").expect("sheet");
    assert_eq!(sheet.cell(1, 8).expect("H1").value, CellValue::number(1.0));
    assert_eq!(sheet.cell(1, 9).expect("I1").value, CellValue::number(3.0));
    assert_eq!(sheet.cell(2, 8).expect("H2").value, CellValue::number(2.0));
    assert_eq!(sheet.cell(2, 9).expect("I2").value, CellValue::number(4.0));
    crate::close(handle).expect("close");
}

#[test]
fn pdf_export_writes_a_paginated_document() {
    let created = crate::create(&json!({ "title": "Report" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [
                    ["name", "qty"],
                    ["alpha", 10],
                    ["beta", 20],
                ] },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "C2", "formula": "=B2*3" },
                { "op": "format.set", "sheet": "Sheet1", "range": "A1:B1", "style": { "bold": true, "fill": "#ffeeaa" } },
                { "op": "merge.set", "sheet": "Sheet1", "start": "D1", "end": "E2", "merge": true },
                { "op": "sheet.add", "name": "Hidden" },
                { "op": "sheet.setVisibility", "sheet": "Hidden", "visibility": "hidden" },
            ]
        }),
    )
    .expect("edit");

    let path = temp_path("report", "pdf");
    let report = crate::export(handle, "pdf", &path, true).expect("export pdf");
    assert_eq!(report["format"], json!("pdf"));
    assert_eq!(report["editable"], json!(false));
    let losses = report["losses"].as_array().expect("losses");
    assert!(losses.iter().any(|loss| loss["code"] == json!("export.pdf.style-approximated")));

    let bytes = fs::read(&path).expect("read pdf");
    let text = String::from_utf8_lossy(&bytes);
    assert!(text.starts_with("%PDF-1.4"));
    assert!(text.contains("/Type /Catalog"));
    assert!(text.contains("Sheet1  Page 1 / 1"));
    assert!(!text.contains("Hidden  Page"));
    assert!(text.contains("/Count 1"));
    assert!(text.contains("xref"));
    assert!(text.contains("trailer"));
    assert!(text.contains("startxref"));
    assert!(text.trim_end().ends_with("%%EOF"));

    let csv_path = temp_path("report", "csv");
    crate::export(handle, "csv", &csv_path, true).expect("csv still works");
    assert!(fs::read_to_string(&csv_path).expect("csv").starts_with("name,qty"));

    crate::close(handle).expect("close");
    for path in [path, csv_path] {
        let _ = fs::remove_file(path);
    }
}

#[test]
fn capabilities_list_the_real_formula_functions() {
    let matrix = crate::capabilities();
    assert_eq!(matrix["formulas"]["unsupportedBehavior"], json!("preserve-cached"));
    let functions = matrix["formulas"]["functions"].as_array().expect("functions");
    for name in [
        "VLOOKUP",
        "XLOOKUP",
        "INDEX",
        "MATCH",
        "IFERROR",
        "TEXTJOIN",
        "ROUNDDOWN",
        "SUMIF",
        "SUMIFS",
        "COUNTIF",
        "COUNTIFS",
        "AVERAGEIF",
        "AVERAGEIFS",
        "MAXIFS",
        "MINIFS",
        "MEDIAN",
        "LARGE",
        "SMALL",
        "RANK",
        "DATE",
        "YEAR",
        "MONTH",
        "DAY",
        "WEEKDAY",
        "EDATE",
        "EOMONTH",
        "TODAY",
        "NOW",
        "IFS",
        "SWITCH",
        "CHOOSE",
        "REPLACE",
        "PMT",
        "FV",
        "PV",
        "NPV",
        "MODE",
    ] {
        assert!(functions.contains(&json!(name)), "`{name}` is missing from the capability list");
    }
    assert_eq!(matrix["exportTargets"]["pdf"], json!("partial"));
    assert_eq!(matrix["exportLimits"]["pdf"], json!(["style-approximated", "no-images"]));
    assert_eq!(matrix["renderers"]["pdf"], json!("unsupported"));
}

#[test]
fn sheet_freeze_uses_a_cell_boundary_and_survives_xlsx() {
    let created = crate::create(&json!({ "title": "Freeze" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [
                    [1, 2, 3],
                    [4, 5, 6],
                    [7, 8, 9],
                ] }
            ]
        }),
    )
    .expect("seed");

    let frozen = crate::edit(handle, &json!({ "operations": [{ "op": "sheet.freeze", "sheet": "Sheet1", "cell": "C2" }] }))
        .expect("freeze at C2");
    assert_eq!(frozen["revision"], json!(2));
    assert_eq!(project_of(handle).sheet("Sheet1").expect("sheet").freeze, Some(Freeze { rows: 1, columns: 2 }));

    let path = temp_path("freeze", "xlsx");
    crate::export(handle, "xlsx", &path, true).expect("export xlsx");
    let read = xlsx::read_project("freeze.xlsx", &fs::read(&path).expect("read xlsx"), "Freeze").expect("read xlsx");
    assert_eq!(read.sheets[0].freeze, Some(Freeze { rows: 1, columns: 2 }));

    let explicit = crate::edit(
        handle,
        &json!({ "operations": [{ "op": "sheet.freeze", "sheet": "Sheet1", "rows": 2, "columns": 1 }] }),
    )
    .expect("explicit rows and columns");
    assert_eq!(explicit["revision"], json!(3));
    assert_eq!(project_of(handle).sheet("Sheet1").expect("sheet").freeze, Some(Freeze { rows: 2, columns: 1 }));

    let cleared = crate::edit(handle, &json!({ "operations": [{ "op": "sheet.freeze", "sheet": "Sheet1", "clear": true }] }))
        .expect("clear freeze");
    assert_eq!(cleared["revision"], json!(4));
    assert_eq!(project_of(handle).sheet("Sheet1").expect("sheet").freeze, None);

    crate::export(handle, "xlsx", &path, true).expect("export cleared xlsx");
    let read = xlsx::read_project("cleared.xlsx", &fs::read(&path).expect("read xlsx"), "Freeze").expect("read cleared");
    assert_eq!(read.sheets[0].freeze, None);

    let implicit = crate::edit(handle, &json!({ "operations": [{ "op": "sheet.freeze", "sheet": "Sheet1" }] }))
        .expect("a request without boundaries clears the freeze");
    assert_eq!(implicit["revision"], json!(5));
    assert_eq!(project_of(handle).sheet("Sheet1").expect("sheet").freeze, None);

    crate::close(handle).expect("close");
    let _ = fs::remove_file(path);
}

#[test]
fn sheet_freeze_rejects_invalid_boundaries_without_changing_the_document() {
    let created = crate::create(&json!({ "title": "FreezeErrors" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [[1, 2, 3], [4, 5, 6], [7, 8, 9]] }
            ]
        }),
    )
    .expect("seed");
    let revision = crate::summary(handle).expect("summary")["revision"].clone();

    for request in [
        json!({ "operations": [{ "op": "sheet.freeze", "sheet": "Sheet1", "cell": "A0" }] }),
        json!({ "operations": [{ "op": "sheet.freeze", "sheet": "Sheet1", "cell": "XFE1" }] }),
        json!({ "operations": [{ "op": "sheet.freeze", "sheet": "Sheet1", "rows": 1_048_577, "columns": 0 }] }),
        json!({ "operations": [{ "op": "sheet.freeze", "sheet": "Sheet1", "rows": 99, "columns": 1 }] }),
        json!({ "operations": [{ "op": "sheet.freeze", "sheet": "Sheet1", "rows": 1 }] }),
        json!({
            "operations": [{ "op": "sheet.freeze", "sheet": "Sheet1", "cell": "C2", "rows": 1, "columns": 1 }]
        }),
    ] {
        let error = crate::edit(handle, &request).expect_err("invalid freeze");
        assert_eq!(error.code, "session/argument-invalid");
        assert_eq!(crate::summary(handle).expect("summary")["revision"], revision);
    }

    let error = crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": "changed" },
                { "op": "sheet.freeze", "sheet": "Sheet1", "rows": 99, "columns": 1 },
            ]
        }),
    )
    .expect_err("a failing operation rolls back the batch");
    assert_eq!(error.code, "session/argument-invalid");
    let project = project_of(handle);
    assert_eq!(project.sheet("Sheet1").expect("sheet").cell(1, 1).expect("A1").value, CellValue::number(1.0));
    assert_eq!(project.sheet("Sheet1").expect("sheet").freeze, None);
    assert_eq!(crate::summary(handle).expect("summary")["revision"], revision);
    crate::close(handle).expect("close");
}

#[test]
fn sheet_visibility_values_round_trip_through_xlsx() {
    let created = crate::create(&json!({ "title": "Visibility" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(
        handle,
        &json!({
            "operations": [
                { "op": "sheet.add", "name": "Hidden" },
                { "op": "sheet.add", "name": "Deep" },
                { "op": "sheet.setVisibility", "sheet": "Hidden", "visibility": "hidden" },
                { "op": "sheet.setVisibility", "sheet": "Deep", "visibility": "veryHidden" },
            ]
        }),
    )
    .expect("set visibility");

    let project = project_of(handle);
    assert_eq!(project.sheet("Sheet1").expect("sheet").visibility, SheetVisibility::Visible);
    assert_eq!(project.sheet("Hidden").expect("sheet").visibility, SheetVisibility::Hidden);
    assert_eq!(project.sheet("Deep").expect("sheet").visibility, SheetVisibility::VeryHidden);

    let path = temp_path("visibility", "xlsx");
    crate::export(handle, "xlsx", &path, true).expect("export xlsx");
    let read =
        xlsx::read_project("visibility.xlsx", &fs::read(&path).expect("read xlsx"), "Visibility").expect("read xlsx");
    assert_eq!(read.sheet("Sheet1").expect("sheet").visibility, SheetVisibility::Visible);
    assert_eq!(read.sheet("Hidden").expect("sheet").visibility, SheetVisibility::Hidden);
    assert_eq!(read.sheet("Deep").expect("sheet").visibility, SheetVisibility::VeryHidden);

    crate::edit(handle, &json!({ "operations": [
        { "op": "sheet.setVisibility", "sheet": "Hidden", "visibility": "visible" },
        { "op": "sheet.setVisibility", "sheet": "Deep", "visibility": "visible" },
    ] }))
    .expect("restore visibility");
    crate::export(handle, "xlsx", &path, true).expect("export restored");
    let read = xlsx::read_project("restored.xlsx", &fs::read(&path).expect("read xlsx"), "Visibility").expect("read restored");
    assert!(read.sheets.iter().all(|sheet| sheet.visibility == SheetVisibility::Visible));

    crate::close(handle).expect("close");
    let _ = fs::remove_file(path);
}

#[test]
fn hiding_the_last_visible_sheet_is_rejected() {
    let created = crate::create(&json!({ "title": "Visible" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    crate::edit(handle, &json!({ "operations": [{ "op": "sheet.add", "name": "Data" }] })).expect("add");
    crate::edit(handle, &json!({ "operations": [
        { "op": "sheet.setVisibility", "sheet": "Sheet1", "visibility": "hidden" },
    ] }))
    .expect("hide one of two sheets");
    let revision = crate::summary(handle).expect("summary")["revision"].clone();

    for visibility in ["hidden", "veryHidden"] {
        let error = crate::edit(
            handle,
            &json!({ "operations": [
                { "op": "sheet.setVisibility", "sheet": "Data", "visibility": visibility },
            ] }),
        )
        .expect_err("the last visible sheet cannot be hidden");
        assert_eq!(error.code, "session/argument-invalid");
        assert_eq!(crate::summary(handle).expect("summary")["revision"], revision);
    }
    let project = project_of(handle);
    assert_eq!(project.sheet("Data").expect("sheet").visibility, SheetVisibility::Visible);
    assert_eq!(project.sheet("Sheet1").expect("sheet").visibility, SheetVisibility::Hidden);

    let single = crate::create(&json!({ "title": "Single" })).expect("create");
    let single_handle = single["handle"].as_u64().expect("handle");
    let error = crate::edit(
        single_handle,
        &json!({ "operations": [{ "op": "sheet.setVisibility", "sheet": "Sheet1", "visibility": "hidden" }] }),
    )
    .expect_err("a single sheet cannot be hidden");
    assert_eq!(error.code, "session/argument-invalid");
    assert_eq!(project_of(single_handle).sheet("Sheet1").expect("sheet").visibility, SheetVisibility::Visible);

    crate::close(handle).expect("close");
    crate::close(single_handle).expect("close");
}
