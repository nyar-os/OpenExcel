//! Layout edits that the CLI and the SDK expose through `edit`.

use std::{
    fs,
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
};

use open_excel_workbook as engine;
use serde_json::{Value, json};

static COUNTER: AtomicU64 = AtomicU64::new(0);

fn temp_path(name: &str, extension: &str) -> PathBuf {
    let directory = std::env::temp_dir().join(format!(
        "open-excel-editing-{}-{}",
        std::process::id(),
        COUNTER.fetch_add(1, Ordering::Relaxed)
    ));
    fs::create_dir_all(&directory).expect("create temp directory");
    directory.join(format!("{name}.{extension}"))
}

fn read_project_json(path: &Path) -> Value {
    let bytes = fs::read(path).expect("read project");
    serde_json::from_slice(&bytes).expect("parse project")
}

#[test]
fn column_widths_and_row_heights_land_in_the_project_json() {
    let project_path = temp_path("layout", "oexcel");
    let created = engine::create(&json!({ "title": "Layout" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let edited = engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "column.setWidth", "sheet": "Sheet1", "column": "A", "width": 24 },
                { "op": "row.setHeight", "sheet": "Sheet1", "row": 2, "height": 30, "hidden": true },
            ]
        }),
    )
    .expect("edit");
    assert_eq!(edited["revision"], json!(1));
    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    let columns = sheet["columns"].as_array().expect("columns");
    assert_eq!(columns.len(), 1);
    assert_eq!(columns[0]["min"], json!(1));
    assert_eq!(columns[0]["max"], json!(1));
    assert_eq!(columns[0]["width"], json!(24.0));
    assert_eq!(columns[0]["custom"], json!(true));
    assert_eq!(columns[0]["hidden"], json!(false));

    let rows = sheet["rows"].as_array().expect("rows");
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0]["index"], json!(2));
    assert_eq!(rows[0]["height"], json!(30.0));
    assert_eq!(rows[0]["customHeight"], json!(true));
    assert_eq!(rows[0]["hidden"], json!(true));

    let _ = fs::remove_file(project_path);
}

#[test]
fn setting_a_column_splits_an_existing_range() {
    let project_path = temp_path("split", "oexcel");
    let created = engine::create(&json!({ "title": "Split" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "column.setWidth", "sheet": "Sheet1", "column": "A", "width": 12 },
                { "op": "column.setWidth", "sheet": "Sheet1", "column": 2, "width": 12 },
                { "op": "column.setWidth", "sheet": "Sheet1", "column": "C", "width": 12 },
            ]
        }),
    )
    .expect("set width on A:C");
    engine::save(handle, &project_path, true).expect("save");

    let project = read_project_json(&project_path);
    let columns = project["sheets"][0]["columns"].as_array().expect("columns");
    assert_eq!(columns.len(), 1, "adjacent equal widths merge into one range");
    assert_eq!(columns[0]["min"], json!(1));
    assert_eq!(columns[0]["max"], json!(3));

    engine::edit(
        handle,
        &json!({
            "operations": [{ "op": "column.setWidth", "sheet": "Sheet1", "column": "B", "width": 20 }]
        }),
    )
    .expect("split middle column");
    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");

    let project = read_project_json(&project_path);
    let columns = project["sheets"][0]["columns"].as_array().expect("columns");
    assert_eq!(columns.len(), 3);
    assert_eq!(
        (columns[0]["min"].clone(), columns[0]["max"].clone(), columns[0]["width"].clone()),
        (json!(1), json!(1), json!(12.0))
    );
    assert_eq!(
        (columns[1]["min"].clone(), columns[1]["max"].clone(), columns[1]["width"].clone()),
        (json!(2), json!(2), json!(20.0))
    );
    assert_eq!(
        (columns[2]["min"].clone(), columns[2]["max"].clone(), columns[2]["width"].clone()),
        (json!(3), json!(3), json!(12.0))
    );

    let _ = fs::remove_file(project_path);
}

#[test]
fn invalid_layout_values_are_rejected() {
    let created = engine::create(&json!({ "title": "Invalid" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let requests = [
        json!({ "operations": [{ "op": "column.setWidth", "sheet": "Sheet1", "column": "A", "width": 0.05 }] }),
        json!({ "operations": [{ "op": "column.setWidth", "sheet": "Sheet1", "column": "A", "width": 300 }] }),
        json!({ "operations": [{ "op": "column.setWidth", "sheet": "Sheet1", "column": "XFE", "width": 10 }] }),
        json!({ "operations": [{ "op": "column.setWidth", "sheet": "Sheet1", "column": "A" }] }),
        json!({ "operations": [{ "op": "row.setHeight", "sheet": "Sheet1", "row": 1, "height": -1 }] }),
        json!({ "operations": [{ "op": "row.setHeight", "sheet": "Sheet1", "row": 1, "height": 500 }] }),
        json!({ "operations": [{ "op": "row.setHeight", "sheet": "Sheet1", "row": 0, "height": 10 }] }),
    ];
    for request in requests {
        let error = engine::edit(handle, &request).expect_err("invalid layout");
        assert_eq!(error.code, "session/argument-invalid");
    }
    engine::close(handle).expect("close");
}

#[test]
fn column_width_and_row_height_round_trip_through_xlsx() {
    let xlsx_path = temp_path("layout", "xlsx");
    let project_path = temp_path("layout-readback", "oexcel");
    let created = engine::create(&json!({ "title": "RoundTrip" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": "wide" },
                { "op": "column.setWidth", "sheet": "Sheet1", "column": "A", "width": 24 },
                { "op": "row.setHeight", "sheet": "Sheet1", "row": 1, "height": 30 },
            ]
        }),
    )
    .expect("edit");
    engine::export(handle, "xlsx", &xlsx_path, true).expect("export");
    engine::close(handle).expect("close");

    let opened = engine::open(&xlsx_path).expect("open xlsx");
    let read_handle = opened["handle"].as_u64().expect("handle");
    engine::save(read_handle, &project_path, true).expect("save");
    engine::close(read_handle).expect("close");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    let columns = sheet["columns"].as_array().expect("columns");
    assert!(
        columns.iter().any(|record| record["min"] == json!(1)
            && record["max"] == json!(1)
            && (record["width"].as_f64().unwrap_or_default() - 24.0).abs() < 0.001),
        "column A keeps its width"
    );
    let rows = sheet["rows"].as_array().expect("rows");
    assert!(
        rows.iter()
            .any(|record| record["index"] == json!(1) && (record["height"].as_f64().unwrap_or_default() - 30.0).abs() < 0.001),
        "row 1 keeps its height"
    );

    let _ = fs::remove_file(xlsx_path);
    let _ = fs::remove_file(project_path);
}

fn cell_json<'a>(sheet: &'a Value, reference: &str) -> &'a Value {
    sheet["cells"]
        .as_array()
        .expect("cells")
        .iter()
        .find(|cell| cell["reference"] == json!(reference))
        .unwrap_or_else(|| panic!("cell {reference} exists"))
}

#[test]
fn inserting_rows_moves_cells_merges_freeze_layout_and_formulas() {
    let project_path = temp_path("insert-rows", "oexcel");
    let created = engine::create(&json!({ "title": "InsertRows" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 1 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "value": 2 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A3", "value": 3 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "B3", "formula": "=$A$3+A1" },
                { "op": "merge.set", "sheet": "Sheet1", "start": "A5", "end": "B5", "merge": true },
                { "op": "sheet.freeze", "sheet": "Sheet1", "rows": 1, "columns": 1 },
                { "op": "column.setWidth", "sheet": "Sheet1", "column": "A", "width": 24 },
                { "op": "row.setHeight", "sheet": "Sheet1", "row": 2, "height": 30 },
            ]
        }),
    )
    .expect("seed");
    let inserted =
        engine::edit(handle, &json!({ "operations": [{ "op": "row.insert", "sheet": "Sheet1", "row": 2, "count": 2 }] }))
            .expect("insert rows");
    assert_eq!(inserted["revision"], json!(2));
    engine::save(handle, &project_path, true).expect("save");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A1")["value"]["number"], json!(1.0));
    assert_eq!(cell_json(sheet, "A4")["value"]["number"], json!(2.0));
    assert_eq!(cell_json(sheet, "A5")["value"]["number"], json!(3.0));
    assert_eq!(cell_json(sheet, "B5")["formula"], json!("=$A$5+A1"));
    assert_eq!(sheet["merged"], json!([{ "startRow": 7, "endRow": 7, "startColumn": 1, "endColumn": 2 }]));
    assert_eq!(sheet["freeze"], json!({ "columns": 1, "rows": 1 }));
    let columns = sheet["columns"].as_array().expect("columns");
    assert_eq!(
        (columns[0]["min"].clone(), columns[0]["max"].clone(), columns[0]["width"].clone()),
        (json!(1), json!(1), json!(24.0))
    );
    let rows = sheet["rows"].as_array().expect("rows");
    assert_eq!((rows[0]["index"].clone(), rows[0]["height"].clone()), (json!(4), json!(30.0)));

    let detail = engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "B5" })).expect("cell");
    assert_eq!(detail["formula"], json!("=$A$5+A1"));
    assert_eq!(detail["value"], json!(4));
    engine::close(handle).expect("close");
    let _ = fs::remove_file(project_path);
}

#[test]
fn deleting_rows_shrinks_merges_and_rewrites_references() {
    let project_path = temp_path("delete-rows", "oexcel");
    let created = engine::create(&json!({ "title": "DeleteRows" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 1 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "value": 2 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A3", "value": 3 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A4", "value": 4 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A5", "value": 5 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A6", "value": 6 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "B1", "formula": "=SUM(A1:A6)" },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "C1", "formula": "=A5" },
                { "op": "merge.set", "sheet": "Sheet1", "start": "A3", "end": "B6", "merge": true },
                { "op": "merge.set", "sheet": "Sheet1", "start": "D5", "end": "E5", "merge": true },
                { "op": "sheet.freeze", "sheet": "Sheet1", "rows": 3, "columns": 0 },
                { "op": "row.setHeight", "sheet": "Sheet1", "row": 2, "height": 20 },
                { "op": "row.setHeight", "sheet": "Sheet1", "row": 5, "height": 40 },
            ]
        }),
    )
    .expect("seed");
    engine::edit(handle, &json!({ "operations": [{ "op": "row.delete", "sheet": "Sheet1", "row": 5, "count": 2 }] }))
        .expect("delete rows");
    engine::save(handle, &project_path, true).expect("save");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A4")["value"]["number"], json!(4.0));
    assert!(sheet["cells"].as_array().expect("cells").iter().all(|cell| cell["reference"] != json!("A5")));
    assert_eq!(cell_json(sheet, "B1")["formula"], json!("=SUM(A1:A4)"));
    assert_eq!(cell_json(sheet, "C1")["formula"], json!("=#REF!"));
    assert_eq!(sheet["merged"], json!([{ "startRow": 3, "endRow": 4, "startColumn": 1, "endColumn": 2 }]));
    assert_eq!(sheet["freeze"], json!({ "columns": 0, "rows": 3 }));
    let rows = sheet["rows"].as_array().expect("rows");
    assert_eq!(rows.len(), 1);
    assert_eq!((rows[0]["index"].clone(), rows[0]["height"].clone()), (json!(2), json!(20.0)));

    let detail = engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "B1" })).expect("cell");
    assert_eq!(detail["value"], json!(10));
    engine::close(handle).expect("close");
    let _ = fs::remove_file(project_path);
}

#[test]
fn inserting_columns_moves_cells_widths_and_formulas() {
    let project_path = temp_path("insert-columns", "oexcel");
    let created = engine::create(&json!({ "title": "InsertColumns" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 1 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "B1", "value": 2 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "C1", "value": 3 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "B2", "formula": "=B1+$C$1" },
                { "op": "merge.set", "sheet": "Sheet1", "start": "B2", "end": "C3", "merge": true },
                { "op": "sheet.freeze", "sheet": "Sheet1", "rows": 1, "columns": 2 },
                { "op": "column.setWidth", "sheet": "Sheet1", "column": "B", "width": 20 },
                { "op": "column.setWidth", "sheet": "Sheet1", "column": "D", "width": 12 },
                { "op": "row.setHeight", "sheet": "Sheet1", "row": 1, "height": 30 },
            ]
        }),
    )
    .expect("seed");
    engine::edit(handle, &json!({ "operations": [{ "op": "column.insert", "sheet": "Sheet1", "column": "B", "count": 2 }] }))
        .expect("insert columns");
    engine::save(handle, &project_path, true).expect("save");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A1")["value"]["number"], json!(1.0));
    assert_eq!(cell_json(sheet, "D1")["value"]["number"], json!(2.0));
    assert_eq!(cell_json(sheet, "E1")["value"]["number"], json!(3.0));
    assert_eq!(cell_json(sheet, "D2")["formula"], json!("=D1+$E$1"));
    assert_eq!(sheet["merged"], json!([{ "startRow": 2, "endRow": 3, "startColumn": 4, "endColumn": 5 }]));
    assert_eq!(sheet["freeze"], json!({ "columns": 4, "rows": 1 }));
    let columns = sheet["columns"].as_array().expect("columns");
    assert_eq!(columns.len(), 2);
    assert_eq!(
        (columns[0]["min"].clone(), columns[0]["max"].clone(), columns[0]["width"].clone()),
        (json!(4), json!(4), json!(20.0))
    );
    assert_eq!(
        (columns[1]["min"].clone(), columns[1]["max"].clone(), columns[1]["width"].clone()),
        (json!(6), json!(6), json!(12.0))
    );
    let rows = sheet["rows"].as_array().expect("rows");
    assert_eq!((rows[0]["index"].clone(), rows[0]["height"].clone()), (json!(1), json!(30.0)));

    let detail = engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "D2" })).expect("cell");
    assert_eq!(detail["value"], json!(5));
    engine::close(handle).expect("close");
    let _ = fs::remove_file(project_path);
}

#[test]
fn deleting_columns_drops_merges_widths_and_marks_removed_references() {
    let project_path = temp_path("delete-columns", "oexcel");
    let created = engine::create(&json!({ "title": "DeleteColumns" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 1 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "B1", "value": 2 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "C1", "value": 3 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "formula": "=A1+B1+C1" },
                { "op": "merge.set", "sheet": "Sheet1", "start": "B1", "end": "C1", "merge": true },
                { "op": "sheet.freeze", "sheet": "Sheet1", "rows": 1, "columns": 2 },
                { "op": "column.setWidth", "sheet": "Sheet1", "column": "B", "width": 20 },
                { "op": "column.setWidth", "sheet": "Sheet1", "column": "D", "width": 12 },
            ]
        }),
    )
    .expect("seed");
    engine::edit(handle, &json!({ "operations": [{ "op": "column.delete", "sheet": "Sheet1", "column": "B" }] }))
        .expect("delete column");
    engine::save(handle, &project_path, true).expect("save");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "B1")["value"]["number"], json!(3.0));
    assert_eq!(cell_json(sheet, "A2")["formula"], json!("=A1+#REF!+B1"));
    assert!(sheet["merged"].is_null(), "a merge shrunk to one cell is dropped");
    assert_eq!(sheet["freeze"], json!({ "columns": 1, "rows": 1 }));
    let columns = sheet["columns"].as_array().expect("columns");
    assert_eq!(columns.len(), 1);
    assert_eq!(
        (columns[0]["min"].clone(), columns[0]["max"].clone(), columns[0]["width"].clone()),
        (json!(3), json!(3), json!(12.0))
    );
    engine::close(handle).expect("close");
    let _ = fs::remove_file(project_path);
}

#[test]
fn structural_edits_reject_out_of_range_requests() {
    let created = engine::create(&json!({ "title": "OutOfRange" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({ "operations": [{ "op": "cell.set", "sheet": "Sheet1", "reference": "A1048576", "value": "edge" }] }),
    )
    .expect("seed edge");
    let requests = [
        json!({ "operations": [{ "op": "row.insert", "sheet": "Sheet1", "row": 1_048_576, "count": 2 }] }),
        json!({ "operations": [{ "op": "row.delete", "sheet": "Sheet1", "row": 1_048_570, "count": 10 }] }),
        json!({ "operations": [{ "op": "row.insert", "sheet": "Sheet1", "row": 0 }] }),
        json!({ "operations": [{ "op": "row.delete", "sheet": "Sheet1", "row": 1, "count": 0 }] }),
        json!({ "operations": [{ "op": "row.insert", "sheet": "Sheet1", "row": 1_048_576 }] }),
        json!({ "operations": [{ "op": "column.insert", "sheet": "Sheet1", "column": "XFE" }] }),
        json!({ "operations": [{ "op": "column.delete", "sheet": "Sheet1", "column": 16_380, "count": 10 }] }),
    ];
    for request in requests {
        let error = engine::edit(handle, &request).expect_err("out-of-range structural edit");
        assert_eq!(error.code, "session/argument-invalid");
    }
    engine::close(handle).expect("close");
}

#[test]
fn a_failing_operation_leaves_the_workbook_unchanged() {
    let project_path = temp_path("atomic", "oexcel");
    let created = engine::create(&json!({ "title": "Atomic" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let edited = engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 1 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "value": 2 },
                { "op": "merge.set", "sheet": "Sheet1", "start": "A3", "end": "A4", "merge": true },
                { "op": "sheet.freeze", "sheet": "Sheet1", "rows": 1, "columns": 0 },
            ]
        }),
    )
    .expect("seed");
    assert_eq!(edited["revision"], json!(1));
    let error = engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "row.insert", "sheet": "Sheet1", "row": 2 },
                { "op": "row.delete", "sheet": "Sheet1", "row": 1_048_570, "count": 10 },
            ]
        }),
    )
    .expect_err("second operation fails");
    assert_eq!(error.code, "session/argument-invalid");
    assert_eq!(engine::summary(handle).expect("summary")["revision"], json!(1));

    engine::save(handle, &project_path, true).expect("save");
    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A2")["value"]["number"], json!(2.0));
    assert_eq!(sheet["merged"], json!([{ "startRow": 3, "endRow": 4, "startColumn": 1, "endColumn": 1 }]));
    assert_eq!(sheet["freeze"], json!({ "columns": 0, "rows": 1 }));
    engine::close(handle).expect("close");
    let _ = fs::remove_file(project_path);
}

#[test]
fn structural_edits_rewrite_qualifying_references_on_other_sheets_only() {
    let created = engine::create(&json!({ "title": "CrossSheet" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 1 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "value": 2 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "B1", "formula": "=A2" },
                { "op": "sheet.add", "name": "Notes" },
                { "op": "cell.set", "sheet": "Notes", "reference": "A1", "formula": "=Sheet1!A2" },
                { "op": "cell.set", "sheet": "Notes", "reference": "A2", "formula": "=A1" },
            ]
        }),
    )
    .expect("seed");
    engine::edit(handle, &json!({ "operations": [{ "op": "row.insert", "sheet": "Sheet1", "row": 2 }] })).expect("insert row");
    let local = engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "B1" })).expect("local");
    assert_eq!(local["formula"], json!("=A3"));
    let qualified = engine::cell(handle, &json!({ "sheet": "Notes", "address": "A1" })).expect("qualified");
    assert_eq!(qualified["formula"], json!("=Sheet1!A3"));
    let unqualified = engine::cell(handle, &json!({ "sheet": "Notes", "address": "A2" })).expect("unqualified");
    assert_eq!(unqualified["formula"], json!("=A1"));
    engine::close(handle).expect("close");
}

#[test]
fn cell_clear_removes_values_and_can_keep_style() {
    let project_path = temp_path("cell-clear", "oexcel");
    let created = engine::create(&json!({ "title": "Clear" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "B2", "value": "Hello" },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "C3", "value": 42 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "D4", "value": "Styled" },
                { "op": "format.set", "sheet": "Sheet1", "range": "C3", "style": { "bold": true } },
                { "op": "format.set", "sheet": "Sheet1", "range": "D4", "style": { "bold": true } },
                { "op": "cell.clear", "sheet": "Sheet1", "reference": "B2" },
                { "op": "cell.clear", "sheet": "Sheet1", "reference": "C3", "keepStyle": true },
                { "op": "cell.clear", "sheet": "Sheet1", "reference": "D4" },
            ]
        }),
    )
    .expect("edit");

    let error = engine::edit(handle, &json!({ "operations": [{ "op": "cell.clear", "sheet": "Sheet1", "reference": "not-a-cell" }] }))
        .expect_err("bad reference");
    assert_eq!(error.code, "session/argument-invalid");

    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    let cells = sheet["cells"].as_array().expect("cells");
    assert!(!cells.iter().any(|cell| cell["reference"] == json!("B2")), "a cleared value cell is removed");
    assert!(!cells.iter().any(|cell| cell["reference"] == json!("D4")), "clearing without keepStyle drops the style too");
    let kept = cell_json(sheet, "C3");
    assert_eq!(kept["value"]["kind"], json!("empty"), "the value is cleared");
    assert_eq!(kept["style"]["bold"], json!(true), "keepStyle preserves formatting");

    let _ = fs::remove_file(project_path);
}

#[test]
fn format_number_lands_kind_codes_on_every_cell_in_the_range() {
    let project_path = temp_path("format-number", "oexcel");
    let created = engine::create(&json!({ "title": "NumberFormats" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 1234.5 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "value": 0.25 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A3", "value": 45000 },
                { "op": "format.number", "sheet": "Sheet1", "range": "A1", "kind": "currency" },
                { "op": "format.number", "sheet": "Sheet1", "range": "A2", "kind": "percent", "decimals": 1 },
                { "op": "format.number", "sheet": "Sheet1", "range": "A3", "kind": "date" },
                { "op": "format.number", "sheet": "Sheet1", "range": "B1", "kind": "plain", "decimals": 3 },
                { "op": "format.number", "sheet": "Sheet1", "range": "B2", "kind": "currency", "symbol": "US$", "decimals": 0 },
                { "op": "format.number", "sheet": "Sheet1", "range": "B3", "kind": "plain" },
            ]
        }),
    )
    .expect("edit");
    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A1")["style"]["numberFormat"], json!("$#,##0.00"));
    assert_eq!(cell_json(sheet, "A1")["value"]["number"], json!(1234.5), "the value stays untouched");
    assert_eq!(cell_json(sheet, "A2")["style"]["numberFormat"], json!("0.0%"));
    assert_eq!(cell_json(sheet, "A3")["style"]["numberFormat"], json!("yyyy-mm-dd"));
    assert_eq!(cell_json(sheet, "B1")["style"]["numberFormat"], json!("0.000"));
    assert_eq!(cell_json(sheet, "B2")["style"]["numberFormat"], json!("US$#,##0"));
    assert_eq!(cell_json(sheet, "B3")["style"]["numberFormat"], json!("General"), "plain without decimals is General");
    let _ = fs::remove_file(project_path);
}

#[test]
fn format_number_keeps_other_style_properties_and_fills_empty_cells() {
    let project_path = temp_path("format-number-merge", "oexcel");
    let created = engine::create(&json!({ "title": "NumberFormatMerge" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "format.set", "sheet": "Sheet1", "range": "A1:B1", "style": { "bold": true } },
                { "op": "format.number", "sheet": "Sheet1", "range": "A1:C1", "kind": "currency", "decimals": 2 },
            ]
        }),
    )
    .expect("edit");
    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A1")["style"]["bold"], json!(true), "the bold style survives");
    assert_eq!(cell_json(sheet, "A1")["style"]["numberFormat"], json!("$#,##0.00"));
    assert_eq!(cell_json(sheet, "C1")["style"]["numberFormat"], json!("$#,##0.00"), "an empty cell gains the format");
    let _ = fs::remove_file(project_path);
}

#[test]
fn format_number_rejects_invalid_requests() {
    let created = engine::create(&json!({ "title": "NumberFormatInvalid" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let requests = [
        json!({ "operations": [{ "op": "format.number", "sheet": "Sheet1", "range": "A1:B2" }] }),
        json!({ "operations": [{ "op": "format.number", "sheet": "Sheet1", "kind": "currency" }] }),
        json!({ "operations": [{ "op": "format.number", "sheet": "Sheet1", "range": "nope", "kind": "currency" }] }),
        json!({ "operations": [{ "op": "format.number", "sheet": "Sheet1", "range": "A1:B2", "kind": "accounting" }] }),
        json!({ "operations": [{ "op": "format.number", "sheet": "Sheet1", "range": "A1:B2", "kind": "" }] }),
        json!({ "operations": [{ "op": "format.number", "sheet": "Sheet1", "range": "A1:B2", "kind": "currency", "decimals": 13 }] }),
        json!({ "operations": [{ "op": "format.number", "sheet": "Sheet1", "range": "A1:B2", "kind": "currency", "decimals": -1 }] }),
        json!({ "operations": [{ "op": "format.number", "sheet": "Sheet1", "range": "A1:B2", "kind": "currency", "symbol": "" }] }),
        json!({ "operations": [{ "op": "format.number", "sheet": "Sheet1", "range": "A1:B2", "kind": "currency", "symbol": 7 }] }),
    ];
    for request in requests {
        let error = engine::edit(handle, &request).expect_err("invalid number format");
        assert_eq!(error.code, "session/argument-invalid");
    }
    engine::close(handle).expect("close");
}

#[test]
fn named_ranges_round_trip_and_redefine_case_insensitively() {
    let project_path = temp_path("names", "oexcel");
    let created = engine::create(&json!({ "title": "Names" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 1 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "value": 2 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A3", "value": 3 },
                { "op": "name.define", "name": "Sales", "sheet": "Sheet1", "range": "A1:A3" },
                { "op": "name.define", "name": "Tax", "sheet": "Sheet1", "range": "B1:B2" },
            ]
        }),
    )
    .expect("define");
    let summary = engine::summary(handle).expect("summary");
    assert_eq!(
        summary["names"],
        json!([
            { "name": "Sales", "sheet": "Sheet1", "range": "A1:A3" },
            { "name": "Tax", "sheet": "Sheet1", "range": "B1:B2" },
        ])
    );

    engine::edit(handle, &json!({ "operations": [{ "op": "name.define", "name": "sales", "sheet": "Sheet1", "range": "A2:A3" }] }))
        .expect("redefine");
    let summary = engine::summary(handle).expect("summary after redefine");
    assert_eq!(
        summary["names"],
        json!([
            { "name": "sales", "sheet": "Sheet1", "range": "A2:A3" },
            { "name": "Tax", "sheet": "Sheet1", "range": "B1:B2" },
        ]),
        "a case-insensitive redefine replaces the name"
    );

    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");
    let project = read_project_json(&project_path);
    assert_eq!(
        project["names"],
        json!([
            { "name": "sales", "sheet": "sheet-1", "range": { "startRow": 2, "endRow": 3, "startColumn": 1, "endColumn": 1 } },
            { "name": "Tax", "sheet": "sheet-1", "range": { "startRow": 1, "endRow": 2, "startColumn": 2, "endColumn": 2 } },
        ])
    );

    let reopened = engine::open(&project_path).expect("reopen");
    let read_handle = reopened["handle"].as_u64().expect("handle");
    assert_eq!(
        engine::summary(read_handle).expect("summary after reopen")["names"],
        json!([
            { "name": "sales", "sheet": "Sheet1", "range": "A2:A3" },
            { "name": "Tax", "sheet": "Sheet1", "range": "B1:B2" },
        ]),
        "names survive the reopen"
    );
    engine::edit(read_handle, &json!({ "operations": [{ "op": "name.remove", "name": "TAX" }] })).expect("remove");
    assert_eq!(engine::summary(read_handle).expect("summary")["names"].as_array().expect("names").len(), 1);
    let error =
        engine::edit(read_handle, &json!({ "operations": [{ "op": "name.remove", "name": "Tax" }] })).expect_err("remove twice");
    assert_eq!(error.code, "object/not-found");
    engine::close(read_handle).expect("close");
    let _ = fs::remove_file(project_path);
}

#[test]
fn text_case_rewrites_text_cells_and_skips_formulas_and_numbers() {
    let created = engine::create(&json!({ "title": "Case" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let edited = engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [["alpha beta"], ["Mixed Text"], [{ "formula": "=A1" }], [12]] },
                { "op": "range.setTextCase", "sheet": "Sheet1", "range": "A1:A4", "case": "upper" },
            ]
        }),
    )
    .expect("upper-case");
    assert_eq!(
        edited["affected"].as_array().and_then(|entries| entries.last()).cloned(),
        Some(json!({ "op": "range.setTextCase", "sheet": "Sheet1", "range": "A1:A4", "changed": 2 }))
    );
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A1" })).expect("A1")["value"], json!("ALPHA BETA"));
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A2" })).expect("A2")["value"], json!("MIXED TEXT"));

    engine::edit(handle, &json!({ "operations": [{ "op": "range.setTextCase", "sheet": "Sheet1", "range": "A1:A2", "case": "title" }] }))
        .expect("title-case");
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A1" })).expect("A1")["value"], json!("Alpha Beta"));
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A2" })).expect("A2")["value"], json!("Mixed Text"));
    assert_eq!(
        engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A3" })).expect("A3")["formula"],
        json!("=A1"),
        "a formula cell is skipped"
    );
    assert_eq!(
        engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A4" })).expect("A4")["value"].as_f64(),
        Some(12.0),
        "a number is not text"
    );

    for request in [
        json!({ "operations": [{ "op": "range.setTextCase", "sheet": "Sheet1", "range": "A1:A2", "case": "sideways" }] }),
        json!({ "operations": [{ "op": "range.setTextCase", "sheet": "Sheet1", "range": "nope", "case": "upper" }] }),
        json!({ "operations": [{ "op": "range.setTextCase", "sheet": "Sheet1", "range": "A1:A2" }] }),
    ] {
        let error = engine::edit(handle, &request).expect_err("invalid case request");
        assert_eq!(error.code, "session/argument-invalid");
    }
    engine::close(handle).expect("close");
}

#[test]
fn trim_text_strips_and_collapses_whitespace() {
    let created = engine::create(&json!({ "title": "Trim" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let edited = engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [["  alpha  beta  "], ["\tgamma\n"], [{ "formula": "=LEN(A1)" }]] },
                { "op": "range.trimText", "sheet": "Sheet1", "range": "A1:A3" },
            ]
        }),
    )
    .expect("trim");
    assert_eq!(
        edited["affected"].as_array().and_then(|entries| entries.last()).cloned(),
        Some(json!({ "op": "range.trimText", "sheet": "Sheet1", "range": "A1:A3", "changed": 2 }))
    );
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A1" })).expect("A1")["value"], json!("alpha  beta"));
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A2" })).expect("A2")["value"], json!("gamma"));

    let collapsed = engine::edit(
        handle,
        &json!({ "operations": [{ "op": "range.trimText", "sheet": "Sheet1", "range": "A1:A2", "collapse": true }] }),
    )
    .expect("collapse");
    assert_eq!(
        collapsed["affected"].as_array().and_then(|entries| entries.last()).cloned(),
        Some(json!({ "op": "range.trimText", "sheet": "Sheet1", "range": "A1:A2", "changed": 1 }))
    );
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A1" })).expect("A1")["value"], json!("alpha beta"));
    assert_eq!(
        engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A3" })).expect("A3")["formula"],
        json!("=LEN(A1)"),
        "a formula cell is skipped"
    );

    for request in [
        json!({ "operations": [{ "op": "range.trimText", "sheet": "Sheet1", "range": "nope" }] }),
        json!({ "operations": [{ "op": "range.trimText", "sheet": "Sheet1", "range": "A1:A2", "collapse": "yes" }] }),
    ] {
        let error = engine::edit(handle, &request).expect_err("invalid trim request");
        assert_eq!(error.code, "session/argument-invalid");
    }
    engine::close(handle).expect("close");
}

#[test]
fn defined_names_can_be_renamed_case_insensitively() {
    let created = engine::create(&json!({ "title": "Rename" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "name.define", "name": "Sales", "sheet": "Sheet1", "range": "A1:A3" },
                { "op": "name.define", "name": "Tax", "sheet": "Sheet1", "range": "B1:B2" },
            ]
        }),
    )
    .expect("define");
    engine::edit(handle, &json!({ "operations": [{ "op": "name.rename", "name": "sales", "to": "Revenue" }] }))
        .expect("rename case-insensitively");
    assert_eq!(
        engine::summary(handle).expect("summary")["names"],
        json!([
            { "name": "Revenue", "sheet": "Sheet1", "range": "A1:A3" },
            { "name": "Tax", "sheet": "Sheet1", "range": "B1:B2" },
        ])
    );

    let taken = engine::edit(handle, &json!({ "operations": [{ "op": "name.rename", "name": "Revenue", "to": "TAX" }] }))
        .expect_err("the target already exists");
    assert_eq!(taken.code, "session/argument-invalid");
    let missing = engine::edit(handle, &json!({ "operations": [{ "op": "name.rename", "name": "Missing", "to": "Other" }] }))
        .expect_err("unknown source name");
    assert_eq!(missing.code, "object/not-found");
    for request in [
        json!({ "operations": [{ "op": "name.rename", "name": "Revenue", "to": "A1" }] }),
        json!({ "operations": [{ "op": "name.rename", "name": "Revenue", "to": " " }] }),
        json!({ "operations": [{ "op": "name.rename", "name": "Revenue" }] }),
    ] {
        let error = engine::edit(handle, &request).expect_err("invalid rename request");
        assert_eq!(error.code, "session/argument-invalid");
    }
    engine::close(handle).expect("close");
}

#[test]
fn named_ranges_move_with_structural_edits_and_dying_sheets() {
    let project_path = temp_path("names-shift", "oexcel");
    let created = engine::create(&json!({ "title": "NamesShift" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "name.define", "name": "Data", "sheet": "Sheet1", "range": "A1:A3" },
                { "op": "name.define", "name": "Wide", "sheet": "Sheet1", "range": "A1:C1" },
                { "op": "name.define", "name": "Top", "sheet": "Sheet1", "range": "A5:A5" },
                { "op": "sheet.add", "name": "Notes" },
                { "op": "name.define", "name": "NotesRange", "sheet": "Notes", "range": "B2:B4" },
            ]
        }),
    )
    .expect("seed");
    engine::edit(handle, &json!({ "operations": [{ "op": "row.insert", "sheet": "Sheet1", "row": 2 }] })).expect("insert row");
    engine::edit(handle, &json!({ "operations": [{ "op": "column.insert", "sheet": "Sheet1", "column": "B" }] }))
        .expect("insert column");
    engine::edit(handle, &json!({ "operations": [{ "op": "row.delete", "sheet": "Sheet1", "row": 6 }] })).expect("delete Top row");
    engine::edit(handle, &json!({ "operations": [{ "op": "row.delete", "sheet": "Sheet1", "row": 2, "count": 2 }] }))
        .expect("shrink Data");
    engine::edit(handle, &json!({ "operations": [{ "op": "sheet.remove", "sheet": "Notes" }] })).expect("drop the named sheet");

    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");

    let project = read_project_json(&project_path);
    let names = project["names"].as_array().expect("names");
    assert_eq!(names.len(), 2, "a name fully inside a deleted band and a name on a removed sheet are dropped");
    let data = names.iter().find(|named| named["name"] == json!("Data")).expect("Data");
    assert_eq!(
        data["range"],
        json!({ "startRow": 1, "endRow": 2, "startColumn": 1, "endColumn": 1 }),
        "Data crossed the insert and shrank across the delete"
    );
    let wide = names.iter().find(|named| named["name"] == json!("Wide")).expect("Wide");
    assert_eq!(
        wide["range"],
        json!({ "startRow": 1, "endRow": 1, "startColumn": 1, "endColumn": 4 }),
        "Wide expanded across the column insert"
    );
    let _ = fs::remove_file(project_path);
}

#[test]
fn name_operations_reject_invalid_requests() {
    let created = engine::create(&json!({ "title": "NamesInvalid" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let argument_requests = [
        json!({ "operations": [{ "op": "name.define", "sheet": "Sheet1", "range": "A1:B2" }] }),
        json!({ "operations": [{ "op": "name.define", "name": "  ", "sheet": "Sheet1", "range": "A1:B2" }] }),
        json!({ "operations": [{ "op": "name.define", "name": "A1", "sheet": "Sheet1", "range": "A1:B2" }] }),
        json!({ "operations": [{ "op": "name.define", "name": "A1:B2", "sheet": "Sheet1", "range": "A1:B2" }] }),
        json!({ "operations": [{ "op": "name.define", "name": "Data", "sheet": "Sheet1" }] }),
        json!({ "operations": [{ "op": "name.define", "name": "Data", "sheet": "Sheet1", "range": "nope" }] }),
        json!({ "operations": [{ "op": "name.remove" }] }),
    ];
    for request in argument_requests {
        let error = engine::edit(handle, &request).expect_err("invalid name request");
        assert_eq!(error.code, "session/argument-invalid");
    }
    let error = engine::edit(handle, &json!({ "operations": [{ "op": "name.define", "name": "Data", "sheet": "Missing", "range": "A1:B2" }] }))
        .expect_err("missing sheet");
    assert_eq!(error.code, "object/not-found");
    engine::close(handle).expect("close");
}

#[test]
fn date_series_fill_writes_serials_with_a_date_format() {
    let project_path = temp_path("date-series", "oexcel");
    let created = engine::create(&json!({ "title": "DateSeries" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 45000 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "value": 45031 },
                { "op": "range.fillSeries", "sheet": "Sheet1", "range": "A1:A5", "kind": "date" },
                { "op": "range.fillSeries", "sheet": "Sheet1", "range": "B1:B3", "kind": "date", "start": 45000, "step": 7 },
                { "op": "range.fillSeries", "sheet": "Sheet1", "range": "C1:C3", "start": 1, "step": 2 },
            ]
        }),
    )
    .expect("edit");
    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A1")["value"]["number"], json!(45000.0));
    assert_eq!(cell_json(sheet, "A2")["value"]["number"], json!(45031.0), "the seeded second cell keeps the inferred step");
    assert_eq!(cell_json(sheet, "A3")["value"]["number"], json!(45062.0));
    assert_eq!(cell_json(sheet, "A5")["value"]["number"], json!(45124.0));
    assert_eq!(cell_json(sheet, "A5")["style"]["numberFormat"], json!("yyyy-mm-dd"));
    assert_eq!(cell_json(sheet, "B3")["value"]["number"], json!(45014.0), "weekly step");
    assert_eq!(cell_json(sheet, "B3")["style"]["numberFormat"], json!("yyyy-mm-dd"));
    assert_eq!(cell_json(sheet, "C1")["value"]["number"], json!(1.0), "the linear default is untouched by the date work");
    assert!(cell_json(sheet, "C3")["style"]["numberFormat"].is_null(), "a linear fill stamps no date format");
    let _ = fs::remove_file(project_path);
}

#[test]
fn growth_series_fill_multiplies_from_a_ratio_or_the_seeds() {
    let project_path = temp_path("growth-series", "oexcel");
    let created = engine::create(&json!({ "title": "GrowthSeries" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 2 },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A2", "value": 6 },
                { "op": "range.fillSeries", "sheet": "Sheet1", "range": "A1:A5", "kind": "growth" },
                { "op": "range.fillSeries", "sheet": "Sheet1", "range": "B1:B3", "kind": "growth", "start": 100, "ratio": 0.5 },
                { "op": "range.fillSeries", "sheet": "Sheet1", "range": "C1:C3", "kind": "growth", "start": 5 },
            ]
        }),
    )
    .expect("edit");
    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A1")["value"]["number"], json!(2.0));
    assert_eq!(cell_json(sheet, "A2")["value"]["number"], json!(6.0));
    assert_eq!(cell_json(sheet, "A3")["value"]["number"], json!(18.0), "the seeds infer a ratio of 3");
    assert_eq!(cell_json(sheet, "A5")["value"]["number"], json!(162.0));
    assert_eq!(cell_json(sheet, "B3")["value"]["number"], json!(25.0), "an explicit ratio may shrink the series");
    assert_eq!(cell_json(sheet, "C3")["value"]["number"], json!(20.0), "one seed defaults to doubling");
    let _ = fs::remove_file(project_path);
}

#[test]
fn growth_series_fill_rejects_invalid_ratios() {
    let created = engine::create(&json!({ "title": "GrowthSeriesInvalid" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let requests = [
        json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Sheet1", "range": "A1:A3", "kind": "growth", "start": 0 }] }),
        json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Sheet1", "range": "A1:A3", "kind": "growth", "start": 2, "ratio": 0 }] }),
        json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Sheet1", "range": "A1:A3", "kind": "linear", "start": 2, "ratio": 2 }] }),
    ];
    for request in requests {
        let error = engine::edit(handle, &request).expect_err("invalid growth request");
        assert_eq!(error.code, "session/argument-invalid");
    }
    engine::close(handle).expect("close");
}

#[test]
fn date_series_fill_rejects_invalid_kind_requests() {
    let created = engine::create(&json!({ "title": "DateSeriesInvalid" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let requests = [
        json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Sheet1", "range": "A1:A3", "kind": "weekday" }] }),
        json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Sheet1", "range": "A1:A3", "kind": 5 }] }),
        json!({ "operations": [{ "op": "range.fillSeries", "sheet": "Sheet1", "range": "A1:A3", "kind": "date", "start": "today" }] }),
    ];
    for request in requests {
        let error = engine::edit(handle, &request).expect_err("invalid series kind");
        assert_eq!(error.code, "session/argument-invalid");
    }
    engine::close(handle).expect("close");
}

#[test]
fn workbook_title_edits_show_up_in_summary_and_project() {
    let project_path = temp_path("workbook-title", "oexcel");
    let created = engine::create(&json!({ "title": "Original" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");

    let edited = engine::edit(
        handle,
        &json!({ "operations": [{ "op": "workbook.setTitle", "title": "  Quarterly plan  " }] }),
    )
    .expect("rename the workbook");
    assert_eq!(edited["workbook"]["title"], json!("Quarterly plan"));
    assert_eq!(edited["revision"], json!(1));

    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");
    let project = read_project_json(&project_path);
    assert_eq!(project["title"], json!("Quarterly plan"));

    let reopened = engine::open(&project_path).expect("reopen");
    let handle = reopened["handle"].as_u64().expect("handle");
    assert_eq!(reopened["workbook"]["title"], json!("Quarterly plan"));
    for request in [
        json!({ "operations": [{ "op": "workbook.setTitle" }] }),
        json!({ "operations": [{ "op": "workbook.setTitle", "title": "   " }] }),
        json!({ "operations": [{ "op": "workbook.setTitle", "title": 7 }] }),
    ] {
        let error = engine::edit(handle, &request).expect_err("invalid title");
        assert_eq!(error.code, "session/argument-invalid");
    }
    assert_eq!(engine::summary(handle).expect("summary")["title"], json!("Quarterly plan"));
    engine::close(handle).expect("close");
    let _ = fs::remove_file(project_path);
}

#[test]
fn remove_duplicates_keeps_the_first_row_and_compacts_the_range() {
    let project_path = temp_path("remove-duplicates", "oexcel");
    let created = engine::create(&json!({ "title": "Dedup" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [
                    ["Name", "Value"], ["alpha", 1], ["beta", 2], ["alpha", 1], ["alpha", 3]
                ] },
                { "op": "range.set", "sheet": "Sheet1", "start": "C1", "rows": [["x"], ["y"], ["x"]] },
            ]
        }),
    )
    .expect("seed rows");

    let edited = engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.removeDuplicates", "sheet": "Sheet1", "range": "A1:B5", "hasHeader": true, "keys": [1] },
                { "op": "range.removeDuplicates", "sheet": "Sheet1", "range": "C1:C3" },
            ]
        }),
    )
    .expect("remove duplicate rows");
    assert_eq!(edited["affected"][0]["op"], json!("range.removeDuplicates"));
    assert_eq!(edited["affected"][0]["range"], json!("A1:B5"));
    assert_eq!(edited["affected"][0]["removed"], json!(2), "two later alpha rows collapse into the first");
    assert_eq!(edited["affected"][1]["removed"], json!(1), "the default keys compare every column");

    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");
    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A1")["value"]["text"], json!("Name"), "the header stays in place");
    assert_eq!(cell_json(sheet, "A2")["value"]["text"], json!("alpha"));
    assert_eq!(cell_json(sheet, "B2")["value"]["number"], json!(1.0));
    assert_eq!(cell_json(sheet, "A3")["value"]["text"], json!("beta"));
    assert_eq!(cell_json(sheet, "B3")["value"]["number"], json!(2.0));
    assert!(
        sheet["cells"].as_array().unwrap().iter().all(|cell| cell["reference"] != "A4"),
        "the tail collapses: {sheet}"
    );
    assert_eq!(cell_json(sheet, "C1")["value"]["text"], json!("x"));
    assert_eq!(cell_json(sheet, "C2")["value"]["text"], json!("y"));
    assert!(
        sheet["cells"].as_array().unwrap().iter().all(|cell| cell["reference"] != "C3"),
        "the repeated x is removed: {sheet}"
    );
    let _ = fs::remove_file(project_path);
}

#[test]
fn remove_duplicates_rejects_bad_keys_and_merged_ranges() {
    let created = engine::create(&json!({ "title": "DedupErrors" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [["a", 1], ["a", 2]] },
                { "op": "merge.set", "sheet": "Sheet1", "start": "D1", "end": "E1", "merge": true },
            ]
        }),
    )
    .expect("seed");
    let revision = engine::summary(handle).expect("summary")["revision"].as_u64().expect("revision");
    for request in [
        json!({ "operations": [{ "op": "range.removeDuplicates", "sheet": "Sheet1", "range": "A1:B2", "keys": [] }] }),
        json!({ "operations": [{ "op": "range.removeDuplicates", "sheet": "Sheet1", "range": "A1:B2", "keys": ["Z"] }] }),
        json!({ "operations": [{ "op": "range.removeDuplicates", "sheet": "Sheet1", "range": "A1:B2", "keys": [9] }] }),
        json!({ "operations": [{ "op": "range.removeDuplicates", "sheet": "Sheet1", "range": "A1:B2", "keys": "A" }] }),
        json!({ "operations": [{ "op": "range.removeDuplicates", "sheet": "Sheet1", "range": "D1:E1" }] }),
        json!({ "operations": [{ "op": "range.removeDuplicates", "sheet": "Sheet1", "range": "not-a-range" }] }),
    ] {
        let error = engine::edit(handle, &request).expect_err("invalid dedup request");
        assert_eq!(error.code, "session/argument-invalid");
    }
    assert_eq!(engine::summary(handle).expect("summary")["revision"].as_u64().expect("revision"), revision);
    engine::close(handle).expect("close");
}

#[test]
fn remove_empty_rows_compacts_only_rows_without_cells() {
    let project_path = temp_path("remove-empty-rows", "oexcel");
    let created = engine::create(&json!({ "title": "Compact" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [["alpha", 1]] },
                { "op": "range.set", "sheet": "Sheet1", "start": "A3", "rows": [["beta", 2]] },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "A4", "value": "" },
                { "op": "merge.set", "sheet": "Sheet1", "start": "D1", "end": "E1", "merge": true },
            ]
        }),
    )
    .expect("seed rows");

    let edited =
        engine::edit(handle, &json!({ "operations": [{ "op": "range.removeEmptyRows", "sheet": "Sheet1", "range": "A1:B4" }] }))
            .expect("compact empty rows");
    assert_eq!(edited["affected"][0]["op"], json!("range.removeEmptyRows"));
    assert_eq!(edited["affected"][0]["range"], json!("A1:B4"));
    assert_eq!(edited["affected"][0]["removed"], json!(1));

    let revision = engine::summary(handle).expect("summary").get("revision").cloned().unwrap_or_default();
    engine::save(handle, &project_path, true).expect("save");
    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A1")["value"]["text"], json!("alpha"));
    assert_eq!(cell_json(sheet, "A2")["value"]["text"], json!("beta"), "the row after the gap moves up");
    assert_eq!(cell_json(sheet, "A3")["value"]["text"], json!(""), "an authored blank cell survives");
    assert!(
        sheet["cells"].as_array().unwrap().iter().all(|cell| !matches!(cell["reference"].as_str(), Some("A4" | "B4"))),
        "the vacated tail is cleared: {sheet}"
    );

    for request in [
        json!({ "operations": [{ "op": "range.removeEmptyRows", "sheet": "Sheet1", "range": "not-a-range" }] }),
        json!({ "operations": [{ "op": "range.removeEmptyRows", "sheet": "Sheet1", "range": "D1" }] }),
    ] {
        let error = engine::edit(handle, &request).expect_err("invalid empty-row request");
        assert_eq!(error.code, "session/argument-invalid");
    }
    let missing =
        engine::edit(handle, &json!({ "operations": [{ "op": "range.removeEmptyRows", "sheet": "Missing", "range": "A1:B2" }] }))
            .expect_err("unknown sheet");
    assert_eq!(missing.code, "object/not-found");
    assert_eq!(engine::summary(handle).expect("summary").get("revision").cloned().unwrap_or_default(), revision);
    engine::close(handle).expect("close");
    let _ = fs::remove_file(project_path);
}

#[test]
fn remove_empty_columns_compacts_only_columns_without_cells() {
    let project_path = temp_path("remove-empty-columns", "oexcel");
    let created = engine::create(&json!({ "title": "Compact columns" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [["a", 1], ["b", 2]] },
                { "op": "range.set", "sheet": "Sheet1", "start": "D1", "rows": [[4], [5]] },
                { "op": "merge.set", "sheet": "Sheet1", "start": "F1", "end": "G1", "merge": true },
            ]
        }),
    )
    .expect("seed columns");

    let edited = engine::edit(
        handle,
        &json!({ "operations": [{ "op": "range.removeEmptyColumns", "sheet": "Sheet1", "range": "A1:E2" }] }),
    )
    .expect("compact empty columns");
    assert_eq!(edited["affected"][0]["op"], json!("range.removeEmptyColumns"));
    assert_eq!(edited["affected"][0]["removed"], json!(2));

    engine::save(handle, &project_path, true).expect("save");
    let revision = engine::summary(handle).expect("summary").get("revision").cloned().unwrap_or_default();
    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A1")["value"]["text"], json!("a"));
    assert_eq!(cell_json(sheet, "B1")["value"]["number"], json!(1.0));
    assert_eq!(cell_json(sheet, "C1")["value"]["number"], json!(4.0), "the D column moves left");
    assert!(
        sheet["cells"].as_array().unwrap().iter().all(|cell| !matches!(cell["reference"].as_str(), Some("D1" | "E1" | "D2" | "E2"))),
        "the vacated right edge is cleared: {sheet}"
    );

    for request in [
        json!({ "operations": [{ "op": "range.removeEmptyColumns", "sheet": "Sheet1", "range": "bad" }] }),
        json!({ "operations": [{ "op": "range.removeEmptyColumns", "sheet": "Sheet1", "range": "F1:G1" }] }),
        json!({ "operations": [{ "op": "range.removeEmptyColumns", "sheet": "Missing", "range": "A1:B2" }] }),
    ] {
        let error = engine::edit(handle, &request).expect_err("invalid empty-column request");
        assert_ne!(error.code, "");
    }
    assert_eq!(engine::summary(handle).expect("summary").get("revision").cloned().unwrap_or_default(), revision);
    engine::close(handle).expect("close");
    let _ = fs::remove_file(project_path);
}

#[test]
fn fill_down_copies_values_styles_and_formulas() {
    let project_path = temp_path("fill-down", "oexcel");
    let created = engine::create(&json!({ "title": "Fill" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let edited = engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [
                    ["top", 1, { "formula": "=B1*2" }],
                    ["old", 2, null],
                    ["old", 3, null],
                    ["old", 4, null]
                ] },
                { "op": "format.set", "sheet": "Sheet1", "range": "A1:C1", "style": { "bold": true, "fill": "#ffcc00" } },
                { "op": "range.fill", "sheet": "Sheet1", "range": "A1:C4", "direction": " DOWN " },
            ]
        }),
    )
    .expect("fill down");
    assert_eq!(
        edited["affected"].as_array().and_then(|entries| entries.last()).cloned(),
        Some(json!({ "op": "range.fill", "sheet": "Sheet1", "range": "A1:C4" }))
    );

    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");
    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A1")["value"]["text"], json!("top"), "the source row is untouched");
    assert_eq!(cell_json(sheet, "A2")["value"]["text"], json!("top"));
    assert_eq!(cell_json(sheet, "A4")["value"]["text"], json!("top"));
    assert_eq!(cell_json(sheet, "B4")["value"]["number"], json!(1.0));
    assert_eq!(cell_json(sheet, "C4")["formula"], json!("=B1*2"), "formula text is copied verbatim");
    assert_eq!(cell_json(sheet, "C4")["reference"], json!("C4"), "the reference follows the target");
    assert_eq!(cell_json(sheet, "B3")["style"]["bold"], json!(true), "styles travel with the source cell");
    assert_eq!(cell_json(sheet, "C2")["style"]["fill"], json!("#ffcc00"));
    let _ = fs::remove_file(project_path);
}

#[test]
fn fill_clears_targets_of_blank_source_cells() {
    let project_path = temp_path("fill-blank", "oexcel");
    let created = engine::create(&json!({ "title": "FillBlank" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [[null, 5], ["stale", "stale"]] },
                { "op": "format.set", "sheet": "Sheet1", "range": "A2", "style": { "bold": true } },
                { "op": "range.fill", "sheet": "Sheet1", "range": "A1:B2", "direction": "down" },
            ]
        }),
    )
    .expect("fill down");

    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");
    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "B2")["value"]["number"], json!(5.0));
    assert!(
        sheet["cells"].as_array().unwrap().iter().all(|cell| cell["reference"] != "A2"),
        "a blank source cell clears its target, including the old style: {sheet}"
    );
    let _ = fs::remove_file(project_path);
}

#[test]
fn fill_right_left_and_up_copy_the_opposite_edge() {
    let created = engine::create(&json!({ "title": "FillEdges" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let edited = engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [["left", "old1", "old2"]] },
                { "op": "range.set", "sheet": "Sheet1", "start": "A3", "rows": [["old1", "old2", "right"]] },
                { "op": "range.set", "sheet": "Sheet1", "start": "A5", "rows": [["old1"], ["old2"], ["bottom"]] },
                { "op": "range.fill", "sheet": "Sheet1", "range": "A1:C1", "direction": "right" },
                { "op": "range.fill", "sheet": "Sheet1", "range": "A3:C3", "direction": "LEFT" },
                { "op": "range.fill", "sheet": "Sheet1", "range": "A5:A7", "direction": "up" },
            ]
        }),
    )
    .expect("fill edges");
    assert!(
        edited["affected"].as_array().unwrap().iter().all(|entry| entry["op"] == json!("range.fill")),
        "every fill is reported: {edited}"
    );
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "B1" })).expect("B1")["value"], json!("left"));
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "C1" })).expect("C1")["value"], json!("left"));
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A3" })).expect("A3")["value"], json!("right"));
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "B3" })).expect("B3")["value"], json!("right"));
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A5" })).expect("A5")["value"], json!("bottom"));
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A6" })).expect("A6")["value"], json!("bottom"));
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A7" })).expect("A7")["value"], json!("bottom"));
    engine::close(handle).expect("close");
}

#[test]
fn fill_rejects_bad_directions_and_merged_ranges_but_single_lines_are_no_ops() {
    let created = engine::create(&json!({ "title": "FillErrors" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [["one", "two", "three"]] },
                { "op": "merge.set", "sheet": "Sheet1", "start": "E1", "end": "F2", "merge": true },
            ]
        }),
    )
    .expect("seed");
    let revision = engine::summary(handle).expect("summary")["revision"].as_u64().expect("revision");
    for request in [
        json!({ "operations": [{ "op": "range.fill", "sheet": "Sheet1", "range": "A1:C1", "direction": "sideways" }] }),
        json!({ "operations": [{ "op": "range.fill", "sheet": "Sheet1", "range": "A1:C1", "direction": " " }] }),
        json!({ "operations": [{ "op": "range.fill", "sheet": "Sheet1", "range": "A1:C1" }] }),
        json!({ "operations": [{ "op": "range.fill", "sheet": "Sheet1", "range": "E1:F2", "direction": "down" }] }),
        json!({ "operations": [{ "op": "range.fill", "sheet": "Sheet1", "range": "nope", "direction": "down" }] }),
    ] {
        let error = engine::edit(handle, &request).expect_err("invalid fill request");
        assert_eq!(error.code, "session/argument-invalid");
    }
    assert_eq!(engine::summary(handle).expect("summary")["revision"].as_u64().expect("revision"), revision);

    let edited = engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.fill", "sheet": "Sheet1", "range": "A1:C1", "direction": "down" },
                { "op": "range.fill", "sheet": "Sheet1", "range": "A1:A3", "direction": "left" },
            ]
        }),
    )
    .expect("single-line fills are no-ops");
    assert_eq!(
        edited["affected"],
        json!([
            { "op": "range.fill", "sheet": "Sheet1", "range": "A1:C1" },
            { "op": "range.fill", "sheet": "Sheet1", "range": "A1:A3" },
        ])
    );
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "A1" })).expect("A1")["value"], json!("one"));
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "B1" })).expect("B1")["value"], json!("two"));
    assert_eq!(engine::cell(handle, &json!({ "sheet": "Sheet1", "address": "C1" })).expect("C1")["value"], json!("three"));
    engine::close(handle).expect("close");
}

#[test]
fn row_swap_moves_cells_heights_merges_and_formula_references() {
    let project_path = temp_path("row-swap", "oexcel");
    let created = engine::create(&json!({ "title": "RowSwap" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [["top", 1]] },
                { "op": "range.set", "sheet": "Sheet1", "start": "A3", "rows": [["bottom", 3]] },
                { "op": "row.setHeight", "sheet": "Sheet1", "row": 1, "height": 40 },
                { "op": "format.set", "sheet": "Sheet1", "range": "B3", "style": { "bold": true } },
                { "op": "merge.set", "sheet": "Sheet1", "start": "D1", "end": "E1", "merge": true },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "F1", "formula": "=A1+A3" },
                { "op": "row.swap", "sheet": "Sheet1", "row": 1, "other": 3 },
            ]
        }),
    )
    .expect("swap rows");
    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "A3")["value"]["text"], json!("top"));
    assert_eq!(cell_json(sheet, "B3")["value"]["number"], json!(1.0));
    assert_eq!(cell_json(sheet, "A1")["value"]["text"], json!("bottom"));
    assert_eq!(cell_json(sheet, "B1")["value"]["number"], json!(3.0));
    assert_eq!(cell_json(sheet, "B1")["style"]["bold"], json!(true), "styles travel with their cells");
    let rows = sheet["rows"].as_array().expect("rows");
    assert_eq!(rows.len(), 1);
    assert_eq!((rows[0]["index"].clone(), rows[0]["height"].clone()), (json!(3), json!(40.0)), "the height follows the row");
    let merged = sheet["merged"].as_array().expect("merged");
    assert_eq!(merged.len(), 1);
    assert_eq!((merged[0]["startRow"].clone(), merged[0]["endRow"].clone()), (json!(3), json!(3)));
    assert_eq!((merged[0]["startColumn"].clone(), merged[0]["endColumn"].clone()), (json!(4), json!(5)));
    assert_eq!(cell_json(sheet, "F3")["formula"], json!("=A3+A1"), "formula cells move and their references follow their rows");
    let _ = fs::remove_file(project_path);
}

#[test]
fn column_swap_moves_cells_widths_merges_and_formula_references() {
    let project_path = temp_path("column-swap", "oexcel");
    let created = engine::create(&json!({ "title": "ColumnSwap" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [["left", "right"]] },
                { "op": "column.setWidth", "sheet": "Sheet1", "column": "A", "width": 30 },
                { "op": "merge.set", "sheet": "Sheet1", "start": "B5", "end": "B6", "merge": true },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "C3", "formula": "=A1*2" },
                { "op": "column.swap", "sheet": "Sheet1", "column": "A", "other": "B" },
            ]
        }),
    )
    .expect("swap columns");
    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    assert_eq!(cell_json(sheet, "B1")["value"]["text"], json!("left"));
    assert_eq!(cell_json(sheet, "A1")["value"]["text"], json!("right"));
    let columns = sheet["columns"].as_array().expect("columns");
    let second = columns.iter().find(|record| record["min"] == json!(2)).expect("column B record");
    assert_eq!((second["width"].clone(), second["custom"].clone()), (json!(30.0), json!(true)), "the width follows the column");
    assert!(
        columns.iter().all(|record| record["min"] != json!(1) || record["custom"] != json!(true)),
        "column A keeps no custom width: {columns:?}"
    );
    let merged = sheet["merged"].as_array().expect("merged");
    assert_eq!(merged.len(), 1);
    assert_eq!((merged[0]["startColumn"].clone(), merged[0]["endColumn"].clone()), (json!(1), json!(1)));
    assert_eq!((merged[0]["startRow"].clone(), merged[0]["endRow"].clone()), (json!(5), json!(6)));
    assert_eq!(cell_json(sheet, "C3")["formula"], json!("=B1*2"), "references follow their columns");
    let _ = fs::remove_file(project_path);
}

#[test]
fn line_swap_rejects_bad_lines_unknown_sheets_and_straddling_merges() {
    let created = engine::create(&json!({ "title": "LineSwapErrors" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "merge.set", "sheet": "Sheet1", "start": "A2", "end": "A4", "merge": true },
                { "op": "merge.set", "sheet": "Sheet1", "start": "C1", "end": "D1", "merge": true },
            ]
        }),
    )
    .expect("seed merges");
    for request in [
        json!({ "operations": [{ "op": "row.swap", "sheet": "Sheet1", "row": 0, "other": 3 }] }),
        json!({ "operations": [{ "op": "row.swap", "sheet": "Sheet1", "row": 2 }] }),
        json!({ "operations": [{ "op": "row.swap", "sheet": "Sheet1", "row": 2, "other": 5 }] }),
        json!({ "operations": [{ "op": "row.swap", "sheet": "Nope", "row": 2, "other": 5 }] }),
        json!({ "operations": [{ "op": "column.swap", "sheet": "Sheet1", "column": "nope", "other": "D" }] }),
        json!({ "operations": [{ "op": "column.swap", "sheet": "Sheet1", "column": "C", "other": "E" }] }),
    ] {
        let error = engine::edit(handle, &request).expect_err("invalid swap request");
        assert_eq!(error.code, if request["operations"][0]["sheet"] == json!("Nope") { "object/not-found" } else { "session/argument-invalid" });
    }
    assert_eq!(engine::summary(handle).expect("summary")["revision"], json!(1), "rejected batches never advance");

    let edited = engine::edit(
        handle,
        &json!({ "operations": [{ "op": "row.swap", "sheet": "Sheet1", "row": 1, "other": 1 }] }),
    )
    .expect("swapping a row with itself is a no-op");
    assert_eq!(edited["revision"], json!(2));
    engine::close(handle).expect("close");
}

#[test]
fn range_insert_shifts_cells_down_or_right() {
    let project_path = temp_path("range-insert", "oexcel");
    let created = engine::create(&json!({ "title": "RangeInsert" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    let edited = engine::edit(
        handle,
        &json!({
            "operations": [
                { "op": "range.set", "sheet": "Sheet1", "start": "A1", "rows": [[1], [2], [3]] },
                { "op": "range.set", "sheet": "Sheet1", "start": "B1", "rows": [["x", "y"]] },
                { "op": "merge.set", "sheet": "Sheet1", "start": "A4", "end": "A5", "merge": true },
                { "op": "cell.set", "sheet": "Sheet1", "reference": "C1", "formula": "=SUM(A1:A3)" },
                { "op": "range.insert", "sheet": "Sheet1", "range": "A2:A2", "shift": "down" },
                { "op": "range.insert", "sheet": "Sheet1", "range": "B1:B1", "shift": "right" },
            ]
        }),
    )
    .expect("insert ranges");
    assert_eq!(
        edited["affected"],
        json!([
            { "op": "range.insert", "sheet": "Sheet1", "range": "A2:A2" },
            { "op": "range.insert", "sheet": "Sheet1", "range": "B1:B1" },
        ])
    );
    engine::save(handle, &project_path, true).expect("save");
    engine::close(handle).expect("close");

    let project = read_project_json(&project_path);
    let sheet = &project["sheets"][0];
    let cells = sheet["cells"].as_array().expect("cells");
    assert_eq!(cell_json(sheet, "A1")["value"]["number"], json!(1.0));
    assert!(cells.iter().all(|cell| cell["reference"] != json!("A2")), "the inserted row is empty");
    assert_eq!(cell_json(sheet, "A3")["value"]["number"], json!(2.0));
    assert_eq!(cell_json(sheet, "A4")["value"]["number"], json!(3.0));
    assert!(cells.iter().all(|cell| cell["reference"] != json!("B1")), "the inserted column is empty");
    assert_eq!(cell_json(sheet, "C1")["value"]["text"], json!("x"), "the right shift moved x from B1");
    assert_eq!(cell_json(sheet, "D1")["formula"], json!("=SUM(A1:A4)"), "the formula cell shifted right and the range follows the inserted row");
    let merged = sheet["merged"].as_array().expect("merged");
    assert_eq!(merged.len(), 1);
    assert_eq!((merged[0]["startRow"].clone(), merged[0]["endRow"].clone()), (json!(5), json!(6)), "the merge below the line moved down");
    let _ = fs::remove_file(project_path);
}

#[test]
fn range_insert_rejects_bad_requests() {
    let created = engine::create(&json!({ "title": "RangeInsertErrors" })).expect("create");
    let handle = created["handle"].as_u64().expect("handle");
    engine::edit(
        handle,
        &json!({ "operations": [{ "op": "merge.set", "sheet": "Sheet1", "start": "A1", "end": "A3", "merge": true }] }),
    )
    .expect("seed merge");
    for request in [
        json!({ "operations": [{ "op": "range.insert", "sheet": "Sheet1", "range": "A2:A2" }] }),
        json!({ "operations": [{ "op": "range.insert", "sheet": "Sheet1", "range": "A2:A2", "shift": "sideways" }] }),
        json!({ "operations": [{ "op": "range.insert", "sheet": "Sheet1", "range": "nope", "shift": "down" }] }),
        json!({ "operations": [{ "op": "range.insert", "sheet": "Nope", "range": "A2:A2", "shift": "down" }] }),
        json!({ "operations": [{ "op": "range.insert", "sheet": "Sheet1", "range": "A2:A2", "shift": "down" }] }),
    ] {
        let error = engine::edit(handle, &request).expect_err("invalid insert request");
        assert_eq!(error.code, if request["operations"][0]["sheet"] == json!("Nope") { "object/not-found" } else { "session/argument-invalid" });
    }
    assert_eq!(engine::summary(handle).expect("summary")["revision"], json!(1), "rejected batches never advance");
    engine::close(handle).expect("close");
}
