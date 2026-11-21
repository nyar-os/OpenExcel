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
                { "op": "freeze.set", "sheet": "Sheet1", "row": 1, "column": 1 },
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
                { "op": "freeze.set", "sheet": "Sheet1", "row": 3, "column": 0 },
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
                { "op": "freeze.set", "sheet": "Sheet1", "row": 1, "column": 2 },
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
                { "op": "freeze.set", "sheet": "Sheet1", "row": 1, "column": 2 },
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
                { "op": "freeze.set", "sheet": "Sheet1", "row": 1, "column": 0 },
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
