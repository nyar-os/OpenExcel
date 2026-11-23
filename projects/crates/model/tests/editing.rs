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
