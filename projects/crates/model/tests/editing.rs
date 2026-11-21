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
    assert_eq!((columns[0]["min"].clone(), columns[0]["max"].clone(), columns[0]["width"].clone()), (json!(1), json!(1), json!(12.0)));
    assert_eq!((columns[1]["min"].clone(), columns[1]["max"].clone(), columns[1]["width"].clone()), (json!(2), json!(2), json!(20.0)));
    assert_eq!((columns[2]["min"].clone(), columns[2]["max"].clone(), columns[2]["width"].clone()), (json!(3), json!(3), json!(12.0)));

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
        columns
            .iter()
            .any(|record| record["min"] == json!(1)
                && record["max"] == json!(1)
                && (record["width"].as_f64().unwrap_or_default() - 24.0).abs() < 0.001),
        "column A keeps its width"
    );
    let rows = sheet["rows"].as_array().expect("rows");
    assert!(
        rows.iter().any(|record| record["index"] == json!(1)
            && (record["height"].as_f64().unwrap_or_default() - 30.0).abs() < 0.001),
        "row 1 keeps its height"
    );

    let _ = fs::remove_file(xlsx_path);
    let _ = fs::remove_file(project_path);
}
