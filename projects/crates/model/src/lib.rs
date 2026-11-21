//! OpenExcel workbook model and engine surface.
//!
//! The crate owns the editable project model, the XLSX/CSV importers, the XLSX
//! writer, the inspection rules, the SVG grid renderer and the process-wide
//! handle store. Bindings (Node-API, WebView, CLI) only map arguments onto the
//! functions below.

#![deny(missing_docs)]

mod csv;
mod error;
mod eval;
mod inspect;
mod model;
mod reference;
mod render;
mod store;
mod xlsx;

use std::{
    fs,
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
};

use serde_json::{Map, Value, json};

pub use error::ServiceError;

use model::{CellValue, PROJECT_FORMAT, PROJECT_VERSION, Project, Sheet};

/// Workbook formats the product classifies.
pub const WORKBOOK_FORMATS: &[&str] = &["xlsx", "xlsm", "xltx", "xls", "csv", "tsv", "ods", "json", "xml", "pdf"];

/// Inspection rules exposed by [`inspect`].
pub const INSPECT_CHECKS: &[&str] = inspect::CHECKS;

/// Render targets the studio exposes.
pub const RENDER_FORMATS: &[&str] = &["svg", "png", "pdf"];

/// Export targets the SDK exposes.
pub const EXPORT_TARGETS: &[&str] = &["xlsx", "xlsm", "csv", "tsv", "json", "html", "pdf"];

static TEMP_COUNTER: AtomicU64 = AtomicU64::new(0);

/// What this build can actually do.
///
/// The matrix reports the wired subset instead of treating an extension or a
/// method name as a capability.
pub fn capabilities() -> Value {
    let mut formats = Map::new();
    for format in WORKBOOK_FORMATS {
        let entry = match *format {
            "xlsx" => format_entry("supported", "supported", "supported", "supported"),
            "xlsm" => format_entry("supported", "supported", "unsupported", "supported"),
            "csv" | "tsv" => format_entry("supported", "supported", "supported", "supported"),
            "json" => format_entry("unsupported", "unsupported", "supported", "unsupported"),
            _ => format_entry("unsupported", "unsupported", "unsupported", "unsupported"),
        };
        formats.insert((*format).to_owned(), entry);
    }

    let checks =
        INSPECT_CHECKS.iter().map(|check| ((*check).to_owned(), Value::String("supported".to_owned()))).collect::<Map<_, _>>();
    json!({
        "engine": { "package": env!("CARGO_PKG_NAME"), "version": env!("CARGO_PKG_VERSION") },
        "project": "supported",
        "formats": Value::Object(formats),
        "checks": Value::Object(checks),
        "renderers": {
            "svg": "supported",
            "png": "unsupported",
            "pdf": "unsupported",
        },
        "exportTargets": {
            "xlsx": "supported",
            "xlsm": "unsupported",
            "csv": "supported",
            "tsv": "supported",
            "json": "supported",
            "html": "supported",
            "pdf": "unsupported",
        },
    })
}

fn format_entry(import: &str, edit: &str, export: &str, inspect: &str) -> Value {
    json!({
        "import": import,
        "edit": edit,
        "export": export,
        "inspect": inspect,
        "lossCodes": [],
    })
}

/// Creates an empty workbook with one `Sheet1`.
pub fn create(request: &Value) -> Result<Value, ServiceError> {
    let title =
        request.get("title").and_then(Value::as_str).map(str::trim).filter(|title| !title.is_empty()).unwrap_or("Workbook");
    if let Some(template) = request.get("template").and_then(Value::as_str).filter(|template| !template.trim().is_empty())
        && !template.eq_ignore_ascii_case("blank")
    {
        return Err(ServiceError::new(
            "object/unsupported",
            format!("template `{template}` is not available; this build starts from `blank`"),
        ));
    }
    let mut project = Project::blank(title);
    eval::recalculate(&mut project);
    let handle = store::insert(project, None);
    Ok(json!({ "handle": handle, "workbook": store::with(handle, |entry| summary_json(handle, entry))? }))
}

/// Opens a workbook project or a supported spreadsheet file.
pub fn open(path: &Path) -> Result<Value, ServiceError> {
    let bytes = fs::read(path).map_err(|error| ServiceError::io(format!("failed to read {}: {error}", path.display())))?;
    let extension = path.extension().and_then(|extension| extension.to_str()).map(str::to_ascii_lowercase).unwrap_or_default();
    let title = path.file_stem().and_then(|stem| stem.to_str()).unwrap_or("Workbook").to_owned();
    let mut project = match extension.as_str() {
        "oexcel" => read_project_json(&bytes)?,
        "xlsx" | "xlsm" => xlsx::read_project(&path.display().to_string(), &bytes, title)?,
        "csv" | "tsv" => {
            let text =
                String::from_utf8(bytes).map_err(|_| ServiceError::invalid(&extension, "the file is not valid UTF-8"))?;
            csv::project_from_csv(title, &text, if extension == "tsv" { '\t' } else { ',' })
        }
        "" => return Err(ServiceError::unsupported("unknown")),
        other => return Err(ServiceError::unsupported(other)),
    };
    eval::recalculate(&mut project);
    let handle = store::insert(project, Some(path.to_path_buf()));
    Ok(json!({ "handle": handle, "workbook": store::with(handle, |entry| summary_json(handle, entry))? }))
}

fn read_project_json(bytes: &[u8]) -> Result<Project, ServiceError> {
    let mut project: Project = serde_json::from_slice(bytes)
        .map_err(|error| ServiceError::invalid("oexcel", format!("the project JSON cannot be parsed: {error}")))?;
    if project.format != PROJECT_FORMAT {
        return Err(ServiceError::invalid("oexcel", format!("unexpected format tag `{}`", project.format)));
    }
    if project.version != PROJECT_VERSION {
        return Err(ServiceError::invalid("oexcel", format!("unsupported project version {}", project.version)));
    }
    project.normalize();
    Ok(project)
}

/// Writes one open workbook to a project path through a temporary file.
pub fn save(handle: u64, path: &Path, overwrite: bool) -> Result<Value, ServiceError> {
    let (project, revision) = store::with(handle, |entry| (entry.project.clone(), entry.revision))?;
    let mut bytes = serde_json::to_vec_pretty(&project)
        .map_err(|error| ServiceError::new("format/invalid", format!("the project cannot be serialized: {error}")))?;
    bytes.push(b'\n');
    write_file(path, &bytes, overwrite)?;
    store::with_mut(handle, |entry| entry.path = Some(path.to_path_buf()))?;
    Ok(json!({
        "target": path.display().to_string(),
        "revision": revision,
        "sheets": project.sheets.len(),
        "bytes": bytes.len(),
    }))
}

/// Returns the summary of one open workbook.
pub fn summary(handle: u64) -> Result<Value, ServiceError> {
    store::with(handle, |entry| summary_json(handle, entry))
}

fn summary_json(handle: u64, entry: &store::Entry) -> Value {
    json!({
        "id": format!("workbook-{handle}"),
        "revision": entry.revision,
        "title": entry.project.title,
        "sheets": entry.project.sheets.iter().map(|sheet| json!({
            "id": sheet.id,
            "index": sheet.index,
            "name": sheet.name,
            "hidden": !sheet.visibility.is_visible(),
            "cells": sheet.non_empty_cells(),
        })).collect::<Vec<_>>(),
        "cells": entry.project.cell_count(),
    })
}

/// Runs the requested rule checks.
pub fn inspect(handle: u64, checks: Option<Vec<String>>) -> Result<Value, ServiceError> {
    store::with(handle, |entry| inspect::run(&entry.project, entry.revision, checks.as_deref()))?
}

/// Writes one open workbook to an export target.
pub fn export(handle: u64, format: &str, target: &Path, overwrite: bool) -> Result<Value, ServiceError> {
    let (project, revision) = store::with(handle, |entry| (entry.project.clone(), entry.revision))?;
    // Export always projects freshly evaluated values, never a stale cache.
    let mut project = project;
    eval::recalculate(&mut project);
    let format = format.trim().to_ascii_lowercase();
    let (bytes, losses, editable) = match format.as_str() {
        "csv" => {
            (csv::write(primary_sheet(&project)?, ',').into_bytes(), skipped_sheets(&project, "export.csv.sheet-skipped"), true)
        }
        "tsv" => (
            csv::write(primary_sheet(&project)?, '\t').into_bytes(),
            skipped_sheets(&project, "export.tsv.sheet-skipped"),
            true,
        ),
        "json" => (
            serde_json::to_vec_pretty(&json_projection(&project))
                .map_err(|error| ServiceError::invalid("json", error.to_string()))?,
            Vec::new(),
            false,
        ),
        "html" => (html_projection(&project).into_bytes(), skipped_sheets(&project, "export.html.sheet-skipped"), false),
        "xlsx" => {
            let bytes = xlsx::write_package(&project)?;
            verify_xlsx(&project, &bytes)?;
            (bytes, Vec::new(), true)
        }
        "xlsm" | "pdf" => return Err(ServiceError::unsupported(&format)),
        other => return Err(ServiceError::new("format/unsupported", format!("`{other}` is not an export target"))),
    };
    write_file(target, &bytes, overwrite)?;
    Ok(json!({
        "revision": revision,
        "format": format,
        "target": target.display().to_string(),
        "losses": losses,
        "editable": editable,
    }))
}

fn primary_sheet(project: &Project) -> Result<&Sheet, ServiceError> {
    project.primary_sheet().ok_or_else(|| ServiceError::new("object/not-found", "the workbook has no sheet"))
}

fn skipped_sheets(project: &Project, code: &str) -> Vec<Value> {
    let primary = project.primary_sheet().map(|sheet| sheet.id.as_str()).unwrap_or_default();
    project
        .sheets
        .iter()
        .filter(|sheet| sheet.id != primary)
        .map(|sheet| {
            json!({
                "code": code,
                "severity": "warning",
                "sheetId": sheet.id,
                "message": format!("sheet `{}` was not written; this target holds one sheet", sheet.name),
            })
        })
        .collect()
}

fn json_projection(project: &Project) -> Value {
    json!({
        "sheets": project.sheets.iter().map(|sheet| json!({
            "name": sheet.name,
            "cells": sheet.cells.iter().filter(|cell| !cell.is_empty()).map(|cell| json!({
                "address": cell.reference,
                "value": primitive(&eval::value_of(cell)),
                "formula": cell.formula,
            })).collect::<Vec<_>>(),
        })).collect::<Vec<_>>(),
    })
}

/// Reads one cell of an open workbook with its evaluated value.
pub fn cell(handle: u64, request: &Value) -> Result<Value, ServiceError> {
    store::with(handle, |entry| -> Result<Value, ServiceError> {
        let key = request.get("sheet").and_then(Value::as_str).unwrap_or("");
        let sheet = if key.is_empty() { entry.project.primary_sheet() } else { entry.project.sheet(key) }
            .ok_or_else(|| missing_sheet(key))?;
        let address = request.get("address").and_then(Value::as_str).unwrap_or("").trim();
        let (row, column) = reference::parse_reference(address)
            .ok_or_else(|| ServiceError::argument(format!("`{address}` is not an A1 cell reference")))?;
        Ok(match sheet.cell(row, column) {
            Some(cell) => json!({
                "sheet": sheet.name,
                "address": cell.reference,
                "kind": cell_kind(cell),
                "formula": cell.formula,
                "value": primitive(&eval::value_of(cell)),
                "cached": cell.cached.as_ref().map(primitive).unwrap_or(Value::Null),
            }),
            None => json!({
                "sheet": sheet.name,
                "address": reference::format_reference(row, column),
                "kind": "empty",
                "formula": Value::Null,
                "value": Value::Null,
                "cached": Value::Null,
            }),
        })
    })?
}

fn cell_kind(cell: &model::Cell) -> &'static str {
    if cell.formula.is_some() {
        return "formula";
    }
    match cell.value {
        CellValue::Text { .. } => "text",
        CellValue::Number { .. } => "number",
        CellValue::Bool { .. } => "boolean",
        CellValue::Error { .. } => "error",
        CellValue::Empty => "empty",
    }
}

/// Reads a rectangle of cells with their evaluated values, row by row.
pub fn range(handle: u64, request: &Value) -> Result<Value, ServiceError> {
    store::with(handle, |entry| -> Result<Value, ServiceError> {
        let key = request.get("sheet").and_then(Value::as_str).unwrap_or("");
        let sheet = if key.is_empty() { entry.project.primary_sheet() } else { entry.project.sheet(key) }
            .ok_or_else(|| missing_sheet(key))?;
        let text = request.get("range").and_then(Value::as_str).unwrap_or("");
        let range = reference::parse_range(text)
            .ok_or_else(|| ServiceError::argument(format!("`{text}` is not an A1 range such as `A1:C4`")))?;
        if u64::from(range.rows()) * u64::from(range.columns()) > 100_000 {
            return Err(ServiceError::argument("the requested range is too large to return in one response"));
        }
        let rows = (range.start_row..=range.end_row)
            .map(|row| {
                (range.start_column..=range.end_column)
                    .map(|column| {
                        let address = reference::format_reference(row, column);
                        match sheet.cell(row, column) {
                            Some(cell) => json!({
                                "address": address,
                                "value": primitive(&eval::value_of(cell)),
                                "formula": cell.formula,
                            }),
                            None => json!({ "address": address, "value": Value::Null, "formula": Value::Null }),
                        }
                    })
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>();
        Ok(json!({ "sheet": sheet.name, "range": range.reference(), "rows": rows }))
    })?
}

/// Recomputes every formula of one open workbook.
pub fn recalculate(handle: u64) -> Result<Value, ServiceError> {
    store::with_mut(handle, |entry| {
        let changed = eval::recalculate(&mut entry.project);
        if changed > 0 {
            entry.revision += 1;
        }
        json!({ "revision": entry.revision, "changed": changed })
    })
}

fn primitive(value: &CellValue) -> Value {
    match value {
        CellValue::Text { text } => Value::String(text.clone()),
        CellValue::Number { number } => number_json(*number),
        CellValue::Bool { value } => Value::Bool(*value),
        CellValue::Error { code } => Value::String(code.clone()),
        CellValue::Empty => Value::Null,
    }
}

/// Projects a number so whole values lose their `.0`, matching CSV and XLSX.
fn number_json(number: f64) -> Value {
    if number.fract() == 0.0 && number.abs() <= 9_007_199_254_740_992.0 {
        Value::Number(serde_json::Number::from(number as i64))
    }
    else {
        serde_json::Number::from_f64(number).map(Value::Number).unwrap_or(Value::Null)
    }
}

fn html_projection(project: &Project) -> String {
    let sheet = project.primary_sheet();
    let mut html = String::new();
    html.push_str("<!doctype html><html><head><meta charset=\"utf-8\"><title>");
    html.push_str(&escape_html(&project.title));
    html.push_str("</title><style>body{font-family:Segoe UI,Arial,sans-serif}table{border-collapse:collapse}td{border:1px solid #d0d5dd;padding:2px 6px;font-size:13px;vertical-align:top}</style></head><body><table>");
    if let Some(sheet) = sheet
        && let Some(bounds) = sheet.actual_range()
    {
        html.push_str("<colgroup>");
        for column in 1..=bounds.end_column {
            html.push_str(&format!(
                "<col style=\"width:{:.0}px\">",
                (sheet.column_width(column) * 7.0 + 5.0).clamp(24.0, 400.0)
            ));
        }
        html.push_str("</colgroup>");
        for row in 1..=bounds.end_row {
            html.push_str("<tr>");
            for column in 1..=bounds.end_column {
                if let Some(covering) = sheet.merge_covering(row, column)
                    && (covering.start_row != row || covering.start_column != column)
                {
                    continue;
                }
                let (row_span, column_span) = match sheet.merge_covering(row, column) {
                    Some(covering) => (covering.rows(), covering.columns()),
                    None => (1, 1),
                };
                let cell = sheet.cell(row, column);
                let style = cell.and_then(|cell| cell.style.as_ref());
                html.push_str("<td");
                if row_span > 1 {
                    html.push_str(&format!(" rowspan=\"{row_span}\""));
                }
                if column_span > 1 {
                    html.push_str(&format!(" colspan=\"{column_span}\""));
                }
                if let Some(style) = style {
                    let mut declarations: Vec<String> = Vec::new();
                    if style.bold.unwrap_or(false) {
                        declarations.push("font-weight:bold".to_owned());
                    }
                    if style.italic.unwrap_or(false) {
                        declarations.push("font-style:italic".to_owned());
                    }
                    if let Some(fill) = style.fill.as_deref().and_then(css_color) {
                        declarations.push(format!("background:{fill}"));
                    }
                    if !declarations.is_empty() {
                        html.push_str(&format!(" style=\"{}\"", declarations.join(";")));
                    }
                }
                html.push('>');
                if let Some(cell) = cell {
                    html.push_str(&escape_html(&eval::value_of(cell).display_text()).replace('\n', "<br>"));
                }
                html.push_str("</td>");
            }
            html.push_str("</tr>");
        }
    }
    html.push_str("</table></body></html>");
    html
}

fn escape_html(text: &str) -> String {
    text.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;")
}

fn css_color(value: &str) -> Option<String> {
    let hex = value.trim().trim_start_matches('#');
    match hex.len() {
        6 if hex.chars().all(|character| character.is_ascii_hexdigit()) => Some(format!("#{hex}")),
        8 if hex.chars().all(|character| character.is_ascii_hexdigit()) => Some(format!("#{}", &hex[2..])),
        _ => None,
    }
}

/// Compares a written package with the project it came from.
fn verify_xlsx(project: &Project, bytes: &[u8]) -> Result<(), ServiceError> {
    let workbook = document_formats::import::xlsx::import_xlsx_bytes("verification", bytes)
        .map_err(|error| ServiceError::io(format!("the written XLSX package cannot be read back: {error}")))?;
    let readback = xlsx::project_from_workbook(workbook, project.title.clone());
    if readback.sheets.len() != project.sheets.len() {
        return Err(ServiceError::io("the written XLSX package lost a sheet"));
    }
    for (written, read) in project.sheets.iter().zip(&readback.sheets) {
        if written.cells.len() != read.cells.len() {
            return Err(ServiceError::io(format!(
                "the written XLSX package lost cells on sheet `{}`: expected {}, read {}",
                written.name,
                written.cells.len(),
                read.cells.len()
            )));
        }
        for (expected, actual) in written.cells.iter().zip(&read.cells) {
            // Formula cells carry their result in `computed`/`cached`, so compare
            // what a reader sees rather than the optional literal fallback.
            if expected.reference != actual.reference
                || eval::value_of(expected) != eval::value_of(actual)
                || expected.formula != actual.formula
            {
                return Err(ServiceError::io(format!(
                    "the written XLSX package changed cell `{}` on sheet `{}`",
                    expected.reference, written.name
                )));
            }
        }
    }
    Ok(())
}

/// Applies edits to one open workbook.
pub fn edit(handle: u64, request: &Value) -> Result<Value, ServiceError> {
    store::with_mut(handle, |entry| -> Result<Value, ServiceError> {
        if let Some(expected) = request.get("expectedRevision").and_then(Value::as_u64)
            && expected != entry.revision
        {
            return Err(ServiceError::new(
                "document/revision-conflict",
                format!("expected revision {expected} but the workbook is at revision {}", entry.revision),
            ));
        }
        let operations = request
            .get("operations")
            .and_then(Value::as_array)
            .ok_or_else(|| ServiceError::argument("an edit request needs an `operations` array"))?;
        let mut project = entry.project.clone();
        for operation in operations {
            apply_operation(&mut project, operation)?;
        }
        project.normalize();
        eval::recalculate(&mut project);
        entry.project = project;
        entry.revision += 1;
        Ok(json!({ "revision": entry.revision, "workbook": summary_json(handle, entry) }))
    })?
}

fn apply_operation(project: &mut Project, operation: &Value) -> Result<(), ServiceError> {
    let kind = operation
        .get("op")
        .and_then(Value::as_str)
        .ok_or_else(|| ServiceError::argument("each operation needs an `op` field"))?;
    match kind {
        "cell.set" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
            let reference = operation.get("reference").and_then(Value::as_str).unwrap_or("");
            let (row, column) = reference::parse_reference(reference)
                .ok_or_else(|| ServiceError::argument(format!("`{reference}` is not an A1 cell reference")))?;
            let formula = operation
                .get("formula")
                .and_then(Value::as_str)
                .map(str::trim)
                .filter(|formula| !formula.is_empty())
                .map(|formula| if formula.starts_with('=') { formula.to_owned() } else { format!("={formula}") });
            let value = match operation.get("value") {
                Some(value) if !value.is_null() => value_from_json(value)?,
                _ => CellValue::Empty,
            };
            if formula.is_none() && value.is_empty() {
                return Err(ServiceError::argument("`cell.set` needs a `value` or a `formula`"));
            }
            sheet.write_cell(row, column, value, formula, None);
        }
        "sheet.add" => {
            let name = required_name(operation)?;
            if project.sheets.iter().any(|sheet| sheet.name.eq_ignore_ascii_case(&name)) {
                return Err(ServiceError::argument(format!("sheet `{name}` already exists")));
            }
            let id = project.allocate_sheet_id();
            let index =
                operation.get("index").and_then(Value::as_u64).map(|index| index as usize).unwrap_or(project.sheets.len());
            let index = index.min(project.sheets.len());
            let mut sheet = Sheet::new(id, name, index as u32);
            sheet.index = index as u32;
            project.sheets.insert(index, sheet);
            project.reindex();
        }
        "sheet.rename" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let name = required_name(operation)?;
            let index = project.sheet_index(key).ok_or_else(|| missing_sheet(key))?;
            if project.sheets.iter().enumerate().any(|(other, sheet)| other != index && sheet.name.eq_ignore_ascii_case(&name))
            {
                return Err(ServiceError::argument(format!("sheet `{name}` already exists")));
            }
            project.sheets[index].name = name;
        }
        "sheet.remove" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let index = project.sheet_index(key).ok_or_else(|| missing_sheet(key))?;
            if project.sheets.len() == 1 {
                return Err(ServiceError::argument("a workbook must keep at least one sheet"));
            }
            project.sheets.remove(index);
            project.reindex();
        }
        "range.clear" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
            let range = operation
                .get("range")
                .and_then(Value::as_str)
                .and_then(reference::parse_range)
                .ok_or_else(|| ServiceError::argument("`range.clear` needs an A1 range such as `A1:C4`"))?;
            sheet.clear_range(&range);
        }
        "range.set" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
            let start = operation
                .get("start")
                .and_then(Value::as_str)
                .and_then(reference::parse_reference)
                .ok_or_else(|| ServiceError::argument("`range.set` needs an A1 `start` such as `B2`"))?;
            let rows = operation
                .get("rows")
                .and_then(Value::as_array)
                .ok_or_else(|| ServiceError::argument("`range.set` needs a `rows` matrix"))?;
            for (row_offset, row) in rows.iter().enumerate() {
                let cells = row
                    .as_array()
                    .ok_or_else(|| ServiceError::argument("each `range.set` row must be an array of cell inputs"))?;
                for (column_offset, input) in cells.iter().enumerate() {
                    let row = start.0.checked_add(row_offset as u32).filter(|row| *row <= 1_048_576);
                    let column = start.1.checked_add(column_offset as u32).filter(|column| *column <= 16_384);
                    let (Some(row), Some(column)) = (row, column)
                    else {
                        return Err(ServiceError::argument("the `range.set` matrix leaves the sheet grid"));
                    };
                    let (value, formula) = range_cell_input(input)?;
                    sheet.write_cell(row, column, value, formula, None);
                }
            }
        }
        "format.set" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
            let range = operation
                .get("range")
                .and_then(Value::as_str)
                .and_then(reference::parse_range)
                .ok_or_else(|| ServiceError::argument("`format.set` needs an A1 range such as `A1:C4`"))?;
            let style = style_from_json(
                operation.get("style").ok_or_else(|| ServiceError::argument("`format.set` needs a `style` object"))?,
            )?;
            for row in range.start_row..=range.end_row {
                for column in range.start_column..=range.end_column {
                    let mut cell = sheet.cell(row, column).cloned().unwrap_or_else(|| model::Cell {
                        reference: reference::format_reference(row, column),
                        row,
                        column,
                        value: CellValue::Empty,
                        formula: None,
                        cached: None,
                        style: None,
                        computed: None,
                    });
                    cell.style = Some(merge_style(cell.style.as_ref(), &style));
                    sheet.set_cell(cell);
                }
            }
        }
        "sheet.duplicate" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let index = project.sheet_index(key).ok_or_else(|| missing_sheet(key))?;
            let mut copy = project.sheets[index].clone();
            copy.id = project.allocate_sheet_id();
            let requested = operation.get("name").and_then(Value::as_str).map(str::trim).filter(|name| !name.is_empty());
            copy.name = match requested {
                Some(name) => {
                    if project.sheets.iter().any(|sheet| sheet.name.eq_ignore_ascii_case(name)) {
                        return Err(ServiceError::argument(format!("sheet `{name}` already exists")));
                    }
                    name.to_owned()
                }
                None => unique_sheet_name(project, &format!("{} Copy", copy.name)),
            };
            let insert = index + 1;
            project.sheets.insert(insert, copy);
            project.reindex();
        }
        "sheet.move" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let index = project.sheet_index(key).ok_or_else(|| missing_sheet(key))?;
            let target = operation
                .get("index")
                .and_then(Value::as_u64)
                .ok_or_else(|| ServiceError::argument("`sheet.move` needs a zero-based `index`"))?
                as usize;
            let sheet = project.sheets.remove(index);
            let target = target.min(project.sheets.len());
            project.sheets.insert(target, sheet);
            project.reindex();
        }
        other => return Err(ServiceError::new("object/unsupported", format!("edit operation `{other}` is not supported"))),
    }
    Ok(())
}

/// Reads one `range.set` entry: a primitive literal or `{ formula, value? }`.
fn range_cell_input(value: &Value) -> Result<(CellValue, Option<String>), ServiceError> {
    match value {
        Value::Null => Ok((CellValue::Empty, None)),
        Value::String(text) => Ok((CellValue::text(text), None)),
        Value::Number(number) => Ok((CellValue::number(number.as_f64().unwrap_or(0.0)), None)),
        Value::Bool(value) => Ok((CellValue::boolean(*value), None)),
        Value::Object(map) => {
            let formula = map
                .get("formula")
                .and_then(Value::as_str)
                .map(str::trim)
                .filter(|formula| !formula.is_empty())
                .map(|formula| if formula.starts_with('=') { formula.to_owned() } else { format!("={formula}") });
            let value = match map.get("value") {
                Some(Value::Null) | None => CellValue::Empty,
                Some(raw) => value_from_json(raw)?,
            };
            if formula.is_none() && value.is_empty() {
                return Err(ServiceError::argument("a `range.set` cell needs a `formula` or a `value`"));
            }
            Ok((value, formula))
        }
        _ => Err(ServiceError::argument(
            "a `range.set` cell must be null, a number, a boolean, a string or an object with `formula` and `value`",
        )),
    }
}

fn style_from_json(value: &Value) -> Result<model::CellStyle, ServiceError> {
    let object = value.as_object().ok_or_else(|| ServiceError::argument("`format.set` needs a `style` object"))?;
    let style = model::CellStyle {
        bold: optional_bool(object, "bold")?,
        italic: optional_bool(object, "italic")?,
        font_size: optional_f32(object, "fontSize")?,
        font_color: optional_string(object, "fontColor")?,
        fill: optional_string(object, "fill")?,
        number_format: optional_string(object, "numberFormat")?,
        align: optional_string(object, "align")?,
        wrap: optional_bool(object, "wrap")?,
    };
    if style.is_empty() {
        return Err(ServiceError::argument("`format.set` needs at least one style property"));
    }
    Ok(style)
}

fn optional_bool(object: &Map<String, Value>, key: &str) -> Result<Option<bool>, ServiceError> {
    match object.get(key) {
        None | Some(Value::Null) => Ok(None),
        Some(Value::Bool(value)) => Ok(Some(*value)),
        Some(_) => Err(ServiceError::argument(format!("`{key}` must be a boolean"))),
    }
}

fn optional_f32(object: &Map<String, Value>, key: &str) -> Result<Option<f32>, ServiceError> {
    match object.get(key) {
        None | Some(Value::Null) => Ok(None),
        Some(Value::Number(number)) => number
            .as_f64()
            .filter(|value| *value > 0.0)
            .map(|value| Some(value as f32))
            .ok_or_else(|| ServiceError::argument(format!("`{key}` must be a positive number"))),
        Some(_) => Err(ServiceError::argument(format!("`{key}` must be a positive number"))),
    }
}

fn optional_string(object: &Map<String, Value>, key: &str) -> Result<Option<String>, ServiceError> {
    match object.get(key) {
        None | Some(Value::Null) => Ok(None),
        Some(Value::String(value)) if !value.trim().is_empty() => Ok(Some(value.clone())),
        Some(_) => Err(ServiceError::argument(format!("`{key}` must be a non-empty string"))),
    }
}

/// Merges one incoming style over the style already on a cell.
fn merge_style(existing: Option<&model::CellStyle>, incoming: &model::CellStyle) -> model::CellStyle {
    let base = existing.cloned().unwrap_or_default();
    model::CellStyle {
        bold: incoming.bold.or(base.bold),
        italic: incoming.italic.or(base.italic),
        font_size: incoming.font_size.or(base.font_size),
        font_color: incoming.font_color.clone().or(base.font_color),
        fill: incoming.fill.clone().or(base.fill),
        number_format: incoming.number_format.clone().or(base.number_format),
        align: incoming.align.clone().or(base.align),
        wrap: incoming.wrap.or(base.wrap),
    }
}

/// Allocates a sheet name that no other sheet uses.
fn unique_sheet_name(project: &Project, base: &str) -> String {
    if project.sheets.iter().all(|sheet| !sheet.name.eq_ignore_ascii_case(base)) {
        return base.to_owned();
    }
    let mut suffix = 2;
    loop {
        let candidate = format!("{base} {suffix}");
        if project.sheets.iter().all(|sheet| !sheet.name.eq_ignore_ascii_case(&candidate)) {
            return candidate;
        }
        suffix += 1;
    }
}

fn missing_sheet(key: &str) -> ServiceError {
    if key.is_empty() {
        ServiceError::argument("an operation needs a `sheet` name or id")
    }
    else {
        ServiceError::new("object/not-found", format!("sheet `{key}` does not exist"))
    }
}

fn required_name(operation: &Value) -> Result<String, ServiceError> {
    operation
        .get("name")
        .and_then(Value::as_str)
        .map(str::trim)
        .filter(|name| !name.is_empty())
        .map(str::to_owned)
        .ok_or_else(|| ServiceError::argument("an operation needs a non-empty sheet `name`"))
}

fn value_from_json(value: &Value) -> Result<CellValue, ServiceError> {
    if let Value::Object(map) = value
        && map.contains_key("kind")
    {
        return serde_json::from_value(value.clone())
            .map_err(|error| ServiceError::argument(format!("invalid cell value: {error}")));
    }
    Ok(match value {
        Value::String(text) => CellValue::text(text),
        Value::Number(number) => CellValue::number(number.as_f64().unwrap_or(0.0)),
        Value::Bool(value) => CellValue::boolean(*value),
        Value::Null => CellValue::Empty,
        _ => return Err(ServiceError::argument("a cell value must be a string, number, boolean, null or a tagged value")),
    })
}

/// Renders one sheet of an open workbook as SVG.
pub fn render(handle: u64, request: &Value) -> Result<Value, ServiceError> {
    let format = request.get("format").and_then(Value::as_str).unwrap_or("svg").to_ascii_lowercase();
    if format != "svg" {
        return Err(ServiceError::unsupported(&format));
    }
    store::with(handle, |entry| {
        let key = request.get("sheet").and_then(Value::as_str).unwrap_or("");
        let sheet = if key.is_empty() { entry.project.primary_sheet() } else { entry.project.sheet(key) }
            .ok_or_else(|| missing_sheet(key))?;
        Ok(json!({
            "revision": entry.revision,
            "format": "svg",
            "sheet": sheet.name,
            "sheetId": sheet.id,
            "svg": render::sheet_svg(sheet),
        }))
    })?
}

/// Releases one open workbook handle. Closing an unknown handle succeeds.
pub fn close(handle: u64) -> Result<(), ServiceError> {
    store::remove(handle);
    Ok(())
}

fn write_file(path: &Path, bytes: &[u8], overwrite: bool) -> Result<(), ServiceError> {
    if path.exists() && !overwrite {
        return Err(ServiceError::new("io/target-exists", format!("{} already exists; pass `--overwrite`", path.display())));
    }
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from("."));
    let name = path.file_name().and_then(|name| name.to_str()).unwrap_or("workbook");
    let counter = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
    let temporary = parent.join(format!(".{name}.{}.{counter}.tmp", std::process::id()));
    fs::write(&temporary, bytes)
        .map_err(|error| ServiceError::io(format!("failed to write {}: {error}", temporary.display())))?;
    if path.exists()
        && let Err(error) = fs::remove_file(path)
    {
        let _ = fs::remove_file(&temporary);
        return Err(ServiceError::io(format!("failed to replace {}: {error}", path.display())));
    }
    if let Err(error) = fs::rename(&temporary, path) {
        let _ = fs::remove_file(&temporary);
        return Err(ServiceError::io(format!("failed to move the temporary file into {}: {error}", path.display())));
    }
    Ok(())
}

#[cfg(test)]
mod tests;
