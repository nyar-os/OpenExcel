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
#[doc(hidden)]
pub mod model;
mod pdf;
mod reference;
mod render;
mod shift;
#[doc(hidden)]
pub mod store;
mod structure;
#[doc(hidden)]
pub mod xlsx;

use std::{
    fs,
    path::{Path, PathBuf},
};

use boost_fs::{OverwritePolicy, PublishOptions, publish_bytes_with_options};
use serde_json::{Map, Value, json};

pub use error::ServiceError;

use model::{CellRange, CellValue, PROJECT_FORMAT, PROJECT_VERSION, Project, Sheet, SheetColumn, SheetRow};

/// Workbook formats the product classifies.
pub const WORKBOOK_FORMATS: &[&str] = &["xlsx", "xlsm", "xltx", "xls", "csv", "tsv", "ods", "json", "xml", "pdf"];

/// Inspection rules exposed by [`inspect`].
pub const INSPECT_CHECKS: &[&str] = inspect::CHECKS;

/// Render targets the studio exposes.
pub const RENDER_FORMATS: &[&str] = &["svg", "png", "pdf"];

/// Export targets the SDK exposes.
pub const EXPORT_TARGETS: &[&str] = &["xlsx", "xlsm", "csv", "tsv", "json", "html", "pdf"];

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
        "formulas": {
            "functions": eval::FUNCTIONS,
            "unsupportedBehavior": "preserve-cached",
        },
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
            "pdf": "partial",
        },
        // `partial` targets list what their projection approximates or drops.
        "exportLimits": {
            "pdf": ["style-approximated", "no-images"],
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
        "pdf" => (
            pdf::write(&project),
            vec![
                json!({
                    "code": "export.pdf.style-approximated",
                    "severity": "warning",
                    "message": "the PDF projection approximates column widths, fonts and fills",
                }),
                json!({
                    "code": "export.pdf.no-images",
                    "severity": "warning",
                    "message": "the project model carries no images, so the PDF has none",
                }),
            ],
            false,
        ),
        "xlsx" => {
            let bytes = xlsx::write_package(&project)?;
            verify_xlsx(&project, &bytes)?;
            (bytes, Vec::new(), true)
        }
        "xlsm" => return Err(ServiceError::unsupported(&format)),
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
                "computedState": cell.computed_state,
            }),
            None => json!({
                "sheet": sheet.name,
                "address": reference::format_reference(row, column),
                "kind": "empty",
                "formula": Value::Null,
                "value": Value::Null,
                "cached": Value::Null,
                "computedState": Value::Null,
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

/// Searches evaluated values and formula sources of one open workbook.
///
/// Returns at most [`MAX_FIND_HITS`] hits with character offsets into the
/// matched field, plus `truncated` when the cap was reached.
pub fn find(handle: u64, request: &Value) -> Result<Value, ServiceError> {
    store::with(handle, |entry| -> Result<Value, ServiceError> {
        let query = request
            .get("query")
            .and_then(Value::as_str)
            .filter(|query| !query.is_empty())
            .ok_or_else(|| ServiceError::argument("`find` needs a non-empty `query`"))?;
        let options = request.get("options");
        let match_case = options.and_then(|options| options.get("matchCase")).and_then(Value::as_bool).unwrap_or(false);
        let search_formulas = options.and_then(|options| options.get("formulas")).and_then(Value::as_bool).unwrap_or(true);
        let search_values = options.and_then(|options| options.get("values")).and_then(Value::as_bool).unwrap_or(true);
        let key = request.get("sheet").and_then(Value::as_str).unwrap_or("");
        let sheets: Vec<usize> = if key.is_empty() {
            (0..entry.project.sheets.len()).collect()
        }
        else {
            vec![entry.project.sheet_index(key).ok_or_else(|| missing_sheet(key))?]
        };

        let mut hits: Vec<Value> = Vec::new();
        'sheets: for sheet_index in sheets {
            let sheet = &entry.project.sheets[sheet_index];
            for cell in &sheet.cells {
                if search_values {
                    let text = eval::value_of(cell).display_text();
                    for (start, end) in occurrence_spans(&text, query, match_case) {
                        if hits.len() >= MAX_FIND_HITS {
                            break 'sheets;
                        }
                        hits.push(json!({
                            "sheet": sheet.name,
                            "address": cell.reference,
                            "field": "value",
                            "start": start,
                            "end": end,
                            "text": slice_chars(&text, start, end),
                        }));
                    }
                }
                if search_formulas && let Some(formula) = &cell.formula {
                    for (start, end) in occurrence_spans(formula, query, match_case) {
                        if hits.len() >= MAX_FIND_HITS {
                            break 'sheets;
                        }
                        hits.push(json!({
                            "sheet": sheet.name,
                            "address": cell.reference,
                            "field": "formula",
                            "start": start,
                            "end": end,
                            "text": slice_chars(formula, start, end),
                        }));
                    }
                }
            }
        }
        Ok(json!({ "hits": hits, "truncated": hits.len() >= MAX_FIND_HITS }))
    })?
}

/// Upper bound on the hits one `find` request returns.
const MAX_FIND_HITS: usize = 500;

fn occurrence_spans(text: &str, query: &str, match_case: bool) -> Vec<(usize, usize)> {
    let haystack: Vec<char> = text.chars().collect();
    let needle: Vec<char> = query.chars().collect();
    let mut spans = Vec::new();
    if needle.is_empty() {
        return spans;
    }
    let mut index = 0usize;
    while index + needle.len() <= haystack.len() {
        if haystack[index..index + needle.len()]
            .iter()
            .zip(&needle)
            .all(|(left, right)| char_matches(*left, *right, match_case))
        {
            spans.push((index, index + needle.len()));
            index += needle.len();
        }
        else {
            index += 1;
        }
    }
    spans
}

fn slice_chars(text: &str, start: usize, end: usize) -> String {
    text.chars().skip(start).take(end.saturating_sub(start)).collect()
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
            // what a reader sees rather than the optional literal fallback. An
            // unsupported formula writes only the imported cache.
            let expected_value = if expected.computed_state == Some(model::ComputedState::Unsupported) {
                expected.cached.clone().unwrap_or(CellValue::Empty)
            }
            else {
                eval::value_of(expected)
            };
            if expected.reference != actual.reference
                || expected_value != eval::value_of(actual)
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
        let mut replacements = 0usize;
        let mut affected: Vec<Value> = Vec::new();
        for operation in operations {
            replacements += apply_operation(&mut project, operation, &mut affected)?;
        }
        project.normalize();
        eval::recalculate(&mut project);
        entry.project = project;
        entry.revision += 1;
        Ok(json!({
            "revision": entry.revision,
            "workbook": summary_json(handle, entry),
            "replacements": replacements,
            "affected": affected,
        }))
    })?
}

fn apply_operation(project: &mut Project, operation: &Value, affected: &mut Vec<Value>) -> Result<usize, ServiceError> {
    let kind = operation
        .get("op")
        .and_then(Value::as_str)
        .ok_or_else(|| ServiceError::argument("each operation needs an `op` field"))?;
    let replacements = match kind {
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
            0
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
            0
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
            0
        }
        "sheet.remove" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let index = project.sheet_index(key).ok_or_else(|| missing_sheet(key))?;
            if project.sheets.len() == 1 {
                return Err(ServiceError::argument("a workbook must keep at least one sheet"));
            }
            project.sheets.remove(index);
            project.reindex();
            0
        }
        "sheet.setVisibility" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let index = project.sheet_index(key).ok_or_else(|| missing_sheet(key))?;
            let requested = operation
                .get("visibility")
                .and_then(Value::as_str)
                .map(str::trim)
                .filter(|visibility| !visibility.is_empty())
                .ok_or_else(|| ServiceError::argument("`sheet.setVisibility` needs `visible`, `hidden` or `veryHidden`"))?;
            let visibility = match requested.to_ascii_lowercase().as_str() {
                "visible" => model::SheetVisibility::Visible,
                "hidden" => model::SheetVisibility::Hidden,
                "veryhidden" => model::SheetVisibility::VeryHidden,
                _ => {
                    return Err(ServiceError::argument(format!(
                        "`{requested}` is not a sheet visibility; expected `visible`, `hidden` or `veryHidden`"
                    )));
                }
            };
            if !visibility.is_visible()
                && project.sheets.iter().enumerate().all(|(other, sheet)| other == index || !sheet.visibility.is_visible())
            {
                return Err(ServiceError::argument("a workbook must keep at least one visible sheet"));
            }
            project.sheets[index].visibility = visibility;
            0
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
            0
        }
        "range.sort" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let range = operation
                .get("range")
                .and_then(Value::as_str)
                .and_then(reference::parse_range)
                .ok_or_else(|| ServiceError::argument("`range.sort` needs an A1 range such as `A1:C4`"))?;
            if !range.is_valid() || range.end_row > 1_048_576 || range.end_column > 16_384 {
                return Err(ServiceError::argument(format!("sort range `{}` leaves the sheet grid", range.reference())));
            }
            let column = sort_column(operation, &range)?;
            let ascending = operation.get("ascending").and_then(Value::as_bool).unwrap_or(true);
            let has_header = operation.get("hasHeader").and_then(Value::as_bool).unwrap_or(false);
            // Formula keys sort by the value a reader sees, not a stale cache.
            eval::recalculate(project);
            let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
            if let Some(merged) = sheet.merged.iter().find(|merged| merged.overlaps(&range)) {
                return Err(ServiceError::argument(format!(
                    "`range.sort` cannot move rows through merged range `{}`",
                    merged.reference()
                )));
            }
            let first_row = range.start_row + u32::from(has_header);
            if first_row <= range.end_row {
                let mut records = Vec::new();
                for row in first_row..=range.end_row {
                    let sort_value = sheet.cell(row, column).map(eval::value_of).unwrap_or(CellValue::Empty);
                    let cells = (range.start_column..=range.end_column)
                        .map(|column| sheet.cell(row, column).cloned())
                        .collect::<Vec<_>>();
                    records.push((sort_value, cells));
                }
                records.sort_by(|left, right| compare_sort_values(&left.0, &right.0, ascending));
                let body = CellRange {
                    start_row: first_row,
                    end_row: range.end_row,
                    start_column: range.start_column,
                    end_column: range.end_column,
                };
                sheet.clear_range(&body);
                for (offset, (_, cells)) in records.into_iter().enumerate() {
                    let row = first_row + offset as u32;
                    for (column_offset, cell) in cells.into_iter().enumerate() {
                        let Some(mut cell) = cell
                        else {
                            continue;
                        };
                        cell.row = row;
                        cell.column = range.start_column + column_offset as u32;
                        cell.reference = reference::format_reference(cell.row, cell.column);
                        sheet.set_cell(cell);
                    }
                }
            }
            affected.push(range_effect(project, operation, "range.sort", &range));
            0
        }
        "range.fillSeries" => {
            let filled = fill_series(project, operation)?;
            affected.push(range_effect(project, operation, "range.fillSeries", &filled));
            0
        }
        "range.transpose" => {
            let written = transpose(project, operation)?;
            affected.push(range_effect(project, operation, "range.transpose", &written));
            0
        }
        "range.copy" => {
            let written = copy_range(project, operation)?;
            affected.push(range_effect(project, operation, "range.copy", &written));
            0
        }
        "range.move" => {
            let written = move_range(project, operation)?;
            affected.push(range_effect(project, operation, "range.move", &written));
            0
        }
        "range.swap" => {
            let written = swap_range(project, operation)?;
            affected.push(range_effect(project, operation, "range.swap", &written));
            0
        }
        "range.flip" => {
            let written = flip_range(project, operation)?;
            affected.push(range_effect(project, operation, "range.flip", &written));
            0
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
            0
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
                        computed_state: None,
                    });
                    cell.style = Some(merge_style(cell.style.as_ref(), &style));
                    sheet.set_cell(cell);
                }
            }
            0
        }
        "format.copy" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
            let source = operation
                .get("source")
                .and_then(Value::as_str)
                .and_then(reference::parse_range)
                .filter(|range| range.start_row == range.end_row && range.start_column == range.end_column)
                .ok_or_else(|| {
                    ServiceError::argument("`format.copy` needs a one-cell `source` such as `B2`")
                })?;
            let target = operation
                .get("target")
                .and_then(Value::as_str)
                .and_then(reference::parse_range)
                .ok_or_else(|| ServiceError::argument("`format.copy` needs an A1 `target` range such as `B2:D8`"))?;
            let style = sheet
                .cell(source.start_row, source.start_column)
                .ok_or_else(|| {
                    ServiceError::argument(format!(
                        "`format.copy` found no cell at `{}` to copy formatting from",
                        source.reference()
                    ))
                })?
                .style
                .clone()
                .unwrap_or_default();
            for row in target.start_row..=target.end_row {
                for column in target.start_column..=target.end_column {
                    let mut cell = sheet.cell(row, column).cloned().unwrap_or_else(|| model::Cell {
                        reference: reference::format_reference(row, column),
                        row,
                        column,
                        value: CellValue::Empty,
                        formula: None,
                        cached: None,
                        style: None,
                        computed: None,
                        computed_state: None,
                    });
                    cell.style = if style.is_empty() { None } else { Some(style.clone()) };
                    sheet.set_cell(cell);
                }
            }
            0
        }
        "format.clear" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
            let range = operation
                .get("range")
                .and_then(Value::as_str)
                .and_then(reference::parse_range)
                .ok_or_else(|| ServiceError::argument("`format.clear` needs an A1 range such as `A1:C4`"))?;
            for row in range.start_row..=range.end_row {
                for column in range.start_column..=range.end_column {
                    if let Some(cell) = sheet.cell(row, column).cloned() {
                        sheet.set_cell(model::Cell { style: None, ..cell });
                    }
                }
            }
            0
        }
        "merge.set" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
            let start = operation
                .get("start")
                .and_then(Value::as_str)
                .and_then(reference::parse_reference)
                .ok_or_else(|| ServiceError::argument("`merge.set` needs an A1 `start` such as `A1`"))?;
            let end = operation
                .get("end")
                .and_then(Value::as_str)
                .and_then(reference::parse_reference)
                .ok_or_else(|| ServiceError::argument("`merge.set` needs an A1 `end` such as `B2`"))?;
            let merge = operation
                .get("merge")
                .and_then(Value::as_bool)
                .ok_or_else(|| ServiceError::argument("`merge.set` needs a boolean `merge`"))?;
            let range = CellRange::between(start.0, start.1, end.0, end.1);
            if !range.is_valid() || range.end_row > 1_048_576 || range.end_column > 16_384 {
                return Err(ServiceError::argument(format!("merged range `{}` leaves the sheet grid", range.reference())));
            }
            if merge {
                if let Some(other) = sheet.merged.iter().find(|other| **other != range && other.overlaps(&range)) {
                    return Err(ServiceError::argument(format!(
                        "merged range `{}` overlaps `{}`",
                        range.reference(),
                        other.reference()
                    )));
                }
                if !sheet.merged.contains(&range) {
                    sheet.merged.push(range);
                }
            }
            else {
                sheet.merged.retain(|existing| !existing.overlaps(&range));
            }
            0
        }
        "sheet.freeze" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let index = project.sheet_index(key).ok_or_else(|| missing_sheet(key))?;
            let count_of = |name: &str| -> Result<Option<u32>, ServiceError> {
                match operation.get(name) {
                    None | Some(Value::Null) => Ok(None),
                    Some(value) => value
                        .as_u64()
                        .and_then(|value| u32::try_from(value).ok())
                        .map(Some)
                        .ok_or_else(|| ServiceError::argument(format!("`{name}` must be a non-negative integer"))),
                }
            };
            let rows = count_of("rows")?;
            let columns = count_of("columns")?;
            let boundary = operation.get("cell").and_then(Value::as_str).map(str::trim).filter(|cell| !cell.is_empty());
            let clear = operation.get("clear").and_then(Value::as_bool).unwrap_or(false);
            if clear || (boundary.is_none() && rows.is_none() && columns.is_none()) {
                project.sheets[index].freeze = None;
            }
            else {
                let (rows, columns) = match boundary {
                    Some(cell) => {
                        if rows.is_some() || columns.is_some() {
                            return Err(ServiceError::argument(
                                "`sheet.freeze` takes either a `cell` boundary or explicit `rows` and `columns`",
                            ));
                        }
                        let (row, column) = reference::parse_reference(cell)
                            .ok_or_else(|| ServiceError::argument(format!("`{cell}` is not an A1 cell reference")))?;
                        (row - 1, column - 1)
                    }
                    None => {
                        let rows = rows.ok_or_else(|| {
                            ServiceError::argument("`sheet.freeze` needs `rows` and `columns` or a `cell` boundary")
                        })?;
                        let columns = columns.ok_or_else(|| {
                            ServiceError::argument("`sheet.freeze` needs `rows` and `columns` or a `cell` boundary")
                        })?;
                        (rows, columns)
                    }
                };
                let sheet = &project.sheets[index];
                if rows > 1_048_576 || columns > 16_384 {
                    return Err(ServiceError::argument("the frozen pane leaves the sheet grid"));
                }
                // A freeze band past the used rectangle would freeze empty space,
                // so the boundary stays inside the rows and columns the sheet uses.
                let used_rows = sheet.last_row();
                let used_columns = sheet.last_column();
                if (used_rows > 0 && rows > used_rows) || (used_columns > 0 && columns > used_columns) {
                    return Err(ServiceError::argument(format!(
                        "freezing {rows} rows and {columns} columns goes past the used range, which ends at row \
                         {used_rows} and column {used_columns}"
                    )));
                }
                project.sheets[index].freeze =
                    if rows == 0 && columns == 0 { None } else { Some(model::Freeze { columns, rows }) };
            }
            0
        }
        "column.setWidth" => {
            let column = layout_column(operation)?;
            let width = required_number(operation, "width", "`column.setWidth` needs a numeric `width`")?;
            if !(0.1..=255.0).contains(&width) {
                return Err(ServiceError::argument(format!("column width {width} must be between 0.1 and 255")));
            }
            let hidden = layout_hidden(operation)?;
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
            set_column_layout(sheet, column, width, hidden);
            0
        }
        "row.setHeight" => {
            let row = operation
                .get("row")
                .and_then(Value::as_u64)
                .filter(|row| (1..=1_048_576).contains(row))
                .ok_or_else(|| ServiceError::argument("`row.setHeight` needs a `row` between 1 and 1048576"))?
                as u32;
            let height = required_number(operation, "height", "`row.setHeight` needs a numeric `height`")?;
            if !(0.0..=409.0).contains(&height) {
                return Err(ServiceError::argument(format!("row height {height} must be between 0 and 409")));
            }
            let hidden = layout_hidden(operation)?;
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
            set_row_layout(sheet, row, height, hidden);
            0
        }
        "cell.replaceText" => {
            let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
            let sheet_name = project.sheet(key).map(|sheet| sheet.name.clone()).unwrap_or_else(|| key.to_owned());
            let requested = operation
                .get("range")
                .and_then(Value::as_str)
                .map(str::trim)
                .filter(|range| !range.is_empty());
            let requested = requested
                .map(|range| {
                    reference::parse_range(range)
                        .ok_or_else(|| ServiceError::argument(format!("`{range}` is not an A1 range such as `A1:C4`")))
                })
                .transpose()?;
            let find = operation
                .get("find")
                .and_then(Value::as_str)
                .filter(|find| !find.is_empty())
                .ok_or_else(|| ServiceError::argument("`cell.replaceText` needs a non-empty `find`"))?;
            let replacement = operation.get("replace").and_then(Value::as_str).unwrap_or("");
            let match_case = operation.get("matchCase").and_then(Value::as_bool).unwrap_or(false);
            let limit = match operation.get("limit") {
                None | Some(Value::Null) => None,
                Some(value) => Some(
                    value
                        .as_u64()
                        .and_then(|limit| usize::try_from(limit).ok())
                        .filter(|limit| *limit >= 1)
                        .ok_or_else(|| ServiceError::argument("`cell.replaceText` `limit` must be a positive integer"))?,
                ),
            };
            let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
            let range = match requested {
                Some(range) => range,
                None => sheet.actual_range().ok_or_else(|| {
                    ServiceError::argument("`cell.replaceText` needs an A1 range when the sheet has no content")
                })?,
            };
            let mut budget = limit.unwrap_or(usize::MAX);
            let mut replacements = 0usize;
            'cells: for row in range.start_row..=range.end_row {
                for column in range.start_column..=range.end_column {
                    if budget == 0 {
                        break 'cells;
                    }
                    let Some(mut cell) = sheet.cell(row, column).cloned()
                    else {
                        continue;
                    };
                    if cell.formula.is_some() {
                        continue;
                    }
                    let CellValue::Text { text } = &cell.value
                    else {
                        continue;
                    };
                    let (next, replaced) = replace_occurrences(text, find, replacement, match_case, budget);
                    if replaced > 0 {
                        budget -= replaced;
                        replacements += replaced;
                        cell.value = CellValue::text(next);
                        sheet.set_cell(cell);
                    }
                }
            }
            affected.push(json!({
                "op": "cell.replaceText",
                "sheet": sheet_name,
                "range": range.reference(),
                "replaced": replacements,
            }));
            replacements
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
            0
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
            0
        }
        "row.insert" => {
            structure::insert_rows(project, operation)?;
            0
        }
        "row.delete" => {
            structure::delete_rows(project, operation)?;
            0
        }
        "column.insert" => {
            structure::insert_columns(project, operation)?;
            0
        }
        "column.delete" => {
            structure::delete_columns(project, operation)?;
            0
        }
        other => return Err(ServiceError::new("object/unsupported", format!("edit operation `{other}` is not supported"))),
    };
    Ok(replacements)
}

/// Resolves the `column` of a `range.sort` request.
///
/// A number is a one-based offset inside the range; a letter is the absolute
/// sheet column. Either form must land inside the sorted range.
fn sort_column(operation: &Value, range: &CellRange) -> Result<u32, ServiceError> {
    let column = match operation.get("column") {
        Some(Value::Number(number)) => {
            let offset = number
                .as_u64()
                .filter(|offset| *offset >= 1)
                .ok_or_else(|| ServiceError::argument("the `column` offset must be a positive integer"))?;
            let offset = u32::try_from(offset).map_err(|_| ServiceError::argument("the `column` offset is too large"))?;
            range
                .start_column
                .checked_add(offset - 1)
                .ok_or_else(|| ServiceError::argument("the `column` offset leaves the sheet grid"))?
        }
        Some(Value::String(text)) => reference::column_number(text.trim())
            .ok_or_else(|| ServiceError::argument(format!("`{text}` is not a column letter such as `B`")))?,
        Some(_) => return Err(ServiceError::argument("`range.sort` needs a numeric or letter `column`")),
        None => return Err(ServiceError::argument("`range.sort` needs a `column`")),
    };
    if !range.contains(range.start_row, column) {
        return Err(ServiceError::argument(format!(
            "sort column `{}` is outside the range `{}`",
            reference::column_name(column),
            range.reference()
        )));
    }
    Ok(column)
}

/// Describes the rectangle one range operation rewrote.
fn range_effect(project: &Project, operation: &Value, op: &str, range: &CellRange) -> Value {
    let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
    let sheet = project.sheet(key).map(|sheet| sheet.name.clone()).unwrap_or_else(|| key.to_owned());
    json!({ "op": op, "sheet": sheet, "range": range.reference() })
}

/// Reads an optional finite number from an operation field.
fn optional_finite(operation: &Value, key: &str) -> Result<Option<f64>, ServiceError> {
    match operation.get(key) {
        None | Some(Value::Null) => Ok(None),
        Some(value) => value
            .as_f64()
            .filter(|value| value.is_finite())
            .map(Some)
            .ok_or_else(|| ServiceError::argument(format!("`{key}` must be a finite number"))),
    }
}

/// Fills one single-row or single-column range with a linear numeric series.
///
/// The first cell seeds the start and the first two cells seed the step unless
/// the request names `start` and `step` explicitly. Existing styles stay on
/// their cells, while formulas inside the range are replaced by the series.
fn fill_series(project: &mut Project, operation: &Value) -> Result<CellRange, ServiceError> {
    let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
    let range = operation
        .get("range")
        .and_then(Value::as_str)
        .and_then(reference::parse_range)
        .ok_or_else(|| ServiceError::argument("`range.fillSeries` needs an A1 range such as `A1:A5`"))?;
    if !range.is_valid() || range.end_row > 1_048_576 {
        return Err(ServiceError::argument(format!("series range `{}` leaves the sheet grid", range.reference())));
    }
    let vertical = match (range.rows() > 1, range.columns() > 1) {
        (true, false) => true,
        (false, true) => false,
        (false, false) => {
            return Err(ServiceError::argument("`range.fillSeries` needs a range that spans at least two cells"));
        }
        (true, true) => return Err(ServiceError::argument("`range.fillSeries` needs a single row or a single column")),
    };
    // Seed values come from evaluated cells, matching how sorting reads keys.
    eval::recalculate(project);
    let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
    if let Some(merged) = sheet.merged.iter().find(|merged| merged.overlaps(&range)) {
        return Err(ServiceError::argument(format!(
            "`range.fillSeries` cannot write through merged range `{}`",
            merged.reference()
        )));
    }
    let coordinates = if vertical {
        (range.start_row..=range.end_row).map(|row| (row, range.start_column)).collect::<Vec<_>>()
    }
    else {
        (range.start_column..=range.end_column).map(|column| (range.start_row, column)).collect::<Vec<_>>()
    };
    let seeds = coordinates
        .iter()
        .map(|(row, column)| sheet.cell(*row, *column).map(eval::value_of).unwrap_or(CellValue::Empty))
        .collect::<Vec<_>>();
    let (start, inferred) = match optional_finite(operation, "start")? {
        Some(start) => (start, false),
        None => match seeds.first() {
            Some(CellValue::Number { number }) => (*number, true),
            _ => {
                return Err(ServiceError::argument("`range.fillSeries` needs a numeric first cell or an explicit `start`"));
            }
        },
    };
    let step = match optional_finite(operation, "step")? {
        Some(step) => step,
        None => match (inferred, seeds.get(1)) {
            (true, Some(CellValue::Number { number })) => number - start,
            _ => 1.0,
        },
    };
    for (offset, (row, column)) in coordinates.iter().enumerate() {
        let style = sheet.cell(*row, *column).and_then(|cell| cell.style.clone());
        let cell = model::Cell {
            reference: reference::format_reference(*row, *column),
            row: *row,
            column: *column,
            value: CellValue::number(start + step * offset as f64),
            formula: None,
            cached: None,
            style,
            computed: None,
            computed_state: None,
        };
        sheet.set_cell(cell);
    }
    Ok(range)
}

/// Moves a rectangle into the transposed rectangle anchored at `target`.
///
/// Cell contents, styles and formula sources move verbatim; relative references
/// inside formulas are not rewritten. Merges and overlapping destinations are
/// rejected because the moved cells would lose their shape.
fn transpose(project: &mut Project, operation: &Value) -> Result<CellRange, ServiceError> {
    let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
    let range = operation
        .get("range")
        .and_then(Value::as_str)
        .and_then(reference::parse_range)
        .ok_or_else(|| ServiceError::argument("`range.transpose` needs an A1 range such as `A1:C2`"))?;
    if !range.is_valid() || range.end_row > 1_048_576 {
        return Err(ServiceError::argument(format!("transpose range `{}` leaves the sheet grid", range.reference())));
    }
    let (target_row, target_column) = operation
        .get("target")
        .and_then(Value::as_str)
        .and_then(reference::parse_reference)
        .ok_or_else(|| ServiceError::argument("`range.transpose` needs an A1 `target` anchor such as `E1`"))?;
    let end_row = target_row
        .checked_add(range.columns() - 1)
        .filter(|row| *row <= 1_048_576)
        .ok_or_else(|| ServiceError::argument("the transposed range leaves the sheet grid"))?;
    let end_column = target_column
        .checked_add(range.rows() - 1)
        .filter(|column| *column <= 16_384)
        .ok_or_else(|| ServiceError::argument("the transposed range leaves the sheet grid"))?;
    let destination = CellRange { start_row: target_row, end_row, start_column: target_column, end_column };
    if destination != range && destination.overlaps(&range) {
        return Err(ServiceError::argument(format!(
            "`range.transpose` cannot write `{}` over its source `{}`",
            destination.reference(),
            range.reference()
        )));
    }
    let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
    if let Some(merged) = sheet.merged.iter().find(|merged| merged.overlaps(&range) || merged.overlaps(&destination)) {
        return Err(ServiceError::argument(format!("`range.transpose` cannot move merged range `{}`", merged.reference())));
    }
    let matrix = (range.start_row..=range.end_row)
        .map(|row| (range.start_column..=range.end_column).map(|column| sheet.cell(row, column).cloned()).collect::<Vec<_>>())
        .collect::<Vec<_>>();
    sheet.clear_range(&range);
    for (row_offset, row) in matrix.iter().enumerate() {
        for (column_offset, cell) in row.iter().enumerate() {
            let Some(cell) = cell.clone()
            else {
                continue;
            };
            let row = destination.start_row + column_offset as u32;
            let column = destination.start_column + row_offset as u32;
            let mut moved = cell;
            moved.row = row;
            moved.column = column;
            moved.reference = reference::format_reference(row, column);
            sheet.set_cell(moved);
        }
    }
    Ok(destination)
}

/// Copies a rectangle into the rectangle anchored at `target`, keeping the source.
///
/// Cell contents, styles and formula sources move verbatim; relative references
/// inside formulas are not rewritten. The source is snapshotted before the
/// destination is cleared, so overlapping source and target stay deterministic.
/// Merges are rejected because copied cells would lose their shape.
fn copy_range(project: &mut Project, operation: &Value) -> Result<CellRange, ServiceError> {
    let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
    let range = operation
        .get("range")
        .and_then(Value::as_str)
        .and_then(reference::parse_range)
        .ok_or_else(|| ServiceError::argument("`range.copy` needs an A1 range such as `A1:C2`"))?;
    if !range.is_valid() || range.end_row > 1_048_576 {
        return Err(ServiceError::argument(format!("copy range `{}` leaves the sheet grid", range.reference())));
    }
    let (target_row, target_column) = operation
        .get("target")
        .and_then(Value::as_str)
        .and_then(reference::parse_reference)
        .ok_or_else(|| ServiceError::argument("`range.copy` needs an A1 `target` anchor such as `E1`"))?;
    let end_row = target_row
        .checked_add(range.rows() - 1)
        .filter(|row| *row <= 1_048_576)
        .ok_or_else(|| ServiceError::argument("the copied range leaves the sheet grid"))?;
    let end_column = target_column
        .checked_add(range.columns() - 1)
        .filter(|column| *column <= 16_384)
        .ok_or_else(|| ServiceError::argument("the copied range leaves the sheet grid"))?;
    let destination = CellRange { start_row: target_row, end_row, start_column: target_column, end_column };
    let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
    if let Some(merged) = sheet.merged.iter().find(|merged| merged.overlaps(&range) || merged.overlaps(&destination)) {
        return Err(ServiceError::argument(format!("`range.copy` cannot copy merged range `{}`", merged.reference())));
    }
    let matrix = (range.start_row..=range.end_row)
        .map(|row| (range.start_column..=range.end_column).map(|column| sheet.cell(row, column).cloned()).collect::<Vec<_>>())
        .collect::<Vec<_>>();
    if destination != range {
        sheet.clear_range(&destination);
    }
    for (row_offset, row) in matrix.iter().enumerate() {
        for (column_offset, cell) in row.iter().enumerate() {
            let Some(cell) = cell.clone()
            else {
                continue;
            };
            let row = destination.start_row + row_offset as u32;
            let column = destination.start_column + column_offset as u32;
            let mut copied = cell;
            copied.row = row;
            copied.column = column;
            copied.reference = reference::format_reference(row, column);
            sheet.set_cell(copied);
        }
    }
    Ok(destination)
}

/// Moves a rectangle into the rectangle anchored at `target`, clearing the source.
///
/// Cell contents, styles and formula sources travel verbatim; relative
/// references inside formulas are not rewritten. The source is snapshotted
/// before either rectangle is cleared, so overlapping source and target stay
/// deterministic. Merges are rejected because moved cells would lose their
/// shape.
fn move_range(project: &mut Project, operation: &Value) -> Result<CellRange, ServiceError> {
    let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
    let range = operation
        .get("range")
        .and_then(Value::as_str)
        .and_then(reference::parse_range)
        .ok_or_else(|| ServiceError::argument("`range.move` needs an A1 range such as `A1:C2`"))?;
    if !range.is_valid() || range.end_row > 1_048_576 || range.end_column > 16_384 {
        return Err(ServiceError::argument(format!("move range `{}` leaves the sheet grid", range.reference())));
    }
    let (target_row, target_column) = operation
        .get("target")
        .and_then(Value::as_str)
        .and_then(reference::parse_reference)
        .ok_or_else(|| ServiceError::argument("`range.move` needs an A1 `target` anchor such as `E1`"))?;
    let end_row = target_row
        .checked_add(range.rows() - 1)
        .filter(|row| *row <= 1_048_576)
        .ok_or_else(|| ServiceError::argument("the moved range leaves the sheet grid"))?;
    let end_column = target_column
        .checked_add(range.columns() - 1)
        .filter(|column| *column <= 16_384)
        .ok_or_else(|| ServiceError::argument("the moved range leaves the sheet grid"))?;
    let destination = CellRange { start_row: target_row, end_row, start_column: target_column, end_column };
    let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
    if let Some(merged) = sheet.merged.iter().find(|merged| merged.overlaps(&range) || merged.overlaps(&destination)) {
        return Err(ServiceError::argument(format!("`range.move` cannot move merged range `{}`", merged.reference())));
    }
    if destination == range {
        return Ok(destination);
    }
    let matrix = (range.start_row..=range.end_row)
        .map(|row| (range.start_column..=range.end_column).map(|column| sheet.cell(row, column).cloned()).collect::<Vec<_>>())
        .collect::<Vec<_>>();
    sheet.clear_range(&destination);
    sheet.clear_range(&range);
    for (row_offset, row) in matrix.iter().enumerate() {
        for (column_offset, cell) in row.iter().enumerate() {
            let Some(cell) = cell.clone()
            else {
                continue;
            };
            let row = destination.start_row + row_offset as u32;
            let column = destination.start_column + column_offset as u32;
            let mut moved = cell;
            moved.row = row;
            moved.column = column;
            moved.reference = reference::format_reference(row, column);
            sheet.set_cell(moved);
        }
    }
    Ok(destination)
}

/// Exchanges two equal-sized rectangles anchored at `range` and `target`.
///
/// Cell contents, styles and formula sources travel verbatim; relative
/// references inside formulas are not rewritten. Both rectangles are
/// snapshotted before either is cleared, so overlapping rectangles stay
/// deterministic. Merges are rejected because swapped cells would lose their
/// shape. The reported range is the bounding rectangle of both rectangles.
fn swap_range(project: &mut Project, operation: &Value) -> Result<CellRange, ServiceError> {
    let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
    let range = operation
        .get("range")
        .and_then(Value::as_str)
        .and_then(reference::parse_range)
        .ok_or_else(|| ServiceError::argument("`range.swap` needs an A1 range such as `A1:B2`"))?;
    if !range.is_valid() || range.end_row > 1_048_576 || range.end_column > 16_384 {
        return Err(ServiceError::argument(format!("swap range `{}` leaves the sheet grid", range.reference())));
    }
    let (target_row, target_column) = operation
        .get("target")
        .and_then(Value::as_str)
        .and_then(reference::parse_reference)
        .ok_or_else(|| ServiceError::argument("`range.swap` needs an A1 `target` anchor such as `E1`"))?;
    let end_row = target_row
        .checked_add(range.rows() - 1)
        .filter(|row| *row <= 1_048_576)
        .ok_or_else(|| ServiceError::argument("the swapped range leaves the sheet grid"))?;
    let end_column = target_column
        .checked_add(range.columns() - 1)
        .filter(|column| *column <= 16_384)
        .ok_or_else(|| ServiceError::argument("the swapped range leaves the sheet grid"))?;
    let destination = CellRange { start_row: target_row, end_row, start_column: target_column, end_column };
    let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
    if let Some(merged) = sheet.merged.iter().find(|merged| merged.overlaps(&range) || merged.overlaps(&destination)) {
        return Err(ServiceError::argument(format!("`range.swap` cannot swap merged range `{}`", merged.reference())));
    }
    if destination == range {
        return Ok(range);
    }
    let snapshot = |rect: &CellRange| -> Vec<Vec<Option<model::Cell>>> {
        (rect.start_row..=rect.end_row)
            .map(|row| (rect.start_column..=rect.end_column).map(|column| sheet.cell(row, column).cloned()).collect())
            .collect()
    };
    let source_cells = snapshot(&range);
    let target_cells = snapshot(&destination);
    sheet.clear_range(&range);
    sheet.clear_range(&destination);
    for (row_offset, row) in target_cells.iter().enumerate() {
        for (column_offset, cell) in row.iter().enumerate() {
            let Some(cell) = cell.clone()
            else {
                continue;
            };
            let row = range.start_row + row_offset as u32;
            let column = range.start_column + column_offset as u32;
            let mut moved = cell;
            moved.row = row;
            moved.column = column;
            moved.reference = reference::format_reference(row, column);
            sheet.set_cell(moved);
        }
    }
    for (row_offset, row) in source_cells.iter().enumerate() {
        for (column_offset, cell) in row.iter().enumerate() {
            let Some(cell) = cell.clone()
            else {
                continue;
            };
            let row = destination.start_row + row_offset as u32;
            let column = destination.start_column + column_offset as u32;
            let mut moved = cell;
            moved.row = row;
            moved.column = column;
            moved.reference = reference::format_reference(row, column);
            sheet.set_cell(moved);
        }
    }
    Ok(CellRange::between(
        range.start_row.min(destination.start_row),
        range.start_column.min(destination.start_column),
        range.end_row.max(destination.end_row),
        range.end_column.max(destination.end_column),
    ))
}

/// Mirrors a rectangle in place along `axis`.
///
/// `horizontal` swaps columns left to right while rows keep their position;
/// `vertical` swaps rows top to bottom while columns keep theirs. Cell
/// contents, styles and formula sources travel verbatim; relative references
/// inside formulas are not rewritten. The rectangle is snapshotted before it is
/// cleared, so every cell lands once. Merges are rejected because mirrored
/// cells would lose their shape. The reported range is the flipped rectangle.
fn flip_range(project: &mut Project, operation: &Value) -> Result<CellRange, ServiceError> {
    let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
    let range = operation
        .get("range")
        .and_then(Value::as_str)
        .and_then(reference::parse_range)
        .ok_or_else(|| ServiceError::argument("`range.flip` needs an A1 range such as `A1:C3`"))?;
    if !range.is_valid() || range.end_row > 1_048_576 || range.end_column > 16_384 {
        return Err(ServiceError::argument(format!("flip range `{}` leaves the sheet grid", range.reference())));
    }
    let requested = operation
        .get("axis")
        .and_then(Value::as_str)
        .map(str::trim)
        .filter(|axis| !axis.is_empty())
        .ok_or_else(|| ServiceError::argument("`range.flip` needs `axis` set to `horizontal` or `vertical`"))?;
    let horizontal = match requested.to_ascii_lowercase().as_str() {
        "horizontal" => true,
        "vertical" => false,
        _ => {
            return Err(ServiceError::argument(format!(
                "`{requested}` is not a flip axis; expected `horizontal` or `vertical`"
            )));
        }
    };
    let sheet = project.sheet_mut(key).ok_or_else(|| missing_sheet(key))?;
    if let Some(merged) = sheet.merged.iter().find(|merged| merged.overlaps(&range)) {
        return Err(ServiceError::argument(format!("`range.flip` cannot flip merged range `{}`", merged.reference())));
    }
    if (horizontal && range.columns() == 1) || (!horizontal && range.rows() == 1) {
        return Ok(range);
    }
    let matrix = (range.start_row..=range.end_row)
        .map(|row| (range.start_column..=range.end_column).map(|column| sheet.cell(row, column).cloned()).collect::<Vec<_>>())
        .collect::<Vec<_>>();
    sheet.clear_range(&range);
    for (row_offset, row) in matrix.iter().enumerate() {
        for (column_offset, cell) in row.iter().enumerate() {
            let Some(cell) = cell.clone()
            else {
                continue;
            };
            let (row, column) = if horizontal {
                (range.start_row + row_offset as u32, range.end_column - column_offset as u32)
            }
            else {
                (range.end_row - row_offset as u32, range.start_column + column_offset as u32)
            };
            let mut moved = cell;
            moved.row = row;
            moved.column = column;
            moved.reference = reference::format_reference(row, column);
            sheet.set_cell(moved);
        }
    }
    Ok(range)
}

/// Reads a finite numeric field of an operation.
fn required_number(operation: &Value, key: &str, message: &str) -> Result<f64, ServiceError> {
    operation.get(key).and_then(Value::as_f64).filter(|value| value.is_finite()).ok_or_else(|| ServiceError::argument(message))
}

/// Reads the optional `hidden` flag of a layout operation.
fn layout_hidden(operation: &Value) -> Result<Option<bool>, ServiceError> {
    match operation.get("hidden") {
        None | Some(Value::Null) => Ok(None),
        Some(Value::Bool(value)) => Ok(Some(*value)),
        Some(_) => Err(ServiceError::argument("`hidden` must be a boolean")),
    }
}

/// Resolves the one-based column of a `column.setWidth` request.
///
/// Accepts a one-based number or a column letter such as `B`.
fn layout_column(operation: &Value) -> Result<u32, ServiceError> {
    let column = match operation.get("column") {
        Some(Value::Number(number)) => number
            .as_u64()
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| ServiceError::argument("`column.setWidth` needs a positive `column`"))?,
        Some(Value::String(text)) => reference::column_number(text.trim())
            .ok_or_else(|| ServiceError::argument(format!("`{text}` is not a column letter such as `B`")))?,
        Some(_) => return Err(ServiceError::argument("`column.setWidth` needs a numeric or letter `column`")),
        None => return Err(ServiceError::argument("`column.setWidth` needs a `column`")),
    };
    if column == 0 || column > 16_384 {
        return Err(ServiceError::argument("the column leaves the sheet grid"));
    }
    Ok(column)
}

/// Writes the width of one column, splitting the record that covers it and
/// merging adjacent records that share the same width, custom and hidden flag.
fn set_column_layout(sheet: &mut Sheet, column: u32, width: f64, hidden: Option<bool>) {
    let hidden = hidden.unwrap_or_else(|| {
        sheet
            .columns
            .iter()
            .rev()
            .find(|record| record.min <= column && column <= record.max)
            .map(|record| record.hidden)
            .unwrap_or(false)
    });
    let mut records: Vec<SheetColumn> = Vec::with_capacity(sheet.columns.len() + 2);
    for record in std::mem::take(&mut sheet.columns) {
        if record.max < column || record.min > column {
            records.push(record);
            continue;
        }
        if record.min < column {
            records.push(SheetColumn { max: column - 1, ..record.clone() });
        }
        if record.max > column {
            records.push(SheetColumn { min: column + 1, ..record });
        }
    }
    records.push(SheetColumn { min: column, max: column, width, custom: true, hidden });
    records.sort_by_key(|record| record.min);
    let mut merged: Vec<SheetColumn> = Vec::with_capacity(records.len());
    for record in records {
        match merged.last_mut() {
            Some(last)
                if last.max + 1 == record.min
                    && last.width == record.width
                    && last.custom == record.custom
                    && last.hidden == record.hidden =>
            {
                last.max = record.max;
            }
            _ => merged.push(record),
        }
    }
    sheet.columns = merged;
}

/// Writes the height of one row, keeping the row records sorted by index.
fn set_row_layout(sheet: &mut Sheet, row: u32, height: f64, hidden: Option<bool>) {
    if let Some(record) = sheet.rows.iter_mut().find(|record| record.index == row) {
        record.height = Some(height);
        record.custom_height = true;
        if let Some(hidden) = hidden {
            record.hidden = hidden;
        }
        return;
    }
    sheet.rows.push(SheetRow { index: row, height: Some(height), custom_height: true, hidden: hidden.unwrap_or(false) });
    sheet.rows.sort_by_key(|record| record.index);
}

/// Orders two sort keys; blanks always trail, everything else follows the
/// number < text < boolean < error ranking and flips with `ascending`.
fn compare_sort_values(left: &CellValue, right: &CellValue, ascending: bool) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    match (left.is_empty(), right.is_empty()) {
        (true, true) => return Ordering::Equal,
        (true, false) => return Ordering::Greater,
        (false, true) => return Ordering::Less,
        _ => {}
    }
    let ordering = match (left, right) {
        (CellValue::Number { number: left }, CellValue::Number { number: right }) => {
            left.partial_cmp(right).unwrap_or(Ordering::Equal)
        }
        (CellValue::Number { .. }, _) => Ordering::Less,
        (_, CellValue::Number { .. }) => Ordering::Greater,
        (CellValue::Text { text: left }, CellValue::Text { text: right }) => {
            left.to_lowercase().cmp(&right.to_lowercase()).then_with(|| left.cmp(right))
        }
        (CellValue::Text { .. }, _) => Ordering::Less,
        (_, CellValue::Text { .. }) => Ordering::Greater,
        (CellValue::Bool { value: left }, CellValue::Bool { value: right }) => left.cmp(right),
        (CellValue::Bool { .. }, _) => Ordering::Less,
        (_, CellValue::Bool { .. }) => Ordering::Greater,
        (CellValue::Error { code: left }, CellValue::Error { code: right }) => left.cmp(right),
        (CellValue::Error { .. }, _) => Ordering::Less,
        (_, CellValue::Error { .. }) => Ordering::Greater,
        (CellValue::Empty, CellValue::Empty) => Ordering::Equal,
    };
    if ascending { ordering } else { ordering.reverse() }
}

/// Replaces up to `limit` non-overlapping occurrences and reports how many it replaced.
fn replace_occurrences(text: &str, find: &str, replacement: &str, match_case: bool, limit: usize) -> (String, usize) {
    let haystack: Vec<char> = text.chars().collect();
    let needle: Vec<char> = find.chars().collect();
    if needle.is_empty() || limit == 0 {
        return (text.to_owned(), 0);
    }
    let mut output = String::new();
    let mut replacements = 0usize;
    let mut index = 0usize;
    while index < haystack.len() {
        if replacements < limit
            && index + needle.len() <= haystack.len()
            && haystack[index..index + needle.len()]
                .iter()
                .zip(&needle)
                .all(|(left, right)| char_matches(*left, *right, match_case))
        {
            output.push_str(replacement);
            replacements += 1;
            index += needle.len();
        }
        else {
            output.push(haystack[index]);
            index += 1;
        }
    }
    (output, replacements)
}

fn char_matches(left: char, right: char, match_case: bool) -> bool {
    left == right || (!match_case && left.eq_ignore_ascii_case(&right))
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

pub(crate) fn missing_sheet(key: &str) -> ServiceError {
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
    let options = PublishOptions {
        overwrite: if overwrite { OverwritePolicy::Replace } else { OverwritePolicy::RefuseExisting },
        ..PublishOptions::default()
    };
    publish_bytes_with_options(publish_target(path), bytes, options)
        .map_err(|error| ServiceError::io(format!("failed to write {}: {error}", path.display())))
}

/// Keeps a bare file name anchored to the working directory so the publisher
/// always has a folder to stage in.
fn publish_target(path: &Path) -> PathBuf {
    match path.parent() {
        Some(parent) if !parent.as_os_str().is_empty() => path.to_path_buf(),
        _ => PathBuf::from(".").join(path),
    }
}
