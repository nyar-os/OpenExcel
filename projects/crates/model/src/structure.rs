//! Whole-row and whole-column structural edits.
//!
//! These edits keep every coordinate-bearing model field in step: cells, merged
//! ranges, frozen panes, column widths, row heights and the declared dimension.
//! Formula text moves through [`crate::shift`]. The row and column arguments are
//! one-based, matching the rest of the model; a column may also be a letter.

use serde_json::Value;

use crate::{
    error::ServiceError,
    model::{CellRange, Project, Sheet, SheetColumn},
    reference,
    shift::{self, Axis, Shift},
};

/// Last row of an XLSX sheet.
const MAX_ROW: u32 = 1_048_576;
/// Last column of an XLSX sheet, `XFD`.
const MAX_COLUMN: u32 = 16_384;

/// Inserts `count` empty rows before `row`.
pub(crate) fn insert_rows(project: &mut Project, operation: &Value) -> Result<(), ServiceError> {
    let index = sheet_index(project, operation)?;
    let start = required_position(operation, "row", "row.insert", MAX_ROW)?;
    let count = count_of(operation)?;
    ensure_band(start, count, MAX_ROW, "row.insert", "row")?;
    check_insert_rows(&project.sheets[index], start, count)?;
    let target = project.sheets[index].name.clone();
    shift_rows(&mut project.sheets[index], start, count);
    shift::shift_project(project, &target, Axis::Row, Shift::Insert { at: start, count });
    project.normalize();
    Ok(())
}

/// Deletes `count` rows starting at `row`.
pub(crate) fn delete_rows(project: &mut Project, operation: &Value) -> Result<(), ServiceError> {
    let index = sheet_index(project, operation)?;
    let start = required_position(operation, "row", "row.delete", MAX_ROW)?;
    let count = count_of(operation)?;
    let end = ensure_band(start, count, MAX_ROW, "row.delete", "row")?;
    let target = project.sheets[index].name.clone();
    delete_rows_from(&mut project.sheets[index], start, end);
    shift::shift_project(project, &target, Axis::Row, Shift::Delete { start, end });
    project.normalize();
    Ok(())
}

/// Inserts `count` empty columns before `column`.
pub(crate) fn insert_columns(project: &mut Project, operation: &Value) -> Result<(), ServiceError> {
    let index = sheet_index(project, operation)?;
    let start = required_column(operation, "column.insert")?;
    let count = count_of(operation)?;
    ensure_band(start, count, MAX_COLUMN, "column.insert", "column")?;
    check_insert_columns(&project.sheets[index], start, count)?;
    let target = project.sheets[index].name.clone();
    shift_columns(&mut project.sheets[index], start, count);
    shift::shift_project(project, &target, Axis::Column, Shift::Insert { at: start, count });
    project.normalize();
    Ok(())
}

/// Deletes `count` columns starting at `column`.
pub(crate) fn delete_columns(project: &mut Project, operation: &Value) -> Result<(), ServiceError> {
    let index = sheet_index(project, operation)?;
    let start = required_column(operation, "column.delete")?;
    let count = count_of(operation)?;
    let end = ensure_band(start, count, MAX_COLUMN, "column.delete", "column")?;
    let target = project.sheets[index].name.clone();
    delete_columns_from(&mut project.sheets[index], start, end);
    shift::shift_project(project, &target, Axis::Column, Shift::Delete { start, end });
    project.normalize();
    Ok(())
}

/// Resolves the sheet an operation targets.
fn sheet_index(project: &Project, operation: &Value) -> Result<usize, ServiceError> {
    let key = operation.get("sheet").and_then(Value::as_str).unwrap_or("");
    project.sheet_index(key).ok_or_else(|| crate::missing_sheet(key))
}

/// Reads the required one-based row or column of an operation.
fn required_position(operation: &Value, key: &str, op: &str, max: u32) -> Result<u32, ServiceError> {
    let value = operation
        .get(key)
        .and_then(Value::as_u64)
        .and_then(|value| u32::try_from(value).ok())
        .filter(|value| (1..=max).contains(value))
        .ok_or_else(|| ServiceError::argument(format!("`{op}` needs a one-based `{key}` between 1 and {max}")))?;
    Ok(value)
}

/// Reads the required column of an operation as a letter or a one-based number.
fn required_column(operation: &Value, op: &str) -> Result<u32, ServiceError> {
    let column = match operation.get("column") {
        Some(Value::Number(number)) => number
            .as_u64()
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| ServiceError::argument(format!("`{op}` needs a one-based `column` or a column letter")))?,
        Some(Value::String(text)) => reference::column_number(text.trim())
            .ok_or_else(|| ServiceError::argument(format!("`{text}` is not a column letter such as `B`")))?,
        Some(_) => {
            return Err(ServiceError::argument(format!("`{op}` needs a one-based `column` or a column letter")));
        }
        None => return Err(ServiceError::argument(format!("`{op}` needs a `column`"))),
    };
    if column == 0 || column > MAX_COLUMN {
        return Err(ServiceError::argument(format!("`{op}` needs a `column` between 1 and {MAX_COLUMN}")));
    }
    Ok(column)
}

/// Reads the optional `count`, defaulting to one.
fn count_of(operation: &Value) -> Result<u32, ServiceError> {
    match operation.get("count") {
        None | Some(Value::Null) => Ok(1),
        Some(value) => value
            .as_u64()
            .and_then(|count| u32::try_from(count).ok())
            .filter(|count| *count >= 1)
            .ok_or_else(|| ServiceError::argument("`count` must be a positive integer")),
    }
}

/// Rejects a band that leaves the sheet grid instead of truncating it.
fn ensure_band(start: u32, count: u32, max: u32, op: &str, unit: &str) -> Result<u32, ServiceError> {
    start.checked_add(count - 1).filter(|end| *end <= max).ok_or_else(|| {
        ServiceError::argument(format!("`{op}` cannot fit {count} {unit}s starting at {start}; the sheet grid ends at {max}"))
    })
}

/// Rejects a row insert that would push existing content past the last row.
fn check_insert_rows(sheet: &Sheet, start: u32, count: u32) -> Result<(), ServiceError> {
    let overflows = |value: u32| value >= start && value + count > MAX_ROW;
    if sheet.cells.iter().any(|cell| overflows(cell.row))
        || sheet.merged.iter().any(|range| overflows(range.end_row))
        || sheet.rows.iter().any(|record| overflows(record.index))
        || sheet.dimension.is_some_and(|range| overflows(range.end_row))
        || sheet.auto_filter.is_some_and(|range| overflows(range.end_row))
        || sheet.freeze.is_some_and(|freeze| start <= freeze.rows && freeze.rows + count > MAX_ROW)
    {
        return Err(ServiceError::argument(format!(
            "inserting {count} rows before row {start} would push content past row {MAX_ROW}"
        )));
    }
    Ok(())
}

/// Rejects a column insert that would push existing content past the last column.
fn check_insert_columns(sheet: &Sheet, start: u32, count: u32) -> Result<(), ServiceError> {
    let overflows = |value: u32| value >= start && value + count > MAX_COLUMN;
    if sheet.cells.iter().any(|cell| overflows(cell.column))
        || sheet.merged.iter().any(|range| overflows(range.end_column))
        || sheet.columns.iter().any(|record| overflows(record.max))
        || sheet.dimension.is_some_and(|range| overflows(range.end_column))
        || sheet.auto_filter.is_some_and(|range| overflows(range.end_column))
        || sheet.freeze.is_some_and(|freeze| start <= freeze.columns && freeze.columns + count > MAX_COLUMN)
    {
        return Err(ServiceError::argument(format!(
            "inserting {count} columns before column {start} would push content past column {MAX_COLUMN}"
        )));
    }
    Ok(())
}

/// Shifts every row-bearing field down by `count`.
fn shift_rows(sheet: &mut Sheet, start: u32, count: u32) {
    for cell in &mut sheet.cells {
        if cell.row >= start {
            cell.row += count;
        }
    }
    for range in &mut sheet.merged {
        if range.end_row < start {
            continue;
        }
        if range.start_row >= start {
            range.start_row += count;
        }
        range.end_row += count;
    }
    for record in &mut sheet.rows {
        if record.index >= start {
            record.index += count;
        }
    }
    sheet.rows.sort_by_key(|record| record.index);
    if let Some(freeze) = &mut sheet.freeze
        && start <= freeze.rows
    {
        freeze.rows += count;
    }
    if let Some(dimension) = &mut sheet.dimension
        && dimension.end_row >= start
    {
        if dimension.start_row >= start {
            dimension.start_row += count;
        }
        dimension.end_row += count;
    }
    if let Some(filter) = &mut sheet.auto_filter
        && filter.end_row >= start
    {
        if filter.start_row >= start {
            filter.start_row += count;
        }
        filter.end_row += count;
    }
}

/// Removes rows `start..=end` and shifts the rows below them up.
fn delete_rows_from(sheet: &mut Sheet, start: u32, end: u32) {
    let count = end - start + 1;
    sheet.cells.retain(|cell| !(start..=end).contains(&cell.row));
    for cell in &mut sheet.cells {
        if cell.row > end {
            cell.row -= count;
        }
    }
    sheet.merged = std::mem::take(&mut sheet.merged)
        .into_iter()
        .filter_map(|mut range| {
            let (start_row, end_row) = delete_span((range.start_row, range.end_row), start, end)?;
            range.start_row = start_row;
            range.end_row = end_row;
            (range.rows() > 1 || range.columns() > 1).then_some(range)
        })
        .collect();
    sheet.rows.retain(|record| !(start..=end).contains(&record.index));
    for record in &mut sheet.rows {
        if record.index > end {
            record.index -= count;
        }
    }
    if let Some(freeze) = &mut sheet.freeze
        && start <= freeze.rows
    {
        freeze.rows = if freeze.rows > end { freeze.rows - count } else { start - 1 };
    }
    if sheet.freeze.is_some_and(|freeze| freeze.rows == 0 && freeze.columns == 0) {
        sheet.freeze = None;
    }
    if let Some(dimension) = sheet.dimension {
        sheet.dimension = delete_span((dimension.start_row, dimension.end_row), start, end)
            .map(|(start_row, end_row)| CellRange { start_row, end_row, ..dimension });
    }
    if let Some(filter) = sheet.auto_filter {
        sheet.auto_filter = delete_span((filter.start_row, filter.end_row), start, end)
            .map(|(start_row, end_row)| CellRange { start_row, end_row, ..filter });
    }
}

/// Shifts every column-bearing field right by `count`.
fn shift_columns(sheet: &mut Sheet, start: u32, count: u32) {
    for cell in &mut sheet.cells {
        if cell.column >= start {
            cell.column += count;
        }
    }
    for range in &mut sheet.merged {
        if range.end_column < start {
            continue;
        }
        if range.start_column >= start {
            range.start_column += count;
        }
        range.end_column += count;
    }
    for record in &mut sheet.columns {
        let (min, max) = insert_span((record.min, record.max), start, count);
        record.min = min;
        record.max = max;
    }
    sheet.columns.sort_by_key(|record| record.min);
    if let Some(freeze) = &mut sheet.freeze
        && start <= freeze.columns
    {
        freeze.columns += count;
    }
    if let Some(dimension) = &mut sheet.dimension
        && dimension.end_column >= start
    {
        if dimension.start_column >= start {
            dimension.start_column += count;
        }
        dimension.end_column += count;
    }
    if let Some(filter) = &mut sheet.auto_filter
        && filter.end_column >= start
    {
        if filter.start_column >= start {
            filter.start_column += count;
        }
        filter.end_column += count;
    }
}

/// Removes columns `start..=end` and shifts the columns right of them left.
fn delete_columns_from(sheet: &mut Sheet, start: u32, end: u32) {
    let count = end - start + 1;
    sheet.cells.retain(|cell| !(start..=end).contains(&cell.column));
    for cell in &mut sheet.cells {
        if cell.column > end {
            cell.column -= count;
        }
    }
    sheet.merged = std::mem::take(&mut sheet.merged)
        .into_iter()
        .filter_map(|mut range| {
            let (start_column, end_column) = delete_span((range.start_column, range.end_column), start, end)?;
            range.start_column = start_column;
            range.end_column = end_column;
            (range.rows() > 1 || range.columns() > 1).then_some(range)
        })
        .collect();
    sheet.columns = std::mem::take(&mut sheet.columns)
        .into_iter()
        .filter_map(|record| {
            let (min, max) = delete_span((record.min, record.max), start, end)?;
            Some(SheetColumn { min, max, ..record })
        })
        .collect();
    if let Some(freeze) = &mut sheet.freeze
        && start <= freeze.columns
    {
        freeze.columns = if freeze.columns > end { freeze.columns - count } else { start - 1 };
    }
    if sheet.freeze.is_some_and(|freeze| freeze.rows == 0 && freeze.columns == 0) {
        sheet.freeze = None;
    }
    if let Some(dimension) = sheet.dimension {
        sheet.dimension = delete_span((dimension.start_column, dimension.end_column), start, end)
            .map(|(start_column, end_column)| CellRange { start_column, end_column, ..dimension });
    }
    if let Some(filter) = sheet.auto_filter {
        sheet.auto_filter = delete_span((filter.start_column, filter.end_column), start, end)
            .map(|(start_column, end_column)| CellRange { start_column, end_column, ..filter });
    }
}

/// Expands or moves a span across an insertion point.
fn insert_span(span: (u32, u32), start: u32, count: u32) -> (u32, u32) {
    let (low, high) = span;
    if high < start {
        (low, high)
    }
    else if low >= start {
        (low + count, high + count)
    }
    else {
        (low, high + count)
    }
}

/// Shrinks or moves a span across a deleted band; `None` means it vanished.
fn delete_span(span: (u32, u32), start: u32, end: u32) -> Option<(u32, u32)> {
    let (low, high) = span;
    if high < start {
        return Some((low, high));
    }
    if low > end {
        let count = end - start + 1;
        return Some((low - count, high - count));
    }
    let new_low = if low < start { low } else { start };
    let new_high = if high > end { high - (end - start + 1) } else { start - 1 };
    (new_high >= new_low).then_some((new_low, new_high))
}
