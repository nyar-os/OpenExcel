//! Whole-row and whole-column structural edits.
//!
//! These edits keep every coordinate-bearing model field in step: cells, merged
//! ranges, frozen panes, column widths, row heights and the declared dimension.
//! Formula text moves through [`crate::shift`]. The row and column arguments are
//! one-based, matching the rest of the model; a column may also be a letter.

use serde_json::Value;

use crate::{
    error::ServiceError,
    model::{Cell, CellRange, NamedRange, Project, Sheet, SheetColumn},
    reference,
    shift::{self, Axis, Shift},
};

/// Last row of an XLSX sheet.
const MAX_ROW: u32 = 1_048_576;
/// Last column of an XLSX sheet, `XFD`.
const MAX_COLUMN: u32 = 16_384;
/// Default column width in Excel character units, matching `Sheet::column_width`.
const DEFAULT_WIDTH: f64 = 8.43;

/// Inserts `count` empty rows before `row`.
pub(crate) fn insert_rows(project: &mut Project, operation: &Value) -> Result<(), ServiceError> {
    let index = sheet_index(project, operation)?;
    let start = required_position(operation, "row", "row.insert", MAX_ROW)?;
    let count = count_of(operation)?;
    ensure_band(start, count, MAX_ROW, "row.insert", "row")?;
    check_insert_rows(&project.sheets[index], start, count)?;
    let target = project.sheets[index].name.clone();
    let id = project.sheets[index].id.clone();
    check_insert_names(&project.names, &id, start, count, MAX_ROW, |range| range.end_row, "row")?;
    shift_rows(&mut project.sheets[index], start, count);
    shift_named_rows(&mut project.names, &id, start, count);
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
    let id = project.sheets[index].id.clone();
    delete_rows_from(&mut project.sheets[index], start, end);
    delete_named_rows(&mut project.names, &id, start, end);
    shift::shift_project(project, &target, Axis::Row, Shift::Delete { start, end });
    project.normalize();
    Ok(())
}

/// Inserts `count` empty columns before `column`.
pub(crate) fn insert_columns(project: &mut Project, operation: &Value) -> Result<(), ServiceError> {
    let index = sheet_index(project, operation)?;
    let start = required_column(operation, "column", "column.insert")?;
    let count = count_of(operation)?;
    ensure_band(start, count, MAX_COLUMN, "column.insert", "column")?;
    check_insert_columns(&project.sheets[index], start, count)?;
    let target = project.sheets[index].name.clone();
    let id = project.sheets[index].id.clone();
    check_insert_names(&project.names, &id, start, count, MAX_COLUMN, |range| range.end_column, "column")?;
    shift_columns(&mut project.sheets[index], start, count);
    shift_named_columns(&mut project.names, &id, start, count);
    shift::shift_project(project, &target, Axis::Column, Shift::Insert { at: start, count });
    project.normalize();
    Ok(())
}

/// Deletes `count` columns starting at `column`.
pub(crate) fn delete_columns(project: &mut Project, operation: &Value) -> Result<(), ServiceError> {
    let index = sheet_index(project, operation)?;
    let start = required_column(operation, "column", "column.delete")?;
    let count = count_of(operation)?;
    let end = ensure_band(start, count, MAX_COLUMN, "column.delete", "column")?;
    let target = project.sheets[index].name.clone();
    let id = project.sheets[index].id.clone();
    delete_columns_from(&mut project.sheets[index], start, end);
    delete_named_columns(&mut project.names, &id, start, end);
    shift::shift_project(project, &target, Axis::Column, Shift::Delete { start, end });
    project.normalize();
    Ok(())
}

/// Exchanges two whole rows.
///
/// Cells, row heights and merges that sit entirely on one of the two rows move
/// with their row, and local formula references to either row follow it. A
/// merge that straddles one of the rows is rejected because the exchange would
/// tear its shape apart.
pub(crate) fn swap_rows(project: &mut Project, operation: &Value) -> Result<(), ServiceError> {
    let index = sheet_index(project, operation)?;
    let first = required_position(operation, "row", "row.swap", MAX_ROW)?;
    let second = required_position(operation, "other", "row.swap", MAX_ROW)?;
    if first == second {
        return Ok(());
    }
    let (first, second) = (first.min(second), first.max(second));
    let sheet = &mut project.sheets[index];
    check_swap_merges(&sheet.merged, Axis::Row, first, second, "row.swap")?;
    swap_cells(sheet, Axis::Row, first, second);
    swap_row_records(sheet, first, second);
    swap_line_merges(sheet, Axis::Row, first, second);
    let target = project.sheets[index].name.clone();
    shift::shift_project(project, &target, Axis::Row, Shift::Swap { first, second });
    project.normalize();
    Ok(())
}

/// Exchanges two whole columns.
///
/// Cells, column widths and merges that sit entirely on one of the two columns
/// move with their column, and local formula references to either column follow
/// it. A merge that straddles one of the columns is rejected.
pub(crate) fn swap_columns(project: &mut Project, operation: &Value) -> Result<(), ServiceError> {
    let index = sheet_index(project, operation)?;
    let first = required_column(operation, "column", "column.swap")?;
    let second = required_column(operation, "other", "column.swap")?;
    if first == second {
        return Ok(());
    }
    let (first, second) = (first.min(second), first.max(second));
    let sheet = &mut project.sheets[index];
    check_swap_merges(&sheet.merged, Axis::Column, first, second, "column.swap")?;
    swap_cells(sheet, Axis::Column, first, second);
    swap_column_records(sheet, first, second);
    swap_line_merges(sheet, Axis::Column, first, second);
    let target = project.sheets[index].name.clone();
    shift::shift_project(project, &target, Axis::Column, Shift::Swap { first, second });
    project.normalize();
    Ok(())
}

/// Inserts an empty rectangle, shifting the cells below or beside it over.
///
/// `shift` is `down` or `right`. Cells inside the band move as a block with
/// their styles and formula sources; references elsewhere on the sheet follow
/// them. Merges fully inside the shifted band move with it, while a merge that
/// crosses the insertion line is rejected. Row heights, column widths, frozen
/// panes and defined names stay put because only part of the grid moves.
pub(crate) fn insert_range(project: &mut Project, operation: &Value) -> Result<CellRange, ServiceError> {
    let index = sheet_index(project, operation)?;
    let range = required_rectangle(operation, "range.insert")?;
    let direction = required_direction(operation, "range.insert", "down", "right")?;
    let axis = if direction == "down" { Axis::Row } else { Axis::Column };
    let (at, count, across_start, across_end) = match axis {
        Axis::Row => (range.start_row, range.rows(), range.start_column, range.end_column),
        Axis::Column => (range.start_column, range.columns(), range.start_row, range.end_row),
    };
    let sheet = &mut project.sheets[index];
    check_band_insert(sheet, at, count, across_start, across_end, axis)?;
    let mut moved: Vec<Cell> = Vec::new();
    sheet.cells.retain(|cell| {
        let (line, across) = line_across(cell, axis);
        if (across_start..=across_end).contains(&across) && line >= at {
            moved.push(cell.clone());
            false
        }
        else {
            true
        }
    });
    for mut cell in moved {
        match axis {
            Axis::Row => cell.row += count,
            Axis::Column => cell.column += count,
        }
        cell.reference = reference::format_reference(cell.row, cell.column);
        sheet.set_cell(cell);
    }
    for merged in &mut sheet.merged {
        let (low, _, across_low, across_high) = band_span(merged, axis);
        if across_low <= across_end && across_high >= across_start && low >= at {
            match axis {
                Axis::Row => {
                    merged.start_row += count;
                    merged.end_row += count;
                }
                Axis::Column => {
                    merged.start_column += count;
                    merged.end_column += count;
                }
            }
        }
    }
    grow_band_span(sheet, axis, at, count);
    let target = project.sheets[index].name.clone();
    shift::shift_project(project, &target, axis, Shift::BandInsert { at, count, across_start, across_end });
    project.normalize();
    Ok(range)
}

/// Deletes a rectangle, shifting the cells below or beside it back.
///
/// `shift` is `up` or `left`. Cells inside the rectangle are removed and the
/// rest of the band moves up or left; references elsewhere follow. Merges fully
/// inside the band vanish, while a merge crossing the rectangle is rejected.
/// Row heights, column widths, frozen panes and defined names stay put.
pub(crate) fn delete_range(project: &mut Project, operation: &Value) -> Result<CellRange, ServiceError> {
    let index = sheet_index(project, operation)?;
    let range = required_rectangle(operation, "range.delete")?;
    let direction = required_direction(operation, "range.delete", "up", "left")?;
    let axis = if direction == "up" { Axis::Row } else { Axis::Column };
    let (start, end, across_start, across_end) = match axis {
        Axis::Row => (range.start_row, range.end_row, range.start_column, range.end_column),
        Axis::Column => (range.start_column, range.end_column, range.start_row, range.end_row),
    };
    let count = end - start + 1;
    let sheet = &mut project.sheets[index];
    for merged in &sheet.merged {
        let (low, high, across_low, across_high) = band_span(merged, axis);
        let overlaps_band = across_low <= across_end && across_high >= across_start;
        if overlaps_band && low <= end && high >= start && !(low >= start && high <= end) {
            return Err(ServiceError::argument(format!(
                "`range.delete` cannot remove merged range `{}` that crosses the rectangle",
                merged.reference()
            )));
        }
    }
    sheet.cells.retain(|cell| {
        let (line, across) = line_across(cell, axis);
        !((across_start..=across_end).contains(&across) && (start..=end).contains(&line))
    });
    for cell in &mut sheet.cells {
        let (line, across) = line_across(cell, axis);
        if (across_start..=across_end).contains(&across) && line > end {
            match axis {
                Axis::Row => cell.row -= count,
                Axis::Column => cell.column -= count,
            }
        }
    }
    sheet.merged = std::mem::take(&mut sheet.merged)
        .into_iter()
        .filter_map(|mut merged| {
            let (low, high, across_low, across_high) = band_span(&merged, axis);
            let overlaps_band = across_low <= across_end && across_high >= across_start;
            if overlaps_band {
                let (low, high) = delete_span((low, high), start, end)?;
                match axis {
                    Axis::Row => {
                        merged.start_row = low;
                        merged.end_row = high;
                    }
                    Axis::Column => {
                        merged.start_column = low;
                        merged.end_column = high;
                    }
                }
            }
            (merged.rows() > 1 || merged.columns() > 1).then_some(merged)
        })
        .collect();
    let target = project.sheets[index].name.clone();
    shift::shift_project(project, &target, axis, Shift::BandDelete { start, end, across_start, across_end });
    project.normalize();
    Ok(range)
}

/// Reads the required A1 rectangle of a range edit.
fn required_rectangle(operation: &Value, op: &str) -> Result<CellRange, ServiceError> {
    let range = operation
        .get("range")
        .and_then(Value::as_str)
        .and_then(reference::parse_range)
        .ok_or_else(|| ServiceError::argument(format!("`{op}` needs an A1 range such as `A1:B2`")))?;
    if !range.is_valid() || range.end_row > MAX_ROW || range.end_column > MAX_COLUMN {
        return Err(ServiceError::argument(format!("range `{}` leaves the sheet grid", range.reference())));
    }
    Ok(range)
}

/// Reads the required shift direction of a range edit.
fn required_direction<'a>(operation: &'a Value, op: &str, first: &'a str, second: &'a str) -> Result<&'a str, ServiceError> {
    let value = operation
        .get("shift")
        .and_then(Value::as_str)
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| ServiceError::argument(format!("`{op}` needs a `shift` of `{first}` or `{second}`")))?;
    if value.eq_ignore_ascii_case(first) {
        return Ok(first);
    }
    if value.eq_ignore_ascii_case(second) {
        return Ok(second);
    }
    Err(ServiceError::argument(format!("`{value}` is not a shift direction; `{op}` needs `{first}` or `{second}`")))
}

/// Line span and across span of a rectangle along an axis.
fn band_span(rectangle: &CellRange, axis: Axis) -> (u32, u32, u32, u32) {
    match axis {
        Axis::Row => (rectangle.start_row, rectangle.end_row, rectangle.start_column, rectangle.end_column),
        Axis::Column => (rectangle.start_column, rectangle.end_column, rectangle.start_row, rectangle.end_row),
    }
}

/// Line and across coordinates of a cell along an axis.
fn line_across(cell: &Cell, axis: Axis) -> (u32, u32) {
    match axis {
        Axis::Row => (cell.row, cell.column),
        Axis::Column => (cell.column, cell.row),
    }
}

/// Rejects a band insert that would push existing content past the grid.
fn check_band_insert(sheet: &Sheet, at: u32, count: u32, across_start: u32, across_end: u32, axis: Axis) -> Result<(), ServiceError> {
    let (max, unit) = match axis {
        Axis::Row => (MAX_ROW, "row"),
        Axis::Column => (MAX_COLUMN, "column"),
    };
    let in_band = |across: u32| (across_start..=across_end).contains(&across);
    let overflows = |line: u32| line >= at && line + count > max;
    let cell_overflows = |cell: &Cell| {
        let (line, across) = line_across(cell, axis);
        in_band(across) && overflows(line)
    };
    let merge_overflows = |merged: &CellRange| {
        let (_, high, across_low, across_high) = band_span(merged, axis);
        across_low <= across_end && across_high >= across_start && overflows(high)
    };
    let merge_crosses = |merged: &CellRange| {
        let (low, high, across_low, across_high) = band_span(merged, axis);
        across_low <= across_end && across_high >= across_start && low < at && high >= at
    };
    if let Some(merged) = sheet.merged.iter().find(|merged| merge_crosses(merged)) {
        return Err(ServiceError::argument(format!(
            "`range.insert` cannot insert through merged range `{}`",
            merged.reference()
        )));
    }
    if sheet.cells.iter().any(cell_overflows) || sheet.merged.iter().any(merge_overflows) {
        return Err(ServiceError::argument(format!(
            "`range.insert` would push content past {unit} {max}"
        )));
    }
    if sheet.dimension.is_some_and(|rectangle| {
        let (_, high, _, _) = band_span(&rectangle, axis);
        overflows(high)
    }) || sheet.auto_filter.is_some_and(|rectangle| {
        let (_, high, _, _) = band_span(&rectangle, axis);
        overflows(high)
    }) {
        return Err(ServiceError::argument(format!(
            "`range.insert` would push the declared range past {unit} {max}"
        )));
    }
    Ok(())
}

/// Grows the declared dimension and auto filter across a band insertion line.
fn grow_band_span(sheet: &mut Sheet, axis: Axis, at: u32, count: u32) {
    for range in [&mut sheet.dimension, &mut sheet.auto_filter].into_iter().flatten() {
        match axis {
            Axis::Row if range.end_row >= at => {
                if range.start_row >= at {
                    range.start_row += count;
                }
                range.end_row += count;
            }
            Axis::Column if range.end_column >= at => {
                if range.start_column >= at {
                    range.start_column += count;
                }
                range.end_column += count;
            }
            _ => {}
        }
    }
}

/// Rejects a swap where a merged range crosses one of the two lines.
fn check_swap_merges(
    merged: &[CellRange],
    axis: Axis,
    first: u32,
    second: u32,
    op: &str,
) -> Result<(), ServiceError> {
    for range in merged {
        let (low, high) = match axis {
            Axis::Row => (range.start_row, range.end_row),
            Axis::Column => (range.start_column, range.end_column),
        };
        let touches = low <= first && first <= high || low <= second && second <= high;
        if touches && low != high {
            return Err(ServiceError::argument(format!(
                "`{op}` cannot move merged range `{}` across the other line",
                range.reference()
            )));
        }
    }
    Ok(())
}

/// Exchanges every cell that sits on one of the two lines.
fn swap_cells(sheet: &mut Sheet, axis: Axis, first: u32, second: u32) {
    let mut from_first: Vec<Cell> = Vec::new();
    let mut from_second: Vec<Cell> = Vec::new();
    sheet.cells.retain(|cell| {
        let line = match axis {
            Axis::Row => cell.row,
            Axis::Column => cell.column,
        };
        if line == first {
            from_first.push(cell.clone());
            false
        }
        else if line == second {
            from_second.push(cell.clone());
            false
        }
        else {
            true
        }
    });
    for mut cell in from_first {
        match axis {
            Axis::Row => cell.row = second,
            Axis::Column => cell.column = second,
        }
        cell.reference = reference::format_reference(cell.row, cell.column);
        sheet.set_cell(cell);
    }
    for mut cell in from_second {
        match axis {
            Axis::Row => cell.row = first,
            Axis::Column => cell.column = first,
        }
        cell.reference = reference::format_reference(cell.row, cell.column);
        sheet.set_cell(cell);
    }
}

/// Moves the row records of the two swapped rows to the other row.
fn swap_row_records(sheet: &mut Sheet, first: u32, second: u32) {
    let at_first = sheet.rows.iter().position(|record| record.index == first);
    let at_second = sheet.rows.iter().position(|record| record.index == second);
    match (at_first, at_second) {
        (Some(a), Some(b)) => {
            sheet.rows[a].index = second;
            sheet.rows[b].index = first;
        }
        (Some(a), None) => sheet.rows[a].index = second,
        (None, Some(b)) => sheet.rows[b].index = first,
        (None, None) => {}
    }
    sheet.rows.sort_by_key(|record| record.index);
}

/// Exchanges the layout of two columns, splitting and merging the width records.
fn swap_column_records(sheet: &mut Sheet, first: u32, second: u32) {
    let mut records: Vec<SheetColumn> = Vec::with_capacity(sheet.columns.len() + 4);
    for record in std::mem::take(&mut sheet.columns) {
        let mut cuts: Vec<u32> =
            [first, first + 1, second, second + 1].into_iter().filter(|cut| *cut > record.min && *cut <= record.max).collect();
        cuts.sort_unstable();
        cuts.dedup();
        let mut start = record.min;
        for cut in cuts {
            records.push(SheetColumn { min: start, max: cut - 1, ..record.clone() });
            start = cut;
        }
        records.push(SheetColumn { min: start, max: record.max, ..record });
    }
    let a = records.iter().position(|record| record.min == first && record.max == first);
    let b = records.iter().position(|record| record.min == second && record.max == second);
    if a.is_none() {
        records.push(SheetColumn { min: first, max: first, width: DEFAULT_WIDTH, custom: false, hidden: false });
    }
    if b.is_none() {
        records.push(SheetColumn { min: second, max: second, width: DEFAULT_WIDTH, custom: false, hidden: false });
    }
    let (a, b) = (
        records.iter().position(|record| record.min == first && record.max == first).expect("first column record"),
        records.iter().position(|record| record.min == second && record.max == second).expect("second column record"),
    );
    let (width, custom, hidden) = (records[b].width, records[b].custom, records[b].hidden);
    records[b].width = records[a].width;
    records[b].custom = records[a].custom;
    records[b].hidden = records[a].hidden;
    records[a].width = width;
    records[a].custom = custom;
    records[a].hidden = hidden;
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

/// Moves merges that sit entirely on one swapped line to the other line.
fn swap_line_merges(sheet: &mut Sheet, axis: Axis, first: u32, second: u32) {
    for range in &mut sheet.merged {
        match axis {
            Axis::Row if range.start_row == range.end_row => {
                if range.start_row == first {
                    range.start_row = second;
                    range.end_row = second;
                }
                else if range.start_row == second {
                    range.start_row = first;
                    range.end_row = first;
                }
            }
            Axis::Column if range.start_column == range.end_column => {
                if range.start_column == first {
                    range.start_column = second;
                    range.end_column = second;
                }
                else if range.start_column == second {
                    range.start_column = first;
                    range.end_column = first;
                }
            }
            _ => {}
        }
    }
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
fn required_column(operation: &Value, key: &str, op: &str) -> Result<u32, ServiceError> {
    let column = match operation.get(key) {
        Some(Value::Number(number)) => number
            .as_u64()
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| ServiceError::argument(format!("`{op}` needs a one-based `{key}` or a column letter")))?,
        Some(Value::String(text)) => reference::column_number(text.trim())
            .ok_or_else(|| ServiceError::argument(format!("`{text}` is not a column letter such as `B`")))?,
        Some(_) => {
            return Err(ServiceError::argument(format!("`{op}` needs a one-based `{key}` or a column letter")));
        }
        None => return Err(ServiceError::argument(format!("`{op}` needs a `{key}`"))),
    };
    if column == 0 || column > MAX_COLUMN {
        return Err(ServiceError::argument(format!("`{op}` needs a `{key}` between 1 and {MAX_COLUMN}")));
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

/// Rejects an insert that would push a named range past the grid on the edited axis.
fn check_insert_names(
    names: &[NamedRange],
    sheet: &str,
    start: u32,
    count: u32,
    max: u32,
    end_of: fn(&CellRange) -> u32,
    unit: &str,
) -> Result<(), ServiceError> {
    let overflows = |value: u32| value >= start && value + count > max;
    if names.iter().any(|named| named.sheet == sheet && overflows(end_of(&named.range))) {
        return Err(ServiceError::argument(format!(
            "inserting {count} {unit}s before {unit} {start} would push a named range past {unit} {max}"
        )));
    }
    Ok(())
}

/// Moves the named ranges on one sheet down by `count`, expanding crossings.
fn shift_named_rows(names: &mut [NamedRange], sheet: &str, start: u32, count: u32) {
    for named in names.iter_mut().filter(|named| named.sheet == sheet) {
        let range = &mut named.range;
        if range.end_row < start {
            continue;
        }
        if range.start_row >= start {
            range.start_row += count;
        }
        range.end_row += count;
    }
}

/// Moves the named ranges on one sheet right by `count`, expanding crossings.
fn shift_named_columns(names: &mut [NamedRange], sheet: &str, start: u32, count: u32) {
    for named in names.iter_mut().filter(|named| named.sheet == sheet) {
        let range = &mut named.range;
        if range.end_column < start {
            continue;
        }
        if range.start_column >= start {
            range.start_column += count;
        }
        range.end_column += count;
    }
}

/// Shrinks or moves named ranges across a deleted row band; a range that
/// falls entirely inside the band loses its name.
fn delete_named_rows(names: &mut Vec<NamedRange>, sheet: &str, start: u32, end: u32) {
    names.retain_mut(|named| {
        if named.sheet != sheet {
            return true;
        }
        match delete_span((named.range.start_row, named.range.end_row), start, end) {
            Some((start_row, end_row)) => {
                named.range.start_row = start_row;
                named.range.end_row = end_row;
                true
            }
            None => false,
        }
    });
}

/// Shrinks or moves named ranges across a deleted column band.
fn delete_named_columns(names: &mut Vec<NamedRange>, sheet: &str, start: u32, end: u32) {
    names.retain_mut(|named| {
        if named.sheet != sheet {
            return true;
        }
        match delete_span((named.range.start_column, named.range.end_column), start, end) {
            Some((start_column, end_column)) => {
                named.range.start_column = start_column;
                named.range.end_column = end_column;
                true
            }
            None => false,
        }
    });
}
