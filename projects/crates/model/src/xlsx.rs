//! XLSX import and export.
//!
//! Reading goes through the pinned upstream XLSX reader and projects the format
//! model into [`Project`]. Writing maps the same project back onto the
//! format-native workbook and delegates the package assembly to the pinned
//! writer, so styles, merges, column widths and frozen panes have one
//! implementation.

use document_formats::{
    export::xlsx::export_xlsx_bytes,
    import::xlsx::{
        import_xlsx_bytes,
        model::{
            CellFormula, CellRange as XlsxRange, CellStyle as XlsxStyle, CellValue as XlsxValue, FormulaKind, FreezePane,
            HorizontalAlign, SheetColumn as XlsxColumn, SheetRow as XlsxRow, SheetVisibility as XlsxVisibility, VerticalAlign,
            XlsxCell, XlsxSheet, XlsxWorkbook,
        },
    },
};

use crate::{
    error::ServiceError,
    model::{
        Cell, CellRange, CellStyle, CellValue, ComputedState, Freeze, Project, Sheet, SheetColumn, SheetRow, SheetVisibility,
        Styles,
    },
};

/// Reads an XLSX/XLSM package into the project model.
pub fn read_project(label: &str, bytes: &[u8], title: impl Into<String>) -> Result<Project, ServiceError> {
    let workbook = import_xlsx_bytes(label, bytes).map_err(|error| ServiceError::invalid("xlsx", error.to_string()))?;
    Ok(project_from_workbook(workbook, title))
}

/// Projects the format-native workbook into the editable project model.
pub fn project_from_workbook(workbook: XlsxWorkbook, title: impl Into<String>) -> Project {
    let sheets = workbook
        .sheets
        .into_iter()
        .enumerate()
        .map(|(index, sheet)| {
            let id = format!("sheet-{}", index + 1);
            let mut projected = Sheet::new(id, sheet.name, sheet.index);
            projected.visibility = match sheet.visibility {
                XlsxVisibility::Visible => SheetVisibility::Visible,
                XlsxVisibility::Hidden => SheetVisibility::Hidden,
                XlsxVisibility::VeryHidden => SheetVisibility::VeryHidden,
            };
            projected.dimension = sheet.dimension.map(range_from);
            projected.columns = sheet
                .columns
                .into_iter()
                .map(|column| SheetColumn {
                    min: column.min,
                    max: column.max,
                    width: column.width,
                    custom: column.custom,
                    hidden: column.hidden,
                })
                .collect();
            projected.rows = sheet
                .rows
                .into_iter()
                .map(|row| SheetRow {
                    index: row.index,
                    height: row.height,
                    custom_height: row.custom_height,
                    hidden: row.hidden,
                })
                .collect();
            projected.merged = sheet.merged.into_iter().map(range_from).collect();
            projected.freeze = sheet.freeze.map(|freeze| Freeze { columns: freeze.columns, rows: freeze.rows });
            projected.auto_filter = sheet.auto_filter.map(range_from);
            projected.cells = sheet
                .cells
                .into_iter()
                .filter_map(|cell| {
                    let style = style_from_xlsx(&cell.style);
                    let formula = cell.formula.as_ref().map(|formula| normalize_formula(&formula.source));
                    let cached = cell.formula.as_ref().and_then(|formula| formula.cached.as_deref()).map(value_from_xlsx);
                    // A formula cell keeps its result in `cached`/`computed`, never
                    // as a literal, so a save can never freeze a stale value.
                    let value = if formula.is_some() { CellValue::Empty } else { value_from_xlsx(&cell.value) };
                    if value.is_empty() && formula.is_none() && style.is_none() {
                        return None;
                    }
                    Some(Cell {
                        reference: cell.reference,
                        row: cell.row,
                        column: cell.column,
                        value,
                        formula,
                        cached,
                        style,
                        computed: None,
                        computed_state: None,
                    })
                })
                .collect();
            projected.normalize();
            projected
        })
        .collect();

    let mut number_formats: Vec<String> = workbook.styles.number_formats.values().cloned().collect();
    number_formats.sort();
    number_formats.dedup();
    let mut cell_formats: Vec<CellStyle> = Vec::new();
    for style in workbook.styles.cell_formats {
        if let Some(projected) = style_from_xlsx(&style)
            && !cell_formats.iter().any(|existing| existing.key() == projected.key())
        {
            cell_formats.push(projected);
        }
    }

    let mut project = Project {
        format: crate::model::PROJECT_FORMAT.to_owned(),
        version: crate::model::PROJECT_VERSION,
        title: title.into(),
        styles: Styles { number_formats, cell_formats },
        sheets,
    };
    project.reindex();
    project
}

fn range_from(range: document_formats::import::xlsx::model::CellRange) -> CellRange {
    CellRange {
        start_row: range.start_row,
        end_row: range.end_row,
        start_column: range.start_column,
        end_column: range.end_column,
    }
}

fn value_from_xlsx(value: &XlsxValue) -> CellValue {
    match value {
        XlsxValue::Empty => CellValue::Empty,
        XlsxValue::Number(number) => CellValue::number(*number),
        XlsxValue::Date(date) => CellValue::number(date.serial),
        XlsxValue::Boolean(value) => CellValue::boolean(*value),
        XlsxValue::Text(text) => CellValue::text(text),
        XlsxValue::SharedString { text, .. } => CellValue::text(text),
        XlsxValue::Error { code } => CellValue::error(code),
    }
}

fn style_from_xlsx(style: &XlsxStyle) -> Option<CellStyle> {
    let projected = CellStyle {
        bold: style.bold.then_some(true),
        italic: style.italic.then_some(true),
        font_size: style.font_size,
        font_color: style.font_color.clone(),
        fill: style.fill_color.clone(),
        number_format: style.number_format.clone(),
        align: match style.horizontal {
            HorizontalAlign::General => None,
            HorizontalAlign::Left => Some("left".to_owned()),
            HorizontalAlign::Center => Some("center".to_owned()),
            HorizontalAlign::Right => Some("right".to_owned()),
            HorizontalAlign::Justify => Some("justify".to_owned()),
            HorizontalAlign::Distributed => Some("distribute".to_owned()),
            HorizontalAlign::Fill => Some("fill".to_owned()),
        },
        wrap: style.wrap_text.then_some(true),
    };
    (!projected.is_empty()).then_some(projected)
}

fn normalize_formula(source: &str) -> String {
    let trimmed = source.trim();
    if trimmed.starts_with('=') { trimmed.to_owned() } else { format!("={trimmed}") }
}

/// Writes the project as an XLSX package.
pub fn write_package(project: &Project) -> Result<Vec<u8>, ServiceError> {
    let workbook = workbook_from_project(project);
    export_xlsx_bytes(&workbook).map_err(|error| ServiceError::io(format!("failed to assemble the XLSX package: {error}")))
}

/// Maps the editable project back onto the format-native workbook.
fn workbook_from_project(project: &Project) -> XlsxWorkbook {
    let mut workbook = XlsxWorkbook::new(&project.title);
    workbook.sheets = project
        .sheets
        .iter()
        .map(|sheet| {
            let mut written = XlsxSheet::new(&sheet.name, sheet.index);
            written.visibility = match sheet.visibility {
                SheetVisibility::Visible => XlsxVisibility::Visible,
                SheetVisibility::Hidden => XlsxVisibility::Hidden,
                SheetVisibility::VeryHidden => XlsxVisibility::VeryHidden,
            };
            written.dimension = sheet.dimension.as_ref().map(range_to);
            written.columns = sheet
                .columns
                .iter()
                .map(|column| XlsxColumn {
                    min: column.min,
                    max: column.max,
                    width: column.width,
                    custom: column.custom,
                    hidden: column.hidden,
                })
                .collect();
            written.rows = sheet
                .rows
                .iter()
                .map(|row| XlsxRow {
                    index: row.index,
                    height: row.height,
                    custom_height: row.custom_height,
                    hidden: row.hidden,
                })
                .collect();
            written.merged = sheet.merged.iter().map(range_to).collect();
            written.freeze = sheet.freeze.map(|freeze| FreezePane { columns: freeze.columns, rows: freeze.rows });
            written.auto_filter = sheet.auto_filter.as_ref().map(range_to);
            written.cells = sheet
                .cells
                .iter()
                .filter_map(|cell| {
                    if cell.is_empty() && cell.style.is_none() {
                        return None;
                    }
                    let formula = cell.formula.as_ref().map(|source| {
                        let cached = if cell.computed_state == Some(ComputedState::Unsupported) {
                            cell.cached.clone()
                        }
                        else {
                            cell.computed
                                .clone()
                                .or_else(|| cell.cached.clone())
                                .or_else(|| (!cell.value.is_empty()).then(|| cell.value.clone()))
                        };
                        CellFormula {
                            source: source.trim().trim_start_matches('=').to_owned(),
                            kind: FormulaKind::Normal,
                            cached: cached.map(|value| Box::new(value_to_xlsx(&value))),
                        }
                    });
                    Some(XlsxCell {
                        reference: cell.reference.clone(),
                        row: cell.row,
                        column: cell.column,
                        value: value_to_xlsx(&cell.value),
                        formula,
                        style_index: None,
                        style: style_to_xlsx(cell.style.as_ref()),
                    })
                })
                .collect();
            written
        })
        .collect();
    workbook
}

fn range_to(range: &CellRange) -> XlsxRange {
    XlsxRange {
        start_row: range.start_row,
        end_row: range.end_row,
        start_column: range.start_column,
        end_column: range.end_column,
    }
}

fn value_to_xlsx(value: &CellValue) -> XlsxValue {
    match value {
        CellValue::Empty => XlsxValue::Empty,
        CellValue::Number { number } => XlsxValue::Number(*number),
        CellValue::Text { text } => XlsxValue::Text(text.clone()),
        CellValue::Bool { value } => XlsxValue::Boolean(*value),
        CellValue::Error { code } => XlsxValue::Error { code: code.clone() },
    }
}

fn style_to_xlsx(style: Option<&CellStyle>) -> XlsxStyle {
    let Some(style) = style
    else {
        return XlsxStyle::default();
    };
    XlsxStyle {
        bold: style.bold.unwrap_or(false),
        italic: style.italic.unwrap_or(false),
        font_size: style.font_size,
        font_color: style.font_color.clone(),
        fill_color: style.fill.clone(),
        number_format: style.number_format.clone(),
        horizontal: align_to(style.align.as_deref()),
        wrap_text: style.wrap.unwrap_or(false),
        vertical: VerticalAlign::Bottom,
        ..XlsxStyle::default()
    }
}

fn align_to(align: Option<&str>) -> HorizontalAlign {
    match align {
        Some("left") => HorizontalAlign::Left,
        Some("center") => HorizontalAlign::Center,
        Some("right") => HorizontalAlign::Right,
        Some("justify") => HorizontalAlign::Justify,
        Some("distribute") => HorizontalAlign::Distributed,
        Some("fill") => HorizontalAlign::Fill,
        _ => HorizontalAlign::General,
    }
}
