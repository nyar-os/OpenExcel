//! The serializable workbook project owned by OpenExcel.
//!
//! The project is the editable model the CLI, the Node-API surface and the
//! studio all share. It is derived from an imported workbook and can be written
//! back to JSON, CSV and XLSX. Upstream format models are never serialized
//! directly.

use serde::{Deserialize, Serialize};

use crate::reference::format_reference;

/// Format tag of the native project container.
pub const PROJECT_FORMAT: &str = "open-excel/project";

/// Version of the native project container.
pub const PROJECT_VERSION: u32 = 0;

/// One editable workbook project.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Project {
    /// Container tag; always [`PROJECT_FORMAT`].
    pub format: String,
    /// Container version; always [`PROJECT_VERSION`].
    pub version: u32,
    /// Caller-visible workbook title.
    pub title: String,
    /// Style tables kept alongside the cells.
    #[serde(default)]
    pub styles: Styles,
    /// Sheets in workbook order.
    pub sheets: Vec<Sheet>,
    /// Workbook-scope defined names bound to sheet rectangles.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub names: Vec<NamedRange>,
}

impl Project {
    /// Creates an empty project with one visible `Sheet1`.
    pub fn blank(title: impl Into<String>) -> Self {
        Self {
            format: PROJECT_FORMAT.to_owned(),
            version: PROJECT_VERSION,
            title: title.into(),
            styles: Styles::default(),
            sheets: vec![Sheet::new("sheet-1", "Sheet1", 0)],
            names: Vec::new(),
        }
    }

    /// Total number of cells that carry a value or a formula.
    pub fn cell_count(&self) -> usize {
        self.sheets.iter().map(|sheet| sheet.non_empty_cells()).sum()
    }

    /// Finds a sheet by id first, then by case-insensitive name.
    pub fn sheet_index(&self, key: &str) -> Option<usize> {
        self.sheets
            .iter()
            .position(|sheet| sheet.id == key)
            .or_else(|| self.sheets.iter().position(|sheet| sheet.name.eq_ignore_ascii_case(key)))
    }

    /// Borrows one sheet by id or name.
    pub fn sheet(&self, key: &str) -> Option<&Sheet> {
        self.sheet_index(key).map(|index| &self.sheets[index])
    }

    /// Mutably borrows one sheet by id or name.
    pub fn sheet_mut(&mut self, key: &str) -> Option<&mut Sheet> {
        self.sheet_index(key).map(move |index| &mut self.sheets[index])
    }

    /// First visible sheet, falling back to the first sheet.
    pub fn primary_sheet(&self) -> Option<&Sheet> {
        self.sheets.iter().find(|sheet| sheet.visibility == SheetVisibility::Visible).or_else(|| self.sheets.first())
    }

    /// Allocates an id that is not used by any sheet yet.
    pub fn allocate_sheet_id(&self) -> String {
        let mut next = self.sheets.len() + 1;
        loop {
            let candidate = format!("sheet-{next}");
            if self.sheets.iter().all(|sheet| sheet.id != candidate) {
                return candidate;
            }
            next += 1;
        }
    }

    /// Rewrites every sheet index so the position in the vector is the order.
    pub fn reindex(&mut self) {
        for (index, sheet) in self.sheets.iter_mut().enumerate() {
            sheet.index = index as u32;
        }
    }

    /// Normalizes references, ordering and indexes after an untrusted deserialize.
    pub fn normalize(&mut self) {
        if self.format != PROJECT_FORMAT {
            self.format = PROJECT_FORMAT.to_owned();
        }
        if self.version != PROJECT_VERSION {
            self.version = PROJECT_VERSION;
        }
        for sheet in &mut self.sheets {
            sheet.normalize();
        }
        self.names.retain(|named| self.sheets.iter().any(|sheet| sheet.id == named.sheet));
        self.names.sort_by(|left, right| left.name.to_lowercase().cmp(&right.name.to_lowercase()));
        self.names.dedup_by(|left, right| left.name.eq_ignore_ascii_case(&right.name));
        self.reindex();
    }
}

/// Workbook-scope defined name bound to one sheet rectangle.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NamedRange {
    /// Defined name as the caller typed it.
    pub name: String,
    /// Id of the sheet the range lives on.
    pub sheet: String,
    /// Rectangle the name resolves to.
    pub range: CellRange,
}

/// Style tables kept with the project.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Styles {
    /// Number format codes seen in the workbook.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub number_formats: Vec<String>,
    /// Distinct cell formats seen in the workbook.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cell_formats: Vec<CellStyle>,
}

/// One sheet of the project.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Sheet {
    /// Stable sheet identity; survives rename.
    pub id: String,
    /// Sheet tab name.
    pub name: String,
    /// Zero-based workbook order.
    pub index: u32,
    /// Visibility state.
    pub visibility: SheetVisibility,
    /// Dimension declared by the source writer, when it declared one.
    #[serde(default)]
    pub dimension: Option<CellRange>,
    /// Column metadata.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<SheetColumn>,
    /// Row metadata.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub rows: Vec<SheetRow>,
    /// Sparse cells sorted by row and column.
    #[serde(default)]
    pub cells: Vec<Cell>,
    /// Merged ranges.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub merged: Vec<CellRange>,
    /// Frozen pane.
    #[serde(default)]
    pub freeze: Option<Freeze>,
    /// Range the sheet's autofilter covers, when one is declared.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<CellRange>,
}

impl Sheet {
    /// Creates an empty visible sheet.
    pub fn new(id: impl Into<String>, name: impl Into<String>, index: u32) -> Self {
        Self {
            id: id.into(),
            name: name.into(),
            index,
            visibility: SheetVisibility::Visible,
            dimension: None,
            columns: Vec::new(),
            rows: Vec::new(),
            cells: Vec::new(),
            merged: Vec::new(),
            freeze: None,
            auto_filter: None,
        }
    }

    /// Number of cells that carry a value or a formula.
    pub fn non_empty_cells(&self) -> usize {
        self.cells.iter().filter(|cell| !cell.is_empty()).count()
    }

    /// Finds one cell by a one-based row and column.
    pub fn cell(&self, row: u32, column: u32) -> Option<&Cell> {
        self.cells.binary_search_by_key(&(row, column), |cell| (cell.row, cell.column)).ok().map(|index| &self.cells[index])
    }

    /// Inserts or replaces one cell, keeping the sparse vector sorted.
    pub fn set_cell(&mut self, cell: Cell) {
        match self.cells.binary_search_by_key(&(cell.row, cell.column), |entry| (entry.row, entry.column)) {
            Ok(index) => {
                if cell.is_blank() {
                    self.cells.remove(index);
                }
                else {
                    self.cells[index] = cell;
                }
            }
            Err(index) => {
                if !cell.is_blank() {
                    self.cells.insert(index, cell);
                }
            }
        }
    }

    /// Writes a value and optional formula at a one-based row and column.
    pub fn write_cell(&mut self, row: u32, column: u32, value: CellValue, formula: Option<String>, cached: Option<CellValue>) {
        let cell = Cell {
            reference: format_reference(row, column),
            row,
            column,
            value,
            formula,
            cached,
            style: None,
            computed: None,
            computed_state: None,
        };
        self.set_cell(cell);
    }

    /// Removes every cell that falls inside the range.
    pub fn clear_range(&mut self, range: &CellRange) {
        self.cells.retain(|cell| !range.contains(cell.row, cell.column));
    }

    /// Last row that carries content or a merge boundary.
    pub fn last_row(&self) -> u32 {
        self.cells
            .iter()
            .filter(|cell| !cell.is_empty())
            .map(|cell| cell.row)
            .chain(self.merged.iter().map(|range| range.end_row))
            .max()
            .unwrap_or(0)
    }

    /// Last column that carries content or a merge boundary.
    pub fn last_column(&self) -> u32 {
        self.cells
            .iter()
            .filter(|cell| !cell.is_empty())
            .map(|cell| cell.column)
            .chain(self.merged.iter().map(|range| range.end_column))
            .max()
            .unwrap_or(0)
    }

    /// Rectangle covered by content and merges, when the sheet is not empty.
    pub fn actual_range(&self) -> Option<CellRange> {
        let last_row = self.last_row();
        let last_column = self.last_column();
        if last_row == 0 && last_column == 0 {
            return None;
        }
        Some(CellRange::between(1, 1, last_row.max(1), last_column.max(1)))
    }

    /// Rectangle an inspector should use: the declared dimension, else the actual cells.
    pub fn bounds(&self) -> Option<CellRange> {
        self.dimension.or_else(|| self.actual_range())
    }

    /// Width of a column in Excel character units, falling back to the default.
    pub fn column_width(&self, column: u32) -> f64 {
        self.columns
            .iter()
            .rev()
            .find(|record| record.min <= column && column <= record.max)
            .map(|record| record.width)
            .unwrap_or(8.43)
    }

    /// Height of a row in points, falling back to the default.
    pub fn row_height(&self, row: u32) -> f64 {
        self.rows.iter().rev().find(|record| record.index == row).and_then(|record| record.height).unwrap_or(15.0)
    }

    /// Merge range covering the cell, when the cell sits inside a merge.
    pub fn merge_covering(&self, row: u32, column: u32) -> Option<CellRange> {
        self.merged.iter().copied().find(|range| {
            range.start_row <= row && row <= range.end_row && range.start_column <= column && column <= range.end_column
        })
    }

    /// Rewrites references, sorts cells, drops duplicates and fixes the index.
    pub fn normalize(&mut self) {
        for cell in &mut self.cells {
            cell.reference = format_reference(cell.row, cell.column);
        }
        self.cells.sort_by_key(|cell| (cell.row, cell.column));
        self.cells.dedup_by_key(|cell| (cell.row, cell.column));
        self.cells.retain(|cell| !cell.is_blank());
    }
}

/// Sheet visibility recorded by the workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum SheetVisibility {
    /// Normal visible sheet.
    Visible,
    /// Hidden by the author but still part of the workbook.
    Hidden,
    /// Hidden and excluded from the sheet tab context menu.
    VeryHidden,
}

impl SheetVisibility {
    /// Whether the sheet is shown as a tab.
    pub fn is_visible(self) -> bool {
        self == Self::Visible
    }
}

/// One cell of a sheet.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Cell {
    /// A1 reference such as `B4`.
    pub reference: String,
    /// One-based row index.
    pub row: u32,
    /// One-based column index.
    pub column: u32,
    /// Stored value.
    pub value: CellValue,
    /// Formula source with the leading `=` when the cell is a formula cell.
    #[serde(default)]
    pub formula: Option<String>,
    /// Cached formula result of the source writer.
    #[serde(default)]
    pub cached: Option<CellValue>,
    /// Resolved visual style.
    #[serde(default)]
    pub style: Option<CellStyle>,
    /// Result of the most recent evaluation of `formula`.
    ///
    /// Projects written before the evaluator existed simply omit the field.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub computed: Option<CellValue>,
    /// How the most recent evaluation of `formula` ended.
    ///
    /// Non-formula cells omit the field, and projects written before it existed
    /// simply omit it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub computed_state: Option<ComputedState>,
}

/// Outcome of evaluating one formula cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum ComputedState {
    /// The engine evaluated every function in the formula.
    Computed,
    /// The formula uses a function or construct this engine does not evaluate.
    ///
    /// The imported cached value, when present, stays the effective value.
    Unsupported,
}

impl Cell {
    /// Whether the cell carries neither a value nor a formula.
    pub fn is_empty(&self) -> bool {
        matches!(self.value, CellValue::Empty) && self.formula.is_none()
    }

    /// Whether the cell carries nothing at all, not even a visual style.
    pub fn is_blank(&self) -> bool {
        self.is_empty() && self.style.as_ref().is_none_or(CellStyle::is_empty)
    }
}

/// Resolved value of one cell.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "kebab-case")]
pub enum CellValue {
    /// Text value.
    Text {
        /// Cell text.
        text: String,
    },
    /// Numeric value.
    Number {
        /// Numeric value.
        number: f64,
    },
    /// Boolean value.
    Bool {
        /// Boolean value.
        value: bool,
    },
    /// Spreadsheet error literal such as `#N/A`.
    Error {
        /// Error literal.
        code: String,
    },
    /// No value.
    #[serde(rename = "empty")]
    Empty,
}

impl CellValue {
    /// Builds a text value.
    pub fn text(text: impl Into<String>) -> Self {
        Self::Text { text: text.into() }
    }

    /// Builds a numeric value.
    pub fn number(number: f64) -> Self {
        Self::Number { number }
    }

    /// Builds a boolean value.
    pub fn boolean(value: bool) -> Self {
        Self::Bool { value }
    }

    /// Builds an error literal.
    pub fn error(code: impl Into<String>) -> Self {
        Self::Error { code: code.into() }
    }

    /// Whether the value carries no data.
    pub fn is_empty(&self) -> bool {
        matches!(self, Self::Empty)
    }

    /// Display text used by CSV and SVG projections.
    pub fn display_text(&self) -> String {
        match self {
            Self::Text { text } => text.clone(),
            Self::Number { number } => format_number(*number),
            Self::Bool { value } => {
                if *value {
                    "TRUE".to_owned()
                }
                else {
                    "FALSE".to_owned()
                }
            }
            Self::Error { code } => code.clone(),
            Self::Empty => String::new(),
        }
    }
}

/// Formats a number without a trailing `.0` for whole values.
pub fn format_number(number: f64) -> String {
    if !number.is_finite() {
        return number.to_string();
    }
    if number.fract() == 0.0 && number.abs() < 1.0e15 { format!("{number:.0}") } else { number.to_string() }
}

/// Visual style subset kept by the project.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", default)]
pub struct CellStyle {
    /// Bold font.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    /// Italic font.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    /// Font size in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f32>,
    /// Font color as an ARGB or `#RRGGBB` string.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    /// Solid fill color.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    /// Number format code.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    /// Horizontal alignment: `left`, `center`, `right`, `justify`, `distribute` or `fill`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    /// Whether the cell wraps text.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap: Option<bool>,
}

impl CellStyle {
    /// Whether the style carries no visible property.
    pub fn is_empty(&self) -> bool {
        self == &Self::default()
    }

    /// Stable deduplication key.
    pub fn key(&self) -> String {
        format!(
            "b{:?}i{:?}s{:?}c{:?}f{:?}n{:?}a{:?}w{:?}",
            self.bold, self.italic, self.font_size, self.font_color, self.fill, self.number_format, self.align, self.wrap
        )
    }
}

/// Inclusive cell rectangle.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellRange {
    /// One-based first row.
    pub start_row: u32,
    /// One-based last row.
    pub end_row: u32,
    /// One-based first column.
    pub start_column: u32,
    /// One-based last column.
    pub end_column: u32,
}

impl CellRange {
    /// Builds a rectangle from two corners, swapping them when needed.
    pub fn between(start_row: u32, start_column: u32, end_row: u32, end_column: u32) -> Self {
        Self {
            start_row: start_row.min(end_row),
            end_row: start_row.max(end_row),
            start_column: start_column.min(end_column),
            end_column: start_column.max(end_column),
        }
    }

    /// Number of rows covered by the range.
    pub fn rows(&self) -> u32 {
        self.end_row.saturating_sub(self.start_row) + 1
    }

    /// Number of columns covered by the range.
    pub fn columns(&self) -> u32 {
        self.end_column.saturating_sub(self.start_column) + 1
    }

    /// Whether a one-based cell sits inside the range.
    pub fn contains(&self, row: u32, column: u32) -> bool {
        self.start_row <= row && row <= self.end_row && self.start_column <= column && column <= self.end_column
    }

    /// Whether two rectangles share at least one cell.
    pub fn overlaps(&self, other: &Self) -> bool {
        self.start_row <= other.end_row
            && other.start_row <= self.end_row
            && self.start_column <= other.end_column
            && other.start_column <= self.end_column
    }

    /// Whether the rectangle has a positive extent.
    pub fn is_valid(&self) -> bool {
        self.start_row >= 1 && self.start_column >= 1 && self.start_row <= self.end_row && self.start_column <= self.end_column
    }

    /// A1 text such as `A1:C4`.
    pub fn reference(&self) -> String {
        format!("{}:{}", format_reference(self.start_row, self.start_column), format_reference(self.end_row, self.end_column))
    }
}

/// Column metadata from the source writer.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetColumn {
    /// First column covered by this record.
    pub min: u32,
    /// Last column covered by this record.
    pub max: u32,
    /// Width in Excel character units.
    pub width: f64,
    /// Whether the author set an explicit width.
    #[serde(default)]
    pub custom: bool,
    /// Whether the column is hidden.
    #[serde(default)]
    pub hidden: bool,
}

/// Row metadata from the source writer.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetRow {
    /// One-based row index.
    pub index: u32,
    /// Height in points.
    #[serde(default)]
    pub height: Option<f64>,
    /// Whether the author set an explicit height.
    #[serde(default)]
    pub custom_height: bool,
    /// Whether the row is hidden.
    #[serde(default)]
    pub hidden: bool,
}

/// Frozen pane state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Freeze {
    /// Number of frozen columns.
    pub columns: u32,
    /// Number of frozen rows.
    pub rows: u32,
}
