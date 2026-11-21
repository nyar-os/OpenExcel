//! XLSX import and export.
//!
//! Reading goes through the pinned upstream XLSX reader and projects the format
//! model into [`Project`]. Writing emits a minimal SpreadsheetML package with
//! inline strings, styles, merges, column widths and frozen panes, and is then
//! read back by the same importer in tests.

use data_binary_zip::{ZipEntry, ZipWriter};
use document_formats::import::xlsx::{
    import_xlsx_bytes,
    model::{CellStyle as XlsxStyle, CellValue as XlsxValue, HorizontalAlign, SheetVisibility as XlsxVisibility, XlsxWorkbook},
};

use crate::{
    error::ServiceError,
    model::{Cell, CellRange, CellStyle, CellValue, Freeze, Project, Sheet, SheetColumn, SheetRow, SheetVisibility, Styles},
    reference::format_reference,
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
            projected.cells = sheet
                .cells
                .into_iter()
                .filter_map(|cell| {
                    let style = style_from_xlsx(&cell.style);
                    let formula = cell.formula.as_ref().map(|formula| normalize_formula(&formula.source));
                    let cached = cell.formula.as_ref().and_then(|formula| formula.cached.as_deref()).map(value_from_xlsx);
                    let value = value_from_xlsx(&cell.value);
                    if value.is_empty() && formula.is_none() && style.is_none() {
                        return None;
                    }
                    Some(Cell { reference: cell.reference, row: cell.row, column: cell.column, value, formula, cached, style })
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
    let table = StyleTable::build(project);
    let mut writer = ZipWriter::new();

    writer.push(ZipEntry::deflated("[Content_Types].xml", content_types(project)));
    writer.push(ZipEntry::deflated("_rels/.rels", root_rels()));
    writer.push(ZipEntry::deflated("xl/workbook.xml", workbook_xml(project)));
    writer.push(ZipEntry::deflated("xl/_rels/workbook.xml.rels", workbook_rels(project)));
    for (index, sheet) in project.sheets.iter().enumerate() {
        writer.push(ZipEntry::deflated(format!("xl/worksheets/sheet{}.xml", index + 1), sheet_xml(sheet, &table)));
    }
    writer.push(ZipEntry::deflated("xl/styles.xml", table.xml()));

    writer.into_bytes().map_err(|error| ServiceError::io(format!("failed to assemble the XLSX package: {error}")))
}

fn content_types(project: &Project) -> String {
    let mut body = String::new();
    body.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    body.push_str("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
    body.push_str("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
    body.push_str("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
    body.push_str("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
    body.push_str("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
    for index in 0..project.sheets.len() {
        body.push_str(&format!(
            "<Override PartName=\"/xl/worksheets/sheet{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
            index + 1
        ));
    }
    body.push_str("</Types>");
    body
}

fn root_rels() -> String {
    concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>",
        "</Relationships>"
    )
    .to_owned()
}

fn workbook_xml(project: &Project) -> String {
    let mut body = String::new();
    body.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    body.push_str("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
    body.push_str("<bookViews><workbookView activeTab=\"0\"/></bookViews>");
    body.push_str("<sheets>");
    for (index, sheet) in project.sheets.iter().enumerate() {
        let state = match sheet.visibility {
            SheetVisibility::Visible => String::new(),
            SheetVisibility::Hidden => " state=\"hidden\"".to_owned(),
            SheetVisibility::VeryHidden => " state=\"veryHidden\"".to_owned(),
        };
        body.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\"{} r:id=\"rId{}\"/>",
            escape_attribute(&sheet.name),
            index + 1,
            state,
            index + 1
        ));
    }
    body.push_str("</sheets></workbook>");
    body
}

fn workbook_rels(project: &Project) -> String {
    let mut body = String::new();
    body.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    body.push_str("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    for index in 0..project.sheets.len() {
        body.push_str(&format!(
            "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.xml\"/>",
            index + 1,
            index + 1
        ));
    }
    body.push_str(&format!(
        "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>",
        project.sheets.len() + 1
    ));
    body.push_str("</Relationships>");
    body
}

fn sheet_xml(sheet: &Sheet, table: &StyleTable) -> String {
    let mut body = String::new();
    body.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    body.push_str("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
    if let Some(bounds) = sheet.bounds() {
        body.push_str(&format!("<dimension ref=\"{}\"/>", bounds.reference()));
    }
    body.push_str("<sheetViews><sheetView workbookViewId=\"0\">");
    if let Some(freeze) = sheet.freeze.filter(|freeze| freeze.columns > 0 || freeze.rows > 0) {
        let top_left = format_reference(freeze.rows + 1, freeze.columns + 1);
        body.push_str(&format!(
            "<pane xSplit=\"{}\" ySplit=\"{}\" topLeftCell=\"{}\" activePane=\"bottomRight\" state=\"frozen\"/>",
            freeze.columns, freeze.rows, top_left
        ));
    }
    body.push_str("</sheetView></sheetViews>");
    body.push_str("<sheetFormatPr defaultRowHeight=\"15\"/>");
    if !sheet.columns.is_empty() {
        body.push_str("<cols>");
        for column in &sheet.columns {
            body.push_str(&format!(
                "<col min=\"{}\" max=\"{}\" width=\"{}\"{}{}/>",
                column.min,
                column.max,
                format_number(column.width),
                if column.custom { " customWidth=\"1\"" } else { "" },
                if column.hidden { " hidden=\"1\"" } else { "" }
            ));
        }
        body.push_str("</cols>");
    }

    body.push_str("<sheetData>");
    let mut row = 0u32;
    let mut index = 0usize;
    while index < sheet.cells.len() {
        let cell = &sheet.cells[index];
        if cell.row != row {
            if row != 0 {
                body.push_str("</row>");
            }
            row = cell.row;
            let metadata = sheet.rows.iter().find(|record| record.index == row);
            let mut start = format!("<row r=\"{row}\"");
            if let Some(metadata) = metadata {
                if let Some(height) = metadata.height {
                    start.push_str(&format!(" ht=\"{}\"", format_number(height)));
                    if metadata.custom_height {
                        start.push_str(" customHeight=\"1\"");
                    }
                }
                if metadata.hidden {
                    start.push_str(" hidden=\"1\"");
                }
            }
            start.push('>');
            body.push_str(&start);
        }
        body.push_str(&cell_xml(cell, table));
        index += 1;
    }
    if row != 0 {
        body.push_str("</row>");
    }
    body.push_str("</sheetData>");

    if !sheet.merged.is_empty() {
        body.push_str(&format!("<mergeCells count=\"{}\">", sheet.merged.len()));
        for range in &sheet.merged {
            body.push_str(&format!("<mergeCell ref=\"{}\"/>", range.reference()));
        }
        body.push_str("</mergeCells>");
    }
    body.push_str("</worksheet>");
    body
}

fn cell_xml(cell: &Cell, table: &StyleTable) -> String {
    if cell.is_empty() {
        return String::new();
    }
    let style = table.xf_index(cell.style.as_ref());
    let style_attribute = match style {
        Some(index) => format!(" s=\"{index}\""),
        None => String::new(),
    };
    let reference = escape_attribute(&cell.reference);
    let cached = cell.cached.clone().or_else(|| (!cell.value.is_empty()).then(|| cell.value.clone()));

    let mut body = format!("<c r=\"{reference}\"{style_attribute}");
    match &cell.formula {
        Some(formula) => {
            let source = formula.trim().strip_prefix('=').unwrap_or(formula.trim());
            let (kind, value) = match &cached {
                Some(CellValue::Text { text }) => (" t=\"str\"", Some(text.clone())),
                Some(CellValue::Bool { value }) => (" t=\"b\"", Some(if *value { "1".to_owned() } else { "0".to_owned() })),
                Some(CellValue::Error { code }) => (" t=\"e\"", Some(code.clone())),
                Some(CellValue::Number { number }) => ("", Some(format_number(*number))),
                _ => ("", None),
            };
            if let Some(value) = value {
                body.push_str(kind);
                body.push('>');
                body.push_str(&format!("<f>{}</f><v>{}</v>", escape_text(source), escape_text(&value)));
            }
            else {
                body.push_str(kind);
                body.push('>');
                body.push_str(&format!("<f>{}</f>", escape_text(source)));
            }
        }
        None => match &cell.value {
            CellValue::Empty => return String::new(),
            CellValue::Text { text } => {
                body.push_str(" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
                body.push_str(&escape_text(text));
                body.push_str("</t></is>");
            }
            CellValue::Number { number } => {
                body.push('>');
                body.push_str(&format!("<v>{}</v>", format_number(*number)));
            }
            CellValue::Bool { value } => {
                body.push_str(" t=\"b\"><v>");
                body.push_str(if *value { "1" } else { "0" });
                body.push_str("</v>");
            }
            CellValue::Error { code } => {
                body.push_str(" t=\"e\"><v>");
                body.push_str(&escape_text(code));
                body.push_str("</v>");
            }
        },
    }
    body.push_str("</c>");
    body
}

/// Style table of one written package.
struct StyleTable {
    numbers: Vec<(u32, String)>,
    fonts: Vec<FontSpec>,
    fills: Vec<String>,
    xfs: Vec<Xf>,
}

#[derive(Debug, Clone, PartialEq, Default)]
struct FontSpec {
    bold: bool,
    italic: bool,
    size: Option<f32>,
    color: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Default)]
struct Xf {
    font: usize,
    fill: usize,
    num_fmt: u32,
    align: Option<String>,
    wrap: bool,
}

impl StyleTable {
    fn build(project: &Project) -> Self {
        let mut table =
            Self { numbers: Vec::new(), fonts: vec![FontSpec::default()], fills: Vec::new(), xfs: vec![Xf::default()] };
        for sheet in &project.sheets {
            for cell in &sheet.cells {
                if let Some(style) = &cell.style {
                    table.register(style);
                }
            }
        }
        table
    }

    /// Registers one style and returns its cellXfs index.
    fn register(&mut self, style: &CellStyle) -> usize {
        if style.is_empty() {
            return 0;
        }
        let font = FontSpec {
            bold: style.bold.unwrap_or(false),
            italic: style.italic.unwrap_or(false),
            size: style.font_size,
            color: style.font_color.clone(),
        };
        let font = match self.fonts.iter().position(|existing| existing == &font) {
            Some(index) => index,
            None => {
                self.fonts.push(font);
                self.fonts.len() - 1
            }
        };
        let fill = match &style.fill {
            Some(color) => match self.fills.iter().position(|existing| existing == color) {
                Some(index) => index + 2,
                None => {
                    self.fills.push(color.clone());
                    self.fills.len() + 1
                }
            },
            None => 0,
        };
        let num_fmt = match style.number_format.as_deref() {
            Some(code) if !code.eq_ignore_ascii_case("general") && !code.trim().is_empty() => {
                match self.numbers.iter().find(|(_, existing)| existing == code) {
                    Some((id, _)) => *id,
                    None => {
                        let id = 164 + self.numbers.len() as u32;
                        self.numbers.push((id, code.to_owned()));
                        id
                    }
                }
            }
            _ => 0,
        };
        let xf = Xf { font, fill, num_fmt, align: valid_align(style.align.as_deref()), wrap: style.wrap.unwrap_or(false) };
        match self.xfs.iter().position(|existing| existing == &xf) {
            Some(index) => index,
            None => {
                self.xfs.push(xf);
                self.xfs.len() - 1
            }
        }
    }

    /// Looks up the cellXfs index of a style already registered during `build`.
    fn xf_index(&self, style: Option<&CellStyle>) -> Option<usize> {
        let style = style?;
        if style.is_empty() {
            return None;
        }
        let font = self.fonts.iter().position(|existing| {
            *existing
                == FontSpec {
                    bold: style.bold.unwrap_or(false),
                    italic: style.italic.unwrap_or(false),
                    size: style.font_size,
                    color: style.font_color.clone(),
                }
        })?;
        let fill = match &style.fill {
            Some(color) => self.fills.iter().position(|existing| existing == color)? + 2,
            None => 0,
        };
        let num_fmt = match style.number_format.as_deref() {
            Some(code) if !code.eq_ignore_ascii_case("general") && !code.trim().is_empty() => {
                self.numbers.iter().find(|(_, existing)| existing == code).map(|(id, _)| *id)?
            }
            _ => 0,
        };
        let xf = Xf { font, fill, num_fmt, align: valid_align(style.align.as_deref()), wrap: style.wrap.unwrap_or(false) };
        self.xfs.iter().position(|existing| existing == &xf)
    }

    fn xml(&self) -> String {
        let mut body = String::new();
        body.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        body.push_str("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        if !self.numbers.is_empty() {
            body.push_str(&format!("<numFmts count=\"{}\">", self.numbers.len()));
            for (id, code) in &self.numbers {
                body.push_str(&format!("<numFmt numFmtId=\"{id}\" formatCode=\"{}\"/>", escape_attribute(code)));
            }
            body.push_str("</numFmts>");
        }
        body.push_str(&format!("<fonts count=\"{}\">", self.fonts.len()));
        for font in &self.fonts {
            body.push_str("<font>");
            if font.bold {
                body.push_str("<b/>");
            }
            if font.italic {
                body.push_str("<i/>");
            }
            if let Some(size) = font.size {
                body.push_str(&format!("<sz val=\"{}\"/>", format_number(f64::from(size))));
            }
            if let Some(color) = &font.color
                && let Some(rgb) = argb(color)
            {
                body.push_str(&format!("<color rgb=\"{rgb}\"/>"));
            }
            body.push_str("</font>");
        }
        body.push_str("</fonts>");
        body.push_str(&format!("<fills count=\"{}\">", self.fills.len() + 2));
        body.push_str("<fill><patternFill patternType=\"none\"/></fill>");
        body.push_str("<fill><patternFill patternType=\"gray125\"/></fill>");
        for fill in &self.fills {
            body.push_str(&format!(
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{}\"/><bgColor indexed=\"64\"/></patternFill></fill>",
                argb(fill).unwrap_or_else(|| "FFEEEEEE".to_owned())
            ));
        }
        body.push_str("</fills>");
        body.push_str("<borders count=\"1\"><border/></borders>");
        body.push_str("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
        body.push_str(&format!("<cellXfs count=\"{}\">", self.xfs.len()));
        for xf in &self.xfs {
            body.push_str(&format!(
                "<xf numFmtId=\"{}\" fontId=\"{}\" fillId=\"{}\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyNumberFormat=\"1\" applyAlignment=\"1\">",
                xf.num_fmt, xf.font, xf.fill
            ));
            if xf.align.is_some() || xf.wrap {
                body.push_str("<alignment");
                if let Some(align) = &xf.align {
                    body.push_str(&format!(" horizontal=\"{}\"", escape_attribute(align)));
                }
                if xf.wrap {
                    body.push_str(" wrapText=\"1\"");
                }
                body.push_str("/>");
            }
            body.push_str("</xf>");
        }
        body.push_str("</cellXfs>");
        body.push_str("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
        body.push_str("</styleSheet>");
        body
    }
}

fn valid_align(align: Option<&str>) -> Option<String> {
    match align {
        Some(value @ ("left" | "center" | "right" | "justify" | "distribute" | "fill")) => Some(value.to_owned()),
        _ => None,
    }
}

/// Converts a stored color into an eight-digit ARGB literal.
fn argb(color: &str) -> Option<String> {
    let hex = color.trim().trim_start_matches('#');
    match hex.len() {
        6 if hex.chars().all(|character| character.is_ascii_hexdigit()) => Some(format!("FF{}", hex.to_ascii_uppercase())),
        8 if hex.chars().all(|character| character.is_ascii_hexdigit()) => Some(hex.to_ascii_uppercase()),
        _ => None,
    }
}

fn format_number(number: f64) -> String {
    crate::model::format_number(number)
}

fn escape_text(text: &str) -> String {
    let mut escaped = String::with_capacity(text.len());
    for character in text.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

fn escape_attribute(text: &str) -> String {
    escape_text(text).replace('"', "&quot;")
}
