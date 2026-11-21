//! Grid rendering of one sheet as a self-contained SVG document.

use crate::{
    model::{Cell, CellValue, Sheet},
    reference::column_name,
};

const HEADER_WIDTH: f64 = 48.0;
const HEADER_HEIGHT: f64 = 22.0;
const MAX_ROWS: u32 = 500;
const MAX_COLUMNS: u32 = 100;

/// Renders one sheet as an SVG string.
pub fn sheet_svg(sheet: &Sheet) -> String {
    let bounds = sheet.actual_range();
    let rows = bounds.map(|bounds| bounds.end_row.min(MAX_ROWS)).unwrap_or(1).max(1);
    let columns = bounds.map(|bounds| bounds.end_column.min(MAX_COLUMNS)).unwrap_or(1).max(1);

    let mut column_widths = Vec::with_capacity(columns as usize);
    let mut x_offsets = Vec::with_capacity(columns as usize + 1);
    let mut x = HEADER_WIDTH;
    for column in 1..=columns {
        let width = (sheet.column_width(column) * 7.0 + 5.0).clamp(24.0, 400.0);
        x_offsets.push(x);
        column_widths.push(width);
        x += width;
    }
    x_offsets.push(x);

    let mut row_heights = Vec::with_capacity(rows as usize);
    let mut y_offsets = Vec::with_capacity(rows as usize + 1);
    let mut y = HEADER_HEIGHT;
    for row in 1..=rows {
        let height = (sheet.row_height(row) * 96.0 / 72.0).clamp(16.0, 200.0);
        y_offsets.push(y);
        row_heights.push(height);
        y += height;
    }
    y_offsets.push(y);

    let total_width = x;
    let total_height = y;
    let mut svg = String::new();
    svg.push_str(&format!(
        "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{total_width:.0}\" height=\"{total_height:.0}\" viewBox=\"0 0 {total_width:.0} {total_height:.0}\" font-family=\"Segoe UI, Arial, sans-serif\">"
    ));
    svg.push_str("<rect width=\"100%\" height=\"100%\" fill=\"#ffffff\"/>");
    svg.push_str("<rect width=\"100%\" height=\"100%\" fill=\"none\" stroke=\"#d0d5dd\"/>");

    // Column and row headers.
    for column in 1..=columns {
        let index = (column - 1) as usize;
        svg.push_str(&format!(
            "<rect x=\"{:.1}\" y=\"0\" width=\"{:.1}\" height=\"{HEADER_HEIGHT}\" fill=\"#f2f4f7\" stroke=\"#d0d5dd\"/>",
            x_offsets[index], column_widths[index]
        ));
        svg.push_str(&format!(
            "<text x=\"{:.1}\" y=\"{:.1}\" font-size=\"11\" fill=\"#475467\" text-anchor=\"middle\">{}</text>",
            x_offsets[index] + column_widths[index] / 2.0,
            HEADER_HEIGHT / 2.0 + 4.0,
            column_name(column)
        ));
    }
    for row in 1..=rows {
        let index = (row - 1) as usize;
        svg.push_str(&format!(
            "<rect x=\"0\" y=\"{:.1}\" width=\"{HEADER_WIDTH}\" height=\"{:.1}\" fill=\"#f2f4f7\" stroke=\"#d0d5dd\"/>",
            y_offsets[index], row_heights[index]
        ));
        svg.push_str(&format!(
            "<text x=\"{:.1}\" y=\"{:.1}\" font-size=\"11\" fill=\"#475467\" text-anchor=\"middle\">{row}</text>",
            HEADER_WIDTH / 2.0,
            y_offsets[index] + row_heights[index] / 2.0 + 4.0
        ));
    }

    // Grid lines.
    for column in 1..=columns {
        let index = (column - 1) as usize;
        svg.push_str(&format!(
            "<rect x=\"{:.1}\" y=\"{HEADER_HEIGHT}\" width=\"{:.1}\" height=\"{:.1}\" fill=\"none\" stroke=\"#e4e7ec\"/>",
            x_offsets[index],
            column_widths[index],
            total_height - HEADER_HEIGHT
        ));
    }
    for row in 1..=rows {
        let index = (row - 1) as usize;
        svg.push_str(&format!(
            "<rect x=\"{HEADER_WIDTH}\" y=\"{:.1}\" width=\"{:.1}\" height=\"{:.1}\" fill=\"none\" stroke=\"#e4e7ec\"/>",
            y_offsets[index],
            total_width - HEADER_WIDTH,
            row_heights[index]
        ));
    }

    for row in 1..=rows {
        for column in 1..=columns {
            let Some(covering) = sheet.merge_covering(row, column)
            else {
                if let Some(cell) = sheet.cell(row, column) {
                    draw_cell(&mut svg, sheet, cell, row, column, 1, 1);
                }
                continue;
            };
            if covering.start_row != row || covering.start_column != column {
                continue;
            }
            if let Some(cell) = sheet.cell(row, column) {
                let row_span = (covering.end_row.min(rows) - row + 1).max(1);
                let column_span = (covering.end_column.min(columns) - column + 1).max(1);
                draw_cell(&mut svg, sheet, cell, row, column, row_span, column_span);
            }
        }
    }

    svg.push_str("</svg>");
    svg
}

fn draw_cell(svg: &mut String, sheet: &Sheet, cell: &Cell, row: u32, column: u32, row_span: u32, column_span: u32) {
    let bounds = sheet.actual_range();
    let rows = bounds.map(|bounds| bounds.end_row.min(MAX_ROWS)).unwrap_or(1).max(1);
    let columns = bounds.map(|bounds| bounds.end_column.min(MAX_COLUMNS)).unwrap_or(1).max(1);
    if row > rows || column > columns {
        return;
    }
    let x = HEADER_WIDTH + width_before(sheet, column);
    let y = HEADER_HEIGHT + height_before(sheet, row);
    let width = sum_widths(sheet, column, column + column_span - 1, columns);
    let height = sum_heights(sheet, row, row + row_span - 1, rows);

    let style = cell.style.as_ref();
    if let Some(fill) = style.and_then(|style| style.fill.as_deref()).and_then(color) {
        svg.push_str(&format!("<rect x=\"{x:.1}\" y=\"{y:.1}\" width=\"{width:.1}\" height=\"{height:.1}\" fill=\"{fill}\"/>"));
    }

    let text = match &cell.value {
        CellValue::Empty => cell.formula.clone().unwrap_or_default(),
        value => value.display_text(),
    };
    if text.is_empty() {
        return;
    }
    let font_size = style.and_then(|style| style.font_size).map_or(12.0, |size| f64::from(size) * 96.0 / 72.0);
    let weight = if style.and_then(|style| style.bold).unwrap_or(false) { "bold" } else { "normal" };
    let italic = if style.and_then(|style| style.italic).unwrap_or(false) { "italic" } else { "normal" };
    let color = style.and_then(|style| style.font_color.as_deref()).and_then(color).unwrap_or_else(|| "#101828".to_owned());
    let align = style.and_then(|style| style.align.as_deref()).unwrap_or(if matches!(cell.value, CellValue::Number { .. }) {
        "right"
    }
    else {
        "left"
    });
    let (anchor, text_x) = match align {
        "center" => ("middle", x + width / 2.0),
        "right" => ("end", x + width - 4.0),
        _ => ("start", x + 4.0),
    };
    svg.push_str(&format!(
        "<text x=\"{text_x:.1}\" y=\"{:.1}\" font-size=\"{font_size:.1}\" font-weight=\"{weight}\" font-style=\"{italic}\" fill=\"{color}\" text-anchor=\"{anchor}\" dominant-baseline=\"middle\">{}</text>",
        y + height / 2.0,
        escape(&text)
    ));
}

fn width_before(sheet: &Sheet, column: u32) -> f64 {
    (1..column).map(|index| (sheet.column_width(index) * 7.0 + 5.0).clamp(24.0, 400.0)).sum()
}

fn height_before(sheet: &Sheet, row: u32) -> f64 {
    (1..row).map(|index| (sheet.row_height(index) * 96.0 / 72.0).clamp(16.0, 200.0)).sum()
}

fn sum_widths(sheet: &Sheet, start: u32, end: u32, limit: u32) -> f64 {
    (start..=end.min(limit)).map(|index| (sheet.column_width(index) * 7.0 + 5.0).clamp(24.0, 400.0)).sum()
}

fn sum_heights(sheet: &Sheet, start: u32, end: u32, limit: u32) -> f64 {
    (start..=end.min(limit)).map(|index| (sheet.row_height(index) * 96.0 / 72.0).clamp(16.0, 200.0)).sum()
}

fn color(value: &str) -> Option<String> {
    let trimmed = value.trim();
    if let Some(hex) = trimmed.strip_prefix('#') {
        return match hex.len() {
            6 => Some(format!("#{}", hex.to_ascii_lowercase())),
            8 => Some(format!("#{}", hex[2..].to_ascii_lowercase())),
            _ => None,
        };
    }
    if trimmed.len() == 8 && trimmed.chars().all(|character| character.is_ascii_hexdigit()) {
        return Some(format!("#{}", trimmed[2..].to_ascii_lowercase()));
    }
    if trimmed.len() == 6 && trimmed.chars().all(|character| character.is_ascii_hexdigit()) {
        return Some(format!("#{}", trimmed.to_ascii_lowercase()));
    }
    if trimmed.starts_with("rgb(") {
        return Some(trimmed.to_owned());
    }
    None
}

fn escape(text: &str) -> String {
    let mut escaped = String::with_capacity(text.len());
    for character in text.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{CellStyle, CellValue, Project};

    #[test]
    fn renders_values_merges_and_escapes() {
        let mut project = Project::blank("Render");
        let sheet = &mut project.sheets[0];
        sheet.write_cell(1, 1, CellValue::text("<A&B>"), None, None);
        sheet.write_cell(1, 2, CellValue::number(42.0), None, None);
        sheet.merged.push(crate::model::CellRange { start_row: 2, end_row: 3, start_column: 1, end_column: 2 });
        sheet.cells[0].style = Some(CellStyle { bold: Some(true), fill: Some("#ffeeaa".to_owned()), ..CellStyle::default() });
        let svg = sheet_svg(sheet);
        assert!(svg.starts_with("<svg"));
        assert!(svg.contains("&lt;A&amp;B&gt;"));
        assert!(svg.contains(">42<") || svg.contains(">42</text>"));
        assert!(svg.contains("<svg"));
    }
}
