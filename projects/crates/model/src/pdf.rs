//! Self-contained PDF projection of the workbook.
//!
//! Every visible sheet is laid out on A4 pages, with a portrait page when the
//! content fits and landscape otherwise. The writer approximates Excel column
//! widths, uses one fixed row height and keeps the resolved styles of the
//! project: solid fill, bold, font size, font color and horizontal alignment.
//! Values come from the evaluator, so formula results and error literals match
//! the CSV, JSON and XLSX projections. Images are not part of the project
//! model, and font metrics stay approximate, which the capability matrix
//! reports as `style-approximated` and `no-images`.

use crate::{
    eval,
    model::{Cell, CellValue, Project, Sheet},
};

const PAGE_WIDTH: f64 = 595.276;
const PAGE_HEIGHT: f64 = 841.89;
const MARGIN: f64 = 28.0;
const ROW_HEIGHT: f64 = 16.0;
const HEADER_SPACE: f64 = 26.0;
const DEFAULT_FONT_SIZE: f64 = 10.0;

/// One planned PDF page: a sheet, a column band and a row band.
struct PagePlan {
    sheet_index: usize,
    columns: Vec<u32>,
    rows: Vec<u32>,
    landscape: bool,
    index: usize,
    total: usize,
}

/// Renders every visible sheet and returns a complete PDF document.
pub(crate) fn write(project: &Project) -> Vec<u8> {
    let plans = plan_pages(project);
    let pages = plans
        .iter()
        .map(|plan| {
            let (width, height) = page_size(plan.landscape);
            (width, height, page_content(project, plan))
        })
        .collect::<Vec<_>>();
    build_document(pages)
}

fn page_size(landscape: bool) -> (f64, f64) {
    if landscape { (PAGE_HEIGHT, PAGE_WIDTH) } else { (PAGE_WIDTH, PAGE_HEIGHT) }
}

/// Approximated column width in points.
fn column_points(sheet: &Sheet, column: u32) -> f64 {
    (sheet.column_width(column) * 6.0 + 4.0).clamp(24.0, 320.0)
}

fn plan_pages(project: &Project) -> Vec<PagePlan> {
    let mut plans = Vec::new();
    for (sheet_index, sheet) in project.sheets.iter().enumerate() {
        if !sheet.visibility.is_visible() {
            continue;
        }
        let bounds = sheet.actual_range();
        let last_row = bounds.map(|bounds| bounds.end_row).unwrap_or(0);
        let last_column = bounds.map(|bounds| bounds.end_column).unwrap_or(0);
        let widths = (1..=last_column).map(|column| column_points(sheet, column)).collect::<Vec<_>>();
        let portrait_width = PAGE_WIDTH - 2.0 * MARGIN;
        let total_width: f64 = widths.iter().sum();
        let landscape = total_width > portrait_width;
        let (page_width, page_height) = page_size(landscape);
        let content_width = page_width - 2.0 * MARGIN;
        let content_height = page_height - 2.0 * MARGIN;

        let mut column_bands: Vec<Vec<u32>> = Vec::new();
        if last_column == 0 {
            column_bands.push(Vec::new());
        }
        else {
            let mut band = Vec::new();
            let mut used = 0.0;
            for column in 1..=last_column {
                let width = widths[(column - 1) as usize].min(content_width);
                if !band.is_empty() && used + width > content_width + 0.01 {
                    column_bands.push(std::mem::take(&mut band));
                    used = 0.0;
                }
                band.push(column);
                used += width;
            }
            if !band.is_empty() {
                column_bands.push(band);
            }
        }

        let row_capacity = (((content_height - HEADER_SPACE) / ROW_HEIGHT).floor() as u32).max(1);
        let mut row_bands: Vec<Vec<u32>> = Vec::new();
        if last_row == 0 {
            row_bands.push(Vec::new());
        }
        else {
            for start in (1..=last_row).step_by(row_capacity as usize) {
                let end = (start + row_capacity - 1).min(last_row);
                row_bands.push((start..=end).collect());
            }
        }

        let total = column_bands.len() * row_bands.len();
        let mut index = 0usize;
        for rows in &row_bands {
            for columns in &column_bands {
                index += 1;
                plans.push(PagePlan { sheet_index, columns: columns.clone(), rows: rows.clone(), landscape, index, total });
            }
        }
    }
    plans
}

fn page_content(project: &Project, plan: &PagePlan) -> String {
    let sheet = &project.sheets[plan.sheet_index];
    let (page_width, page_height) = page_size(plan.landscape);
    let content_width = page_width - 2.0 * MARGIN;
    let mut ops = String::new();

    let header_y = page_height - MARGIN - 10.0;
    ops.push_str(&format!(
        "BT /F1 9 Tf 0.25 0.29 0.35 rg 1 0 0 1 {MARGIN:.2} {header_y:.2} Tm {} Tj ET\n",
        literal(&format!("{}  Page {} / {}", sheet.name, plan.index, plan.total))
    ));
    ops.push_str(&format!(
        "0.6 w 0.78 0.82 0.88 RG {:.2} {:.2} m {:.2} {:.2} l S\n",
        MARGIN,
        header_y - 6.0,
        page_width - MARGIN,
        header_y - 6.0
    ));

    let widths = plan.columns.iter().map(|column| column_points(sheet, *column).min(content_width)).collect::<Vec<_>>();
    let mut x_offsets = Vec::with_capacity(widths.len() + 1);
    let mut x = MARGIN;
    for width in &widths {
        x_offsets.push(x);
        x += width;
    }
    x_offsets.push(x);

    let top = page_height - MARGIN - HEADER_SPACE;
    for (row_index, row) in plan.rows.iter().enumerate() {
        let bottom = top - (row_index as f64 + 1.0) * ROW_HEIGHT;
        for (column_index, column) in plan.columns.iter().enumerate() {
            let cell = sheet.cell(*row, *column);
            let merge = sheet.merge_covering(*row, *column);
            if let Some(merge) = merge
                && (merge.start_row != *row || merge.start_column != *column)
            {
                continue;
            }
            let row_span =
                merge.map(|merge| (merge.end_row - *row + 1).min((plan.rows.len() - row_index) as u32)).unwrap_or(1).max(1);
            let column_span = merge
                .map(|merge| (merge.end_column - *column + 1).min((plan.columns.len() - column_index) as u32))
                .unwrap_or(1)
                .max(1);
            let width = widths[column_index..column_index + column_span as usize].iter().sum::<f64>();
            let height = ROW_HEIGHT * f64::from(row_span);
            draw_cell(&mut ops, cell, x_offsets[column_index], bottom - ROW_HEIGHT * f64::from(row_span - 1), width, height);
        }
    }
    ops
}

fn draw_cell(ops: &mut String, cell: Option<&Cell>, x: f64, bottom: f64, width: f64, height: f64) {
    let style = cell.and_then(|cell| cell.style.as_ref());
    if let Some((red, green, blue)) = style.and_then(|style| style.fill.as_deref()).and_then(rgb) {
        ops.push_str(&format!("{red:.3} {green:.3} {blue:.3} rg {x:.2} {bottom:.2} {width:.2} {height:.2} re f\n"));
    }
    ops.push_str(&format!("0.5 w 0.86 0.88 0.91 RG {x:.2} {bottom:.2} {width:.2} {height:.2} re S\n"));

    let Some(cell) = cell
    else {
        return;
    };
    let value = eval::value_of(cell);
    let text = value.display_text();
    if text.is_empty() {
        return;
    }
    let font_size = style.and_then(|style| style.font_size).map(f64::from).unwrap_or(DEFAULT_FONT_SIZE).clamp(6.0, 20.0);
    let font = if style.and_then(|style| style.bold).unwrap_or(false) { "F2" } else { "F1" };
    let (red, green, blue) = style.and_then(|style| style.font_color.as_deref()).and_then(rgb).unwrap_or((0.06, 0.09, 0.13));
    let default_align = if matches!(value, CellValue::Number { .. }) { "right" } else { "left" };
    let align = style.and_then(|style| style.align.as_deref()).unwrap_or(default_align);
    let text = fit_text(&text, width, font_size);
    let estimated = text.chars().count() as f64 * font_size * 0.5;
    let text_x = match align {
        "center" => x + (width - estimated) / 2.0,
        "right" => x + width - estimated - 3.0,
        _ => x + 3.0,
    };
    let baseline = bottom + (height - font_size) / 2.0 + font_size * 0.32;
    ops.push_str(&format!(
        "BT /{font} {font_size:.2} Tf {red:.3} {green:.3} {blue:.3} rg 1 0 0 1 {text_x:.2} {baseline:.2} Tm {} Tj ET\n",
        literal(&text)
    ));
}

fn fit_text(text: &str, width: f64, font_size: f64) -> String {
    let capacity = (((width - 6.0) / (font_size * 0.5)).floor() as usize).max(1);
    if text.chars().count() <= capacity {
        return text.to_owned();
    }
    let mut clipped = text.chars().take(capacity.saturating_sub(1)).collect::<String>();
    clipped.push('.');
    clipped
}

/// Parses `#RRGGBB`, `AARRGGBB` or `RGB` into unit components.
fn rgb(value: &str) -> Option<(f64, f64, f64)> {
    let hex = value.trim().trim_start_matches('#');
    let hex = if hex.len() == 8 { &hex[2..] } else { hex };
    if hex.len() != 6 || !hex.chars().all(|character| character.is_ascii_hexdigit()) {
        return None;
    }
    let red = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let green = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let blue = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some((f64::from(red) / 255.0, f64::from(green) / 255.0, f64::from(blue) / 255.0))
}

/// Encodes a PDF literal string, replacing glyphs the base fonts cannot show.
fn literal(text: &str) -> String {
    let mut output = String::with_capacity(text.len() + 2);
    output.push('(');
    for character in text.chars() {
        match character {
            '\\' => output.push_str("\\\\"),
            '(' => output.push_str("\\("),
            ')' => output.push_str("\\)"),
            '\n' | '\r' => output.push(' '),
            character if (' '..='~').contains(&character) => output.push(character),
            _ => output.push('?'),
        }
    }
    output.push(')');
    output
}

fn build_document(pages: Vec<(f64, f64, String)>) -> Vec<u8> {
    let mut objects: Vec<String> = vec![
        String::new(),
        String::new(),
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>".to_owned(),
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>".to_owned(),
    ];
    let mut kids = Vec::new();
    for (width, height, content) in pages {
        objects.push(format!("<< /Length {} >>\nstream\n{}\nendstream", content.len(), content));
        let content_number = objects.len();
        objects.push(format!(
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width:.2} {height:.2}] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents {content_number} 0 R >>"
        ));
        kids.push(objects.len());
    }
    objects[0] = "<< /Type /Catalog /Pages 2 0 R >>".to_owned();
    objects[1] = format!(
        "<< /Type /Pages /Kids [{}] /Count {} >>",
        kids.iter().map(|number| format!("{number} 0 R")).collect::<Vec<_>>().join(" "),
        kids.len()
    );
    serialize(&objects)
}

fn serialize(objects: &[String]) -> Vec<u8> {
    let mut output = b"%PDF-1.4\n".to_vec();
    output.extend_from_slice(b"%\xE2\xE3\xCF\xD3\n");
    let mut offsets = Vec::with_capacity(objects.len());
    for (index, body) in objects.iter().enumerate() {
        offsets.push(output.len());
        output.extend_from_slice(format!("{} 0 obj\n{}\nendobj\n", index + 1, body).as_bytes());
    }
    let start = output.len();
    output.extend_from_slice(format!("xref\n0 {}\n0000000000 65535 f \n", objects.len() + 1).as_bytes());
    for offset in offsets {
        output.extend_from_slice(format!("{offset:010} 00000 n \n").as_bytes());
    }
    output.extend_from_slice(
        format!("trailer\n<< /Size {} /Root 1 0 R >>\nstartxref\n{}\n%%EOF\n", objects.len() + 1, start).as_bytes(),
    );
    output
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{CellRange, CellStyle, CellValue, Project, SheetVisibility};

    #[test]
    fn writes_a_paginated_document_with_an_xref() {
        let mut project = Project::blank("Pdf");
        let sheet = &mut project.sheets[0];
        sheet.write_cell(1, 1, CellValue::text("name"), None, None);
        sheet.write_cell(1, 2, CellValue::number(42.0), None, None);
        sheet.write_cell(2, 1, CellValue::Empty, Some("=B1*2".to_owned()), None);
        sheet.cells[0].style = Some(CellStyle { bold: Some(true), fill: Some("#ffeeaa".to_owned()), ..CellStyle::default() });
        sheet.merged.push(CellRange { start_row: 3, end_row: 4, start_column: 1, end_column: 2 });
        project.sheets.push({
            let mut hidden = Sheet::new("sheet-2", "Hidden", 1);
            hidden.visibility = SheetVisibility::Hidden;
            hidden
        });
        crate::eval::recalculate(&mut project);
        let bytes = write(&project);
        let text = String::from_utf8_lossy(&bytes);
        assert!(text.starts_with("%PDF-1.4"));
        assert!(text.contains("/Type /Catalog"));
        assert!(text.contains("Sheet1  Page 1 / 1"));
        assert!(!text.contains("Hidden  Page"));
        assert!(text.contains("xref"));
        assert!(text.contains("trailer"));
        assert!(text.contains("startxref"));
        assert!(text.trim_end().ends_with("%%EOF"));
    }

    #[test]
    fn splits_rows_across_pages() {
        let mut project = Project::blank("Long");
        for row in 1..=120 {
            project.sheets[0].write_cell(row, 1, CellValue::number(f64::from(row)), None, None);
        }
        crate::eval::recalculate(&mut project);
        let bytes = write(&project);
        let text = String::from_utf8_lossy(&bytes);
        assert!(text.contains("/Count 3"), "pages were `{text}`");
        assert!(text.contains("Sheet1  Page 3 / 3"));
    }
}
