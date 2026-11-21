//! Delimited text (CSV/TSV) reading and writing.
//!
//! The quote state machine lives in `data-text-csv` so every tabular product
//! shares one parser. This module keeps only the workbook projection: records
//! become cells, numeric unquoted fields become numbers, and a sheet writes back
//! through its displayed values.

use data_text_csv::table::{escape_field, parse_records, write_records};

use crate::model::{Cell, CellValue, Project, Sheet};

pub use data_text_csv::table::Field;

/// Splits delimited text into records. Blank lines are skipped.
pub fn parse(text: &str, delimiter: char) -> Vec<Vec<Field>> {
    parse_records(text, delimiter)
}

/// Builds a one-sheet project from delimited text.
pub fn project_from_csv(title: impl Into<String>, text: &str, delimiter: char) -> Project {
    let records = parse(text, delimiter);
    let mut project = Project::blank(title);
    let sheet = &mut project.sheets[0];
    for (row_index, record) in records.iter().enumerate() {
        for (column_index, field) in record.iter().enumerate() {
            if field.text.is_empty() && !field.quoted {
                continue;
            }
            let value = if !field.quoted && looks_numeric(&field.text) {
                CellValue::number(field.text.parse::<f64>().unwrap_or(0.0))
            }
            else {
                CellValue::text(&field.text)
            };
            let row = row_index as u32 + 1;
            let column = column_index as u32 + 1;
            sheet.set_cell(Cell {
                reference: crate::reference::format_reference(row, column),
                row,
                column,
                value,
                formula: None,
                cached: None,
                style: None,
                computed: None,
                computed_state: None,
            });
        }
    }
    sheet.dimension = sheet.actual_range();
    project
}

fn looks_numeric(text: &str) -> bool {
    let mut has_digit = false;
    for character in text.chars() {
        if character.is_ascii_digit() {
            has_digit = true;
        }
        else if !matches!(character, '+' | '-' | '.' | 'e' | 'E') {
            return false;
        }
    }
    has_digit
}

/// Serializes one sheet as delimited text with a trailing newline per row.
pub fn write(sheet: &Sheet, delimiter: char) -> String {
    let Some(bounds) = sheet.actual_range()
    else {
        return String::new();
    };
    let mut rows = Vec::new();
    for row in 1..=bounds.end_row {
        let mut cells = Vec::new();
        for column in 1..=bounds.end_column {
            // Delimited text has no formula channel, so a formula projects its
            // freshly evaluated value.
            cells.push(sheet.cell(row, column).map(|cell| crate::eval::value_of(cell).display_text()).unwrap_or_default());
        }
        rows.push(cells);
    }
    write_records(&rows, delimiter)
}

/// Quotes a field when it contains the delimiter, a quote or a line break.
pub fn escape(text: &str, delimiter: char) -> String {
    escape_field(text, delimiter)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reads_quotes_escapes_and_embedded_newlines() {
        let records = parse("name,note\r\n\"Doe, Jane\",\"line 1\nline 2\"\nplain,\"say \"\"hi\"\"\"\n", ',');
        assert_eq!(records.len(), 3);
        assert_eq!(records[0][0].text, "name");
        assert_eq!(records[1][0].text, "Doe, Jane");
        assert_eq!(records[1][1].text, "line 1\nline 2");
        assert_eq!(records[2][1].text, "say \"hi\"");
    }

    #[test]
    fn infers_numbers_only_for_unquoted_fields() {
        let project = project_from_csv("Data", "id,amount\n1,\"2.5\"\n2,3.5\n", ',');
        let sheet = &project.sheets[0];
        assert_eq!(sheet.cell(2, 1).expect("cell").value, CellValue::number(1.0));
        assert_eq!(sheet.cell(2, 2).expect("cell").value, CellValue::text("2.5"));
        assert_eq!(sheet.cell(3, 2).expect("cell").value, CellValue::number(3.5));
    }

    #[test]
    fn writes_escaped_fields() {
        let project = project_from_csv("Data", "a\n\"x,y\"\n", ',');
        let text = write(&project.sheets[0], ',');
        assert_eq!(text, "a\n\"x,y\"\n");
    }
}
