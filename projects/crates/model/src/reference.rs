//! A1 notation as the workbook model needs it.
//!
//! The notation itself lives in `document_formats::a1` so the reader, the writers
//! and every product agree on one parser. What stays here is the workbook's own
//! bound: an XLSX sheet ends at column XFD, so a reference that would name column
//! XFE is rejected here even though the notation alone can spell it.

use document_formats::a1;

use crate::model::CellRange;

/// Last column of an XLSX sheet, `XFD`.
const MAX_COLUMN: u32 = 16_384;

/// Converts an A1 column name such as `AB` into a one-based column index.
pub fn column_number(letters: &str) -> Option<u32> {
    a1::column_number(letters).filter(|column| *column <= MAX_COLUMN)
}

/// Converts a one-based column index into its A1 letters.
pub fn column_name(column: u32) -> String {
    a1::column_name(column)
}

/// Formats one one-based row and column as an A1 reference.
pub fn format_reference(row: u32, column: u32) -> String {
    a1::format_reference(row, column)
}

/// Parses a single A1 cell reference such as `$B$4` into a one-based row and column.
pub fn parse_reference(text: &str) -> Option<(u32, u32)> {
    a1::parse_reference(text).filter(|(_, column)| *column <= MAX_COLUMN)
}

/// Parses `A1` or `A1:B2` text into an inclusive rectangle.
pub fn parse_range(text: &str) -> Option<CellRange> {
    let ((start_row, start_column), (end_row, end_column)) = a1::parse_range(text)?;
    (start_column <= MAX_COLUMN && end_column <= MAX_COLUMN).then(|| CellRange::between(start_row, start_column, end_row, end_column))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn columns_round_trip() {
        for (number, name) in [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (702, "ZZ"), (703, "AAA"), (16384, "XFD")] {
            assert_eq!(column_name(number), name);
            assert_eq!(column_number(name), Some(number));
        }
    }

    #[test]
    fn parses_absolute_and_relative_references() {
        assert_eq!(parse_reference("A1"), Some((1, 1)));
        assert_eq!(parse_reference("$B$4"), Some((4, 2)));
        assert_eq!(parse_reference("XFD1048576"), Some((1048576, 16384)));
        assert_eq!(parse_reference("A"), None);
        assert_eq!(parse_reference("1"), None);
        assert_eq!(parse_reference("A0"), None);
    }

    #[test]
    fn rejects_columns_past_the_sheet_edge() {
        assert_eq!(column_number("XFE"), None);
        assert_eq!(parse_reference("XFE1"), None);
        assert_eq!(parse_range("A1:XFE9"), None);
    }

    #[test]
    fn parses_ranges() {
        let range = parse_range("B2:D5").expect("range");
        assert_eq!((range.start_row, range.end_row, range.start_column, range.end_column), (2, 5, 2, 4));
        let single = parse_range("C3").expect("single");
        assert_eq!((single.start_row, single.end_row, single.start_column, single.end_column), (3, 3, 3, 3));
        assert_eq!(parse_range("A1:B2:C3"), None);
    }
}
