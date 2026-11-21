//! A1 reference parsing and formatting shared by the model, the inspector and the writers.

use crate::model::CellRange;

/// Converts an A1 column name such as `AB` into a one-based column index.
pub fn column_number(letters: &str) -> Option<u32> {
    if letters.is_empty() || letters.len() > 3 {
        return None;
    }
    let mut value = 0u32;
    for character in letters.chars() {
        let upper = character.to_ascii_uppercase();
        if !upper.is_ascii_uppercase() {
            return None;
        }
        value = value.checked_mul(26)?.checked_add(u32::from(upper as u8 - b'A' + 1))?;
    }
    Some(value)
}

/// Converts a one-based column index into its A1 letters.
pub fn column_name(mut column: u32) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let remainder = ((column - 1) % 26) as u8;
        letters.push((b'A' + remainder) as char);
        column = (column - 1) / 26;
    }
    letters.reverse();
    letters.into_iter().collect()
}

/// Formats one one-based row and column as an A1 reference.
pub fn format_reference(row: u32, column: u32) -> String {
    format!("{}{}", column_name(column), row)
}

/// Parses a single A1 cell reference such as `$B$4` into a one-based row and column.
pub fn parse_reference(text: &str) -> Option<(u32, u32)> {
    let mut letters = String::new();
    let mut digits = String::new();
    for character in text.trim().chars() {
        if character == '$' {
            continue;
        }
        if character.is_ascii_alphabetic() && digits.is_empty() {
            letters.push(character);
        }
        else if character.is_ascii_digit() {
            digits.push(character);
        }
        else {
            return None;
        }
    }
    let column = column_number(&letters)?;
    let row: u32 = digits.parse().ok()?;
    (row > 0 && column > 0).then_some((row, column))
}

/// Parses `A1`, `A1:B2` or `A1:B2:C3` text into an inclusive rectangle.
pub fn parse_range(text: &str) -> Option<CellRange> {
    let mut parts = text.split(':');
    let first = parse_reference(parts.next()?)?;
    let second = match parts.next() {
        Some(part) => parse_reference(part)?,
        None => first,
    };
    if parts.next().is_some() {
        return None;
    }
    Some(CellRange::between(first.0, first.1, second.0, second.1))
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
    fn parses_ranges() {
        let range = parse_range("B2:D5").expect("range");
        assert_eq!((range.start_row, range.end_row, range.start_column, range.end_column), (2, 5, 2, 4));
        let single = parse_range("C3").expect("single");
        assert_eq!((single.start_row, single.end_row, single.start_column, single.end_column), (3, 3, 3, 3));
    }
}
