//! Moves A1 references when whole rows or columns are inserted, deleted or
//! swapped, or when a range edit shifts a band of cells.
//!
//! A formula is scanned as text. String literals, error literals, function names,
//! defined names and worksheet qualifiers stay verbatim. Only references that
//! point at the edited sheet move; references to other sheets, to another
//! workbook, or to a 3-D sheet span are left alone. A reference removed by a
//! deletion becomes `#REF!` instead of being dropped.

use crate::{model::Project, reference};

/// Last row of an XLSX sheet.
const MAX_ROW: u32 = 1_048_576;
/// Last column of an XLSX sheet, `XFD`.
const MAX_COLUMN: u32 = 16_384;

/// Axis one structural edit moves.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Axis {
    /// Rows move; columns stay.
    Row,
    /// Columns move; rows stay.
    Column,
}

/// How one structural edit moves an axis.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Shift {
    /// Insert `count` lines before line `at`.
    Insert {
        /// First inserted line.
        at: u32,
        /// Number of inserted lines.
        count: u32,
    },
    /// Delete lines `start..=end`.
    Delete {
        /// First deleted line.
        start: u32,
        /// Last deleted line.
        end: u32,
    },
    /// Exchange lines `first` and `second`.
    Swap {
        /// One of the exchanged lines.
        first: u32,
        /// The other exchanged line.
        second: u32,
    },
    /// Insert `count` lines before `at`, moving only lines whose across-axis
    /// coordinate lies in `across_start..=across_end`.
    BandInsert {
        /// First inserted line.
        at: u32,
        /// Number of inserted lines.
        count: u32,
        /// First across-axis line of the edited band.
        across_start: u32,
        /// Last across-axis line of the edited band.
        across_end: u32,
    },
    /// Delete lines `start..=end`, removing and shifting only lines whose
    /// across-axis coordinate lies in `across_start..=across_end`.
    BandDelete {
        /// First deleted line.
        start: u32,
        /// Last deleted line.
        end: u32,
        /// First across-axis line of the edited band.
        across_start: u32,
        /// Last across-axis line of the edited band.
        across_end: u32,
    },
}

/// Rewrites the formulas of every sheet after a structural edit.
///
/// `target_sheet` names the edited sheet. An unqualified reference is local to
/// its host sheet, so it moves only when the host is the edited sheet; a
/// qualified reference moves only when its qualifier names the edited sheet.
pub(crate) fn shift_project(project: &mut Project, target_sheet: &str, axis: Axis, shift: Shift) {
    for sheet in &mut project.sheets {
        let host = sheet.name.clone();
        for cell in &mut sheet.cells {
            let Some(source) = cell.formula.as_ref()
            else {
                continue;
            };
            let next = shift_formula(source, target_sheet, &host, axis, shift);
            if next != *source {
                cell.formula = Some(next);
                // The cached result belonged to the old references.
                cell.cached = None;
            }
        }
    }
}

/// Shifts the references of one formula that belong to `target_sheet`.
///
/// `host_sheet` names the sheet the formula lives on, which decides whether an
/// unqualified reference is local. The returned text is the formula source with
/// the same leading `=`, if any.
pub(crate) fn shift_formula(source: &str, target_sheet: &str, host_sheet: &str, axis: Axis, shift: Shift) -> String {
    let chars: Vec<char> = source.chars().collect();
    let mut out = String::with_capacity(source.len());
    let mut index = 0;
    while index < chars.len() {
        let character = chars[index];
        if character == '"' {
            index = copy_quoted(&chars, index, '"', &mut out);
            continue;
        }
        if character == '[' {
            index = copy_external(&chars, index, &mut out);
            continue;
        }
        if let Some((qualifier, after)) = parse_qualifier(&chars, index) {
            let left_local = qualifier.name.eq_ignore_ascii_case(target_sheet);
            out.push_str(&qualifier.raw);
            if let Some(group) = parse_group(&chars, after) {
                let local = left_local && right_is_local(&group, target_sheet);
                emit_group(&chars, &group, local, axis, shift, &mut out);
                index = group.end;
            }
            else {
                index = after;
            }
            continue;
        }
        if character == '\'' {
            index = copy_quoted(&chars, index, '\'', &mut out);
            continue;
        }
        let after_name = index > 0 && is_name_char(chars[index - 1]);
        if !after_name && let Some(group) = parse_group(&chars, index) {
            let local = host_sheet.eq_ignore_ascii_case(target_sheet) && right_is_local(&group, target_sheet);
            emit_group(&chars, &group, local, axis, shift, &mut out);
            index = group.end;
            continue;
        }
        out.push(character);
        index += 1;
    }
    out
}

/// Writes one reference group, or its verbatim source text when it stays put.
fn emit_group(chars: &[char], group: &Group, local: bool, axis: Axis, shift: Shift, out: &mut String) {
    if local {
        match rewrite_group(group, axis, shift) {
            Transformed::Unchanged => out.extend(chars[group.start..group.end].iter()),
            Transformed::Text(text) => out.push_str(&text),
            Transformed::Deleted => out.push_str("#REF!"),
        }
    }
    else {
        out.extend(chars[group.start..group.end].iter());
    }
}

/// Copies a double-quoted string literal, including doubled quote escapes.
fn copy_quoted(chars: &[char], start: usize, delimiter: char, out: &mut String) -> usize {
    let mut index = start + 1;
    while index < chars.len() {
        if chars[index] == delimiter {
            if chars.get(index + 1) == Some(&delimiter) {
                index += 2;
                continue;
            }
            index += 1;
            break;
        }
        index += 1;
    }
    let end = index.min(chars.len());
    out.extend(chars[start..end].iter());
    end
}

/// Copies an external-workbook or structured-reference bracket verbatim.
fn copy_external(chars: &[char], start: usize, out: &mut String) -> usize {
    let mut index = start;
    while index < chars.len() && chars[index] != ']' {
        index += 1;
    }
    if index < chars.len() {
        index += 1;
    }
    out.extend(chars[start..index].iter());
    if let Some((qualifier, after)) = parse_qualifier(chars, index) {
        out.push_str(&qualifier.raw);
        if let Some(group) = parse_group(chars, after) {
            out.extend(chars[group.start..group.end].iter());
            return group.end;
        }
        return after;
    }
    index
}

/// Whether a character continues a defined name, sheet name or number.
fn is_name_char(character: char) -> bool {
    character.is_ascii_alphanumeric() || character == '_' || character == '.' || character == '\\'
}

/// Skips whitespace between a name and a following `(`.
fn skip_spaces(chars: &[char], mut index: usize) -> usize {
    while index < chars.len() && chars[index].is_whitespace() {
        index += 1;
    }
    index
}

/// A worksheet qualifier such as `Sheet2!` or `'My Sheet'!`.
#[derive(Clone)]
struct Qualifier {
    /// Unquoted sheet name the qualifier spells.
    name: String,
    /// Source text of the qualifier, including the trailing `!`.
    raw: String,
}

/// Which kind of reference a token is.
#[derive(Clone, Copy, PartialEq, Eq)]
enum RefKind {
    /// A cell such as `B4`.
    Cell,
    /// A whole column such as `B`, valid only as a range end.
    Column,
    /// A whole row such as `4`, valid only as a range end.
    Row,
}

/// One lexed A1 reference.
#[derive(Clone, Copy)]
struct RefToken {
    kind: RefKind,
    column: u32,
    row: u32,
    absolute_column: bool,
    absolute_row: bool,
    end: usize,
}

/// One single reference or one `start:end` range.
struct Group {
    start: usize,
    end: usize,
    left: RefToken,
    right: Option<(Option<Qualifier>, RefToken)>,
}

/// Outcome of moving one reference or range.
enum Transformed {
    /// The reference does not move.
    Unchanged,
    /// The reference moves to the given text.
    Text(String),
    /// The reference was deleted and becomes `#REF!`.
    Deleted,
}

/// Outcome of moving one line on the edited axis.
enum PointShift {
    Kept,
    Moved(u32),
    Removed,
}

/// Parses a worksheet qualifier ending in `!`.
fn parse_qualifier(chars: &[char], start: usize) -> Option<(Qualifier, usize)> {
    match *chars.get(start)? {
        '\'' => {
            let mut index = start + 1;
            let mut name = String::new();
            loop {
                let character = *chars.get(index)?;
                if character == '\'' {
                    if chars.get(index + 1) == Some(&'\'') {
                        name.push('\'');
                        index += 2;
                    }
                    else {
                        index += 1;
                        break;
                    }
                }
                else {
                    name.push(character);
                    index += 1;
                }
            }
            if chars.get(index) != Some(&'!') {
                return None;
            }
            let raw: String = chars[start..=index].iter().collect();
            Some((Qualifier { name, raw }, index + 1))
        }
        character if character.is_ascii_alphabetic() || character == '_' || character == '\\' => {
            let mut index = start;
            while index < chars.len() && is_name_char(chars[index]) {
                index += 1;
            }
            if chars.get(index) != Some(&'!') {
                return None;
            }
            let name: String = chars[start..index].iter().collect();
            let raw: String = chars[start..=index].iter().collect();
            Some((Qualifier { name, raw }, index + 1))
        }
        _ => None,
    }
}

/// Parses one cell, column or row reference token.
fn parse_ref(chars: &[char], start: usize) -> Option<RefToken> {
    let mut index = start;
    let absolute_column = chars.get(index) == Some(&'$');
    if absolute_column {
        index += 1;
    }
    match *chars.get(index)? {
        first if first.is_ascii_alphabetic() => {
            let letters_start = index;
            while index < chars.len() && chars[index].is_ascii_alphabetic() {
                index += 1;
            }
            let letters: String = chars[letters_start..index].iter().collect();
            if letters.len() > 3 {
                return None;
            }
            let column = reference::column_number(&letters)?;
            let absolute_row = chars.get(index) == Some(&'$');
            if absolute_row {
                index += 1;
            }
            if chars.get(index).is_some_and(|character| character.is_ascii_digit()) {
                let row_start = index;
                while index < chars.len() && chars[index].is_ascii_digit() {
                    index += 1;
                }
                if index - row_start > 7 {
                    return None;
                }
                let row: u32 = chars[row_start..index].iter().collect::<String>().parse().ok()?;
                if !(1..=MAX_ROW).contains(&row) {
                    return None;
                }
                if chars.get(index).is_some_and(|character| is_name_char(*character)) {
                    return None;
                }
                // `LOG10(...)` is a function name, not the cell `LOG10`.
                if chars.get(skip_spaces(chars, index)) == Some(&'(') {
                    return None;
                }
                Some(RefToken { kind: RefKind::Cell, column, row, absolute_column, absolute_row, end: index })
            }
            else {
                if chars.get(index).is_some_and(|character| is_name_char(*character)) {
                    return None;
                }
                Some(RefToken { kind: RefKind::Column, column, row: 0, absolute_column, absolute_row: false, end: index })
            }
        }
        first if first.is_ascii_digit() => {
            let row_start = index;
            while index < chars.len() && chars[index].is_ascii_digit() {
                index += 1;
            }
            if index - row_start > 7 {
                return None;
            }
            let row: u32 = chars[row_start..index].iter().collect::<String>().parse().ok()?;
            if !(1..=MAX_ROW).contains(&row) {
                return None;
            }
            if chars.get(index).is_some_and(|character| is_name_char(*character)) {
                return None;
            }
            Some(RefToken {
                kind: RefKind::Row,
                column: 0,
                row,
                absolute_column: false,
                absolute_row: absolute_column,
                end: index,
            })
        }
        _ => None,
    }
}

/// Parses a single reference or a `reference:reference` range.
///
/// A bare column or row token is only a reference when it is the end of a range,
/// so function names and plain numbers never become references.
fn parse_group(chars: &[char], start: usize) -> Option<Group> {
    let left = parse_ref(chars, start)?;
    let mut end = left.end;
    let mut right = None;
    if chars.get(end) == Some(&':') {
        let after_colon = end + 1;
        let (qualifier, next) = match parse_qualifier(chars, after_colon) {
            Some((qualifier, next)) => (Some(qualifier), next),
            None => (None, after_colon),
        };
        if let Some(right_ref) = parse_ref(chars, next)
            && right_ref.kind == left.kind
        {
            end = right_ref.end;
            right = Some((qualifier, right_ref));
        }
    }
    if right.is_none() && left.kind != RefKind::Cell {
        return None;
    }
    Some(Group { start, end, left, right })
}

/// Whether the right end of a group stays on the edited sheet.
///
/// A missing right qualifier inherits the left one, so an unqualified right end
/// is local whenever the left end is. This helper is only consulted after the
/// left end was found local.
fn right_is_local(group: &Group, target_sheet: &str) -> bool {
    match &group.right {
        Some((Some(qualifier), _)) => qualifier.name.eq_ignore_ascii_case(target_sheet),
        _ => true,
    }
}

/// Whether a reference kind moves along the edited axis.
fn applies(kind: RefKind, axis: Axis) -> bool {
    matches!((kind, axis), (RefKind::Cell, _) | (RefKind::Column, Axis::Column) | (RefKind::Row, Axis::Row))
}

/// Coordinate of a reference on the edited axis.
fn coordinate(reference: &RefToken, axis: Axis) -> u32 {
    match axis {
        Axis::Row => reference.row,
        Axis::Column => reference.column,
    }
}

/// Last valid coordinate on the edited axis.
fn axis_max(axis: Axis) -> u32 {
    match axis {
        Axis::Row => MAX_ROW,
        Axis::Column => MAX_COLUMN,
    }
}

/// Maps one line through an insert or delete.
fn shift_point(value: u32, axis: Axis, shift: Shift) -> PointShift {
    match shift {
        Shift::Insert { at, count } => {
            if value >= at {
                let moved = value + count;
                if moved > axis_max(axis) { PointShift::Removed } else { PointShift::Moved(moved) }
            }
            else {
                PointShift::Kept
            }
        }
        Shift::Delete { start, end } => {
            if value < start {
                PointShift::Kept
            }
            else if value > end {
                PointShift::Moved(value - (end - start + 1))
            }
            else {
                PointShift::Removed
            }
        }
        Shift::Swap { first, second } => {
            if value == first {
                PointShift::Moved(second)
            }
            else if value == second {
                PointShift::Moved(first)
            }
            else {
                PointShift::Kept
            }
        }
        // Band shifts need the across-axis coordinate and are resolved by `shift_line`.
        Shift::BandInsert { .. } | Shift::BandDelete { .. } => PointShift::Kept,
    }
}

/// Whether a shift only moves a band of across-axis lines.
fn is_band(shift: Shift) -> bool {
    matches!(shift, Shift::BandInsert { .. } | Shift::BandDelete { .. })
}

/// Coordinate of a reference on the axis perpendicular to the edit.
fn across_coordinate(reference: &RefToken, axis: Axis) -> u32 {
    match axis {
        Axis::Row => reference.column,
        Axis::Column => reference.row,
    }
}

/// Maps one line through a shift, given its across-axis coordinate.
fn shift_line(line: u32, across: u32, axis: Axis, shift: Shift) -> PointShift {
    match shift {
        Shift::BandInsert { at, count, across_start, across_end } => {
            if (across_start..=across_end).contains(&across) {
                shift_point(line, axis, Shift::Insert { at, count })
            }
            else {
                PointShift::Kept
            }
        }
        Shift::BandDelete { start, end, across_start, across_end } => {
            if (across_start..=across_end).contains(&across) {
                shift_point(line, axis, Shift::Delete { start, end })
            }
            else {
                PointShift::Kept
            }
        }
        other => shift_point(line, axis, other),
    }
}

/// New line of one range endpoint under a band shift; `None` means it vanished.
fn shift_endpoint(reference: &RefToken, axis: Axis, shift: Shift) -> Option<u32> {
    match shift_line(coordinate(reference, axis), across_coordinate(reference, axis), axis, shift) {
        PointShift::Kept => Some(coordinate(reference, axis)),
        PointShift::Moved(value) => Some(value),
        PointShift::Removed => None,
    }
}

/// Line an endpoint inside a deleted band clamps to, given the surviving side.
fn clamp_to_band(survivor: u32, shift: Shift) -> u32 {
    let Shift::BandDelete { start, .. } = shift
    else {
        return survivor;
    };
    if survivor >= start { start } else { start - 1 }
}

/// Orders two moved endpoints the way the original endpoints were ordered.
fn orient(first: u32, second: u32, moved_first: u32, moved_second: u32) -> (u32, u32) {
    if first <= second {
        (moved_first.min(moved_second), moved_first.max(moved_second))
    }
    else {
        (moved_first.max(moved_second), moved_first.min(moved_second))
    }
}

/// Rewrites a single reference or a range.
fn rewrite_group(group: &Group, axis: Axis, shift: Shift) -> Transformed {
    match &group.right {
        None => rewrite_single(&group.left, axis, shift),
        Some(_) => rewrite_range(group, axis, shift),
    }
}

/// Rewrites one single reference.
fn rewrite_single(reference: &RefToken, axis: Axis, shift: Shift) -> Transformed {
    if !applies(reference.kind, axis) || (is_band(shift) && reference.kind != RefKind::Cell) {
        return Transformed::Unchanged;
    }
    match shift_line(coordinate(reference, axis), across_coordinate(reference, axis), axis, shift) {
        PointShift::Kept => Transformed::Unchanged,
        PointShift::Moved(value) => Transformed::Text(render_ref(reference, axis, value)),
        PointShift::Removed => Transformed::Deleted,
    }
}

/// Rewrites a two-ended range, clamping its ends when a deletion cuts it.
fn rewrite_range(group: &Group, axis: Axis, shift: Shift) -> Transformed {
    let left = group.left;
    let Some((qualifier, right)) = group.right.as_ref()
    else {
        return Transformed::Unchanged;
    };
    if !applies(left.kind, axis) {
        return Transformed::Unchanged;
    }
    let first = coordinate(&left, axis);
    let second = coordinate(right, axis);
    if is_band(shift) {
        if left.kind != RefKind::Cell {
            return Transformed::Unchanged;
        }
        // Each endpoint moves on its own across-axis coordinate. A deleted
        // endpoint clamps toward the surviving one, like a whole-line delete.
        let (new_first, new_second) = match (shift_endpoint(&left, axis, shift), shift_endpoint(right, axis, shift)) {
            (None, None) => return Transformed::Deleted,
            (None, Some(survivor)) => orient(first, second, clamp_to_band(survivor, shift), survivor),
            (Some(survivor), None) => orient(first, second, survivor, clamp_to_band(survivor, shift)),
            (Some(moved_first), Some(moved_second)) => orient(first, second, moved_first, moved_second),
        };
        if new_first == first && new_second == second {
            return Transformed::Unchanged;
        }
        let mut text = render_ref(&left, axis, new_first);
        text.push(':');
        if let Some(qualifier) = qualifier {
            text.push_str(&qualifier.raw);
        }
        text.push_str(&render_ref(right, axis, new_second));
        return Transformed::Text(text);
    }
    let (low, high) = (first.min(second), first.max(second));
    let (new_low, new_high, changed) = match shift {
        Shift::Insert { at, count } => {
            let move_line = |value: u32| if value >= at { value + count } else { value };
            let new_low = move_line(low);
            let new_high = move_line(high);
            (new_low, new_high, new_low != low || new_high != high)
        }
        Shift::Swap { first, second } => {
            let swap_line = |value: u32| {
                if value == first {
                    second
                }
                else if value == second {
                    first
                }
                else {
                    value
                }
            };
            let (new_low, new_high) = (swap_line(low).min(swap_line(high)), swap_line(low).max(swap_line(high)));
            (new_low, new_high, new_low != low || new_high != high)
        }
        Shift::Delete { start, end } => {
            let count = end - start + 1;
            if high < start {
                (low, high, false)
            }
            else if low > end {
                (low - count, high - count, true)
            }
            else {
                let new_low = if low < start { low } else { start };
                let new_high = if high > end { high - count } else { start - 1 };
                if new_high < new_low {
                    return Transformed::Deleted;
                }
                (new_low, new_high, new_low != low || new_high != high)
            }
        }
        // Band shifts return above; whole-line rules never see them here.
        Shift::BandInsert { .. } | Shift::BandDelete { .. } => unreachable!("band ranges are rewritten per endpoint"),
    };
    if new_high > axis_max(axis) {
        return Transformed::Deleted;
    }
    if !changed {
        return Transformed::Unchanged;
    }
    let (first_value, second_value) = if first <= second { (new_low, new_high) } else { (new_high, new_low) };
    let mut text = render_ref(&left, axis, first_value);
    text.push(':');
    if let Some(qualifier) = qualifier {
        text.push_str(&qualifier.raw);
    }
    text.push_str(&render_ref(right, axis, second_value));
    Transformed::Text(text)
}

/// Formats one reference with a replacement coordinate on the edited axis.
fn render_ref(reference: &RefToken, axis: Axis, value: u32) -> String {
    let column = if axis == Axis::Column { value } else { reference.column };
    let row = if axis == Axis::Row { value } else { reference.row };
    match reference.kind {
        RefKind::Cell => format!(
            "{}{}{}{}",
            if reference.absolute_column { "$" } else { "" },
            reference::column_name(column),
            if reference.absolute_row { "$" } else { "" },
            row
        ),
        RefKind::Column => format!("{}{}", if reference.absolute_column { "$" } else { "" }, reference::column_name(column)),
        RefKind::Row => format!("{}{}", if reference.absolute_row { "$" } else { "" }, row),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn insert_row(source: &str, at: u32, count: u32) -> String {
        shift_formula(source, "Sheet1", "Sheet1", Axis::Row, Shift::Insert { at, count })
    }

    fn delete_row(source: &str, start: u32, end: u32) -> String {
        shift_formula(source, "Sheet1", "Sheet1", Axis::Row, Shift::Delete { start, end })
    }

    fn insert_column(source: &str, at: u32, count: u32) -> String {
        shift_formula(source, "Sheet1", "Sheet1", Axis::Column, Shift::Insert { at, count })
    }

    fn delete_column(source: &str, start: u32, end: u32) -> String {
        shift_formula(source, "Sheet1", "Sheet1", Axis::Column, Shift::Delete { start, end })
    }

    #[test]
    fn inserting_rows_moves_cells_at_and_after_the_point() {
        assert_eq!(insert_row("=A1+A2", 2, 1), "=A1+A3");
        assert_eq!(insert_row("=B4", 1, 2), "=B6");
        assert_eq!(insert_row("=$B$4", 1, 1), "=$B$5");
        assert_eq!(insert_row("=B$4+$B4", 2, 1), "=B$5+$B5");
    }

    #[test]
    fn inserting_rows_expands_ranges_that_cross_the_point() {
        assert_eq!(insert_row("=SUM(A1:A3)", 2, 2), "=SUM(A1:A5)");
        assert_eq!(insert_row("=SUM(A2:A3)", 2, 1), "=SUM(A3:A4)");
        assert_eq!(insert_row("=SUM(A1:A2)", 5, 1), "=SUM(A1:A2)");
    }

    #[test]
    fn deleting_rows_marks_references_inside_the_band() {
        assert_eq!(delete_row("=A5", 2, 3), "=A3");
        assert_eq!(delete_row("=A2", 2, 3), "=#REF!");
        assert_eq!(delete_row("=$A$2", 2, 3), "=#REF!");
    }

    #[test]
    fn deleting_rows_clamps_ranges() {
        assert_eq!(delete_row("=SUM(A1:A5)", 3, 4), "=SUM(A1:A3)");
        assert_eq!(delete_row("=SUM(A5:A9)", 3, 6), "=SUM(A3:A5)");
        assert_eq!(delete_row("=SUM(A1:A9)", 3, 6), "=SUM(A1:A5)");
        assert_eq!(delete_row("=SUM(A3:A4)", 3, 4), "=SUM(#REF!)");
        assert_eq!(delete_row("=SUM(A2:A3)", 5, 6), "=SUM(A2:A3)");
    }

    #[test]
    fn inserting_and_deleting_columns_use_the_same_rules() {
        assert_eq!(insert_column("=B1+C1", 2, 1), "=C1+D1");
        assert_eq!(insert_column("=SUM(B2:C3)", 3, 2), "=SUM(B2:E3)");
        assert_eq!(delete_column("=C1", 1, 2), "=A1");
        assert_eq!(delete_column("=SUM(A1:C1)", 2, 2), "=SUM(A1:B1)");
        assert_eq!(delete_column("=SUM(B1:C1)", 1, 3), "=SUM(#REF!)");
    }

    #[test]
    fn strings_error_literals_and_function_names_stay_verbatim() {
        assert_eq!(insert_row(r#"="A1"&A2"#, 1, 1), r#"="A1"&A3"#);
        assert_eq!(insert_row("=LOG10(A1)", 1, 1), "=LOG10(A2)");
        assert_eq!(insert_row("=#REF!+A1", 1, 1), "=#REF!+A2");
        assert_eq!(insert_row("=MyName+A1", 1, 1), "=MyName+A2");
    }

    #[test]
    fn other_sheets_and_external_workbooks_do_not_move() {
        assert_eq!(insert_row("=Sheet2!A1", 1, 1), "=Sheet2!A1");
        assert_eq!(insert_row("=Sheet1!A1", 1, 1), "=Sheet1!A2");
        assert_eq!(insert_row("=[Book.xlsx]Sheet1!A1", 1, 1), "=[Book.xlsx]Sheet1!A1");
        assert_eq!(insert_row("=Sheet1:Sheet3!A1", 1, 1), "=Sheet1:Sheet3!A1");
        assert_eq!(insert_row("=A1:Sheet2!B2", 1, 1), "=A1:Sheet2!B2");
    }

    #[test]
    fn quoted_sheet_names_move_only_for_the_edited_sheet() {
        assert_eq!(
            shift_formula("='My Sheet'!A1+A1", "My Sheet", "My Sheet", Axis::Row, Shift::Insert { at: 1, count: 1 }),
            "='My Sheet'!A2+A2"
        );
        assert_eq!(insert_row("='My Sheet'!A1", 1, 1), "='My Sheet'!A1");
    }

    #[test]
    fn unqualified_references_on_other_sheets_do_not_move() {
        assert_eq!(shift_formula("=A1", "Sheet1", "Sheet2", Axis::Row, Shift::Insert { at: 1, count: 1 }), "=A1");
        assert_eq!(shift_formula("=Sheet1!A1", "Sheet1", "Sheet2", Axis::Row, Shift::Insert { at: 1, count: 1 }), "=Sheet1!A2");
    }

    #[test]
    fn whole_column_and_row_ranges_move_on_their_own_axis() {
        assert_eq!(insert_column("=SUM(A:A)", 1, 1), "=SUM(B:B)");
        assert_eq!(insert_column("=SUM($A:$A)", 2, 1), "=SUM($A:$A)");
        assert_eq!(insert_row("=SUM(1:1)", 1, 2), "=SUM(3:3)");
        assert_eq!(insert_row("=SUM(A:A)", 1, 1), "=SUM(A:A)");
    }

    #[test]
    fn formulas_without_references_are_untouched() {
        assert_eq!(insert_row("=1+2", 1, 1), "=1+2");
        assert_eq!(insert_row("=SUM(1,2,3)", 1, 1), "=SUM(1,2,3)");
    }

    #[test]
    fn band_inserts_only_move_cells_inside_the_band() {
        let shift = Shift::BandInsert { at: 2, count: 1, across_start: 1, across_end: 2 };
        assert_eq!(shift_formula("=A5+B5+C5", "Sheet1", "Sheet1", Axis::Row, shift), "=A6+B6+C5");
        assert_eq!(shift_formula("=A1", "Sheet1", "Sheet1", Axis::Row, shift), "=A1");
        assert_eq!(shift_formula("=SUM(A2:A5)", "Sheet1", "Sheet1", Axis::Row, shift), "=SUM(A3:A6)");
        assert_eq!(shift_formula("=SUM(A2:C2)", "Sheet1", "Sheet1", Axis::Row, shift), "=SUM(A2:C3)");
        assert_eq!(shift_formula("=SUM(1:1)", "Sheet1", "Sheet1", Axis::Row, shift), "=SUM(1:1)");
    }

    #[test]
    fn band_deletes_only_remove_cells_inside_the_band() {
        let shift = Shift::BandDelete { start: 2, end: 3, across_start: 1, across_end: 2 };
        assert_eq!(shift_formula("=A5+C5", "Sheet1", "Sheet1", Axis::Row, shift), "=A3+C5");
        assert_eq!(shift_formula("=A2", "Sheet1", "Sheet1", Axis::Row, shift), "=#REF!");
        assert_eq!(shift_formula("=SUM(A2:A5)", "Sheet1", "Sheet1", Axis::Row, shift), "=SUM(A2:A3)");
        assert_eq!(shift_formula("=SUM(A1:A2)", "Sheet1", "Sheet1", Axis::Row, shift), "=SUM(A1:A1)");
        assert_eq!(shift_formula("=SUM(B2:B3)", "Sheet1", "Sheet1", Axis::Row, shift), "=SUM(#REF!)");
        assert_eq!(shift_formula("=SUM(A2:C2)", "Sheet1", "Sheet1", Axis::Row, shift), "=SUM(A2:C2)");
    }
}
