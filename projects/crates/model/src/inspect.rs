//! Rule checks over the editable project model.

use serde_json::{Value, json};

use crate::{
    error::ServiceError,
    model::{CellValue, ComputedState, Project, Sheet},
    reference::parse_reference,
};

/// Stable version of the rule set behind every inspection report.
pub const CHECKER_VERSION: &str = "open-excel/checker/1";

/// Rule names accepted by the inspector.
pub const CHECKS: &[&str] = &["formulas", "references", "cycles", "types", "ranges", "styles"];

const MAX_ROW: u32 = 1_048_576;
const MAX_COLUMN: u32 = 16_384;

/// One inspection finding.
#[derive(Debug, Clone)]
struct Issue {
    code: &'static str,
    severity: &'static str,
    sheet_id: String,
    sheet_name: String,
    cell: String,
    message: String,
    recoverable: bool,
}

impl Issue {
    fn new(
        code: &'static str,
        severity: &'static str,
        sheet: &Sheet,
        cell: impl Into<String>,
        message: impl Into<String>,
        recoverable: bool,
    ) -> Self {
        Self {
            code,
            severity,
            sheet_id: sheet.id.clone(),
            sheet_name: sheet.name.clone(),
            cell: cell.into(),
            message: message.into(),
            recoverable,
        }
    }

    fn json(&self, revision: u64) -> Value {
        json!({
            "code": self.code,
            "severity": self.severity,
            "revision": revision,
            "sheetId": self.sheet_id,
            "sheetName": self.sheet_name,
            "cell": self.cell,
            "recoverable": self.recoverable,
            "message": self.message,
        })
    }
}

/// Runs the requested checks and returns the inspection report.
pub fn run(project: &Project, revision: u64, checks: Option<&[String]>) -> Result<Value, ServiceError> {
    let requested = match checks {
        None => CHECKS.to_vec(),
        Some(list) => {
            let mut selected = Vec::new();
            for name in list {
                let normalized = name.trim();
                if normalized.is_empty() {
                    continue;
                }
                let Some(check) = CHECKS.iter().find(|check| check.eq_ignore_ascii_case(normalized))
                else {
                    return Err(ServiceError::argument(format!(
                        "unknown check `{normalized}`; expected one of {}",
                        CHECKS.join(", ")
                    )));
                };
                if !selected.contains(check) {
                    selected.push(*check);
                }
            }
            selected
        }
    };

    // Checks always run against freshly evaluated values, so `types`, `cycles`
    // and `references` report what a reader would actually see.
    let mut evaluated = project.clone();
    crate::eval::recalculate(&mut evaluated);
    let project = &evaluated;

    let mut issues = Vec::new();
    for check in requested {
        match check {
            "formulas" => check_formulas(project, &mut issues),
            "references" => check_references(project, &mut issues),
            "cycles" => check_cycles(project, &mut issues),
            "types" => check_types(project, &mut issues),
            "ranges" => check_ranges(project, &mut issues),
            "styles" => check_styles(project, &mut issues),
            _ => {}
        }
    }

    let passed = issues.iter().all(|issue| issue.severity != "error");
    Ok(json!({
        "revision": revision,
        "checkerVersion": CHECKER_VERSION,
        "issues": issues.iter().map(|issue| issue.json(revision)).collect::<Vec<_>>(),
        "passed": passed,
    }))
}

fn check_formulas(project: &Project, issues: &mut Vec<Issue>) {
    for sheet in &project.sheets {
        for cell in &sheet.cells {
            let Some(formula) = &cell.formula
            else {
                continue;
            };
            let trimmed = formula.trim();
            if cell.computed_state == Some(ComputedState::Unsupported) {
                issues.push(Issue::new(
                    "formulas.unsupported-function",
                    "warning",
                    sheet,
                    cell.reference.clone(),
                    "formula uses a function this engine does not evaluate; the imported cached value is kept",
                    false,
                ));
            }
            if !trimmed.starts_with('=') {
                issues.push(Issue::new(
                    "inspect.formulas.missing-equals",
                    "error",
                    sheet,
                    cell.reference.clone(),
                    format!("formula `{trimmed}` does not start with `=`"),
                    true,
                ));
                continue;
            }
            let body = trimmed[1..].trim();
            if body.is_empty() {
                issues.push(Issue::new(
                    "inspect.formulas.empty",
                    "error",
                    sheet,
                    cell.reference.clone(),
                    "formula body is empty",
                    true,
                ));
                continue;
            }
            let mut depth: i32 = 0;
            let mut quotes = 0i32;
            let mut characters = body.chars().peekable();
            let mut previous: Option<char> = None;
            while let Some(character) = characters.next() {
                match character {
                    '"' => quotes += 1,
                    '(' => {
                        if quotes % 2 == 0 {
                            depth += 1;
                            let mut lookahead = characters.clone();
                            let next = lookahead.find(|candidate| !candidate.is_whitespace());
                            if next == Some(')') {
                                let before = previous.unwrap_or('=');
                                if !(before.is_ascii_alphanumeric() || before == '_' || before == '.' || before == '$') {
                                    issues.push(Issue::new(
                                        "inspect.formulas.empty-function",
                                        "error",
                                        sheet,
                                        cell.reference.clone(),
                                        "formula calls a function without a name",
                                        true,
                                    ));
                                }
                            }
                        }
                    }
                    ')' if quotes % 2 == 0 => {
                        depth -= 1;
                    }
                    ')' => {}
                    _ => {}
                }
                previous = Some(character);
            }
            if quotes % 2 != 0 {
                issues.push(Issue::new(
                    "inspect.formulas.unbalanced-quotes",
                    "error",
                    sheet,
                    cell.reference.clone(),
                    "formula has an unterminated string literal",
                    true,
                ));
            }
            if depth != 0 {
                issues.push(Issue::new(
                    "inspect.formulas.unbalanced-parentheses",
                    "error",
                    sheet,
                    cell.reference.clone(),
                    format!("formula parentheses are not balanced (depth {depth})"),
                    true,
                ));
            }
            if let Err(message) = crate::eval::validate(trimmed) {
                issues.push(Issue::new(
                    "inspect.formulas.syntax",
                    "error",
                    sheet,
                    cell.reference.clone(),
                    format!("formula cannot be parsed: {message}"),
                    true,
                ));
            }
        }
    }
}

/// One A1 reference found in a formula body.
struct ScannedReference {
    sheet: Option<String>,
    row: u32,
    column: u32,
}

fn check_references(project: &Project, issues: &mut Vec<Issue>) {
    for sheet in &project.sheets {
        for cell in &sheet.cells {
            if let Some(CellValue::Error { code }) = &cell.computed
                && code == "#REF!"
            {
                issues.push(Issue::new(
                    "inspect.references.unresolved",
                    "error",
                    sheet,
                    cell.reference.clone(),
                    "formula resolves to `#REF!`; a sheet or cell reference no longer exists",
                    true,
                ));
            }
            let Some(formula) = &cell.formula
            else {
                continue;
            };
            for reference in scan_references(formula) {
                let target = match &reference.sheet {
                    Some(name) => match project.sheet(name) {
                        Some(target) => target,
                        None => {
                            issues.push(Issue::new(
                                "inspect.references.missing-sheet",
                                "error",
                                sheet,
                                cell.reference.clone(),
                                format!("formula references sheet `{name}`, which does not exist"),
                                true,
                            ));
                            continue;
                        }
                    },
                    None => sheet,
                };
                if reference.row > MAX_ROW || reference.column > MAX_COLUMN {
                    issues.push(Issue::new(
                        "inspect.references.invalid",
                        "error",
                        sheet,
                        cell.reference.clone(),
                        format!(
                            "formula references `{}`, which is outside the sheet grid",
                            crate::reference::format_reference(reference.row, reference.column)
                        ),
                        true,
                    ));
                    continue;
                }
                if let Some(bounds) = target.bounds()
                    && (reference.row > bounds.end_row || reference.column > bounds.end_column)
                {
                    issues.push(Issue::new(
                        "inspect.references.out-of-range",
                        "warning",
                        sheet,
                        cell.reference.clone(),
                        format!(
                            "formula references `{}!{}`, which is outside the sheet bounds `{}`",
                            target.name,
                            crate::reference::format_reference(reference.row, reference.column),
                            bounds.reference()
                        ),
                        false,
                    ));
                }
            }
        }
    }
}

fn check_cycles(project: &Project, issues: &mut Vec<Issue>) {
    for cycle in crate::eval::find_cycles(project) {
        let Some((sheet_index, row, column)) = cycle.first().copied()
        else {
            continue;
        };
        let sheet = &project.sheets[sheet_index];
        let address = crate::reference::format_reference(row, column);
        let path = cycle
            .iter()
            .map(|(index, row, column)| {
                format!("{}!{}", project.sheets[*index].name, crate::reference::format_reference(*row, *column))
            })
            .collect::<Vec<_>>()
            .join(" -> ");
        issues.push(Issue::new(
            "inspect.cycles.cycle",
            "error",
            sheet,
            address,
            format!("formula cycle detected: {path}"),
            true,
        ));
    }
}

fn check_types(project: &Project, issues: &mut Vec<Issue>) {
    for sheet in &project.sheets {
        for cell in &sheet.cells {
            if let CellValue::Error { code } = &cell.value {
                issues.push(Issue::new(
                    "inspect.types.error-value",
                    "error",
                    sheet,
                    cell.reference.clone(),
                    format!("cell holds the error value `{code}`"),
                    true,
                ));
            }
            if cell.formula.is_some() {
                if cell.computed_state == Some(ComputedState::Unsupported) {
                    // The cached value stays the effective value, so an
                    // unsupported formula is a warning, never an error.
                    continue;
                }
                match &cell.computed {
                    Some(CellValue::Error { code }) => issues.push(Issue::new(
                        "inspect.types.error-value",
                        "error",
                        sheet,
                        cell.reference.clone(),
                        format!("formula result is the error value `{code}`"),
                        true,
                    )),
                    None => issues.push(Issue::new(
                        "inspect.types.formula-without-cache",
                        "warning",
                        sheet,
                        cell.reference.clone(),
                        "formula cell has no evaluated result",
                        false,
                    )),
                    _ => {}
                }
            }
        }
    }
}

fn check_ranges(project: &Project, issues: &mut Vec<Issue>) {
    for sheet in &project.sheets {
        let actual = sheet.actual_range();
        if let Some(dimension) = sheet.dimension
            && let Some(actual) = actual
            && (dimension.end_row < actual.end_row || dimension.end_column < actual.end_column)
        {
            issues.push(Issue::new(
                "inspect.ranges.dimension-too-small",
                "warning",
                sheet,
                dimension.reference(),
                format!(
                    "declared dimension `{}` is smaller than the used range `{}`",
                    dimension.reference(),
                    actual.reference()
                ),
                false,
            ));
        }

        for (index, range) in sheet.merged.iter().enumerate() {
            if !range.is_valid() {
                issues.push(Issue::new(
                    "inspect.ranges.merge-invalid",
                    "error",
                    sheet,
                    range.reference(),
                    format!("merged range `{}` has an invalid extent", range.reference()),
                    true,
                ));
                continue;
            }
            let bounds = sheet.bounds();
            if let Some(bounds) = bounds
                && (range.end_row > bounds.end_row || range.end_column > bounds.end_column)
            {
                issues.push(Issue::new(
                    "inspect.ranges.merge-out-of-bounds",
                    "error",
                    sheet,
                    range.reference(),
                    format!("merged range `{}` extends past the sheet bounds", range.reference()),
                    true,
                ));
            }
            for other in sheet.merged.iter().skip(index + 1) {
                if range.overlaps(other) && range != other {
                    issues.push(Issue::new(
                        "inspect.ranges.merge-overlap",
                        "error",
                        sheet,
                        range.reference(),
                        format!("merged ranges `{}` and `{}` overlap", range.reference(), other.reference()),
                        true,
                    ));
                }
            }
        }
    }
}

fn check_styles(project: &Project, issues: &mut Vec<Issue>) {
    for sheet in &project.sheets {
        for cell in &sheet.cells {
            let Some(style) = &cell.style
            else {
                continue;
            };
            if let Some(size) = style.font_size
                && !(1.0..=409.0).contains(&size)
            {
                issues.push(Issue::new(
                    "inspect.styles.font-size",
                    "error",
                    sheet,
                    cell.reference.clone(),
                    format!("font size {size} is outside the range 1..=409"),
                    true,
                ));
            }
            if let Some(format) = &style.number_format
                && !plausible_number_format(format)
            {
                issues.push(Issue::new(
                    "inspect.styles.unknown-number-format",
                    "error",
                    sheet,
                    cell.reference.clone(),
                    format!("number format `{format}` is not recognized"),
                    true,
                ));
            }
            if let Some(align) = &style.align
                && !matches!(align.as_str(), "general" | "left" | "center" | "right" | "justify" | "distribute" | "fill")
            {
                issues.push(Issue::new(
                    "inspect.styles.align",
                    "error",
                    sheet,
                    cell.reference.clone(),
                    format!("alignment `{align}` is not a valid horizontal alignment"),
                    true,
                ));
            }
        }
    }
}

fn plausible_number_format(format: &str) -> bool {
    let format = format.trim();
    if format.is_empty() {
        return false;
    }
    if format.eq_ignore_ascii_case("general") || format == "@" {
        return true;
    }
    let lower = format.to_ascii_lowercase();
    lower.contains(['0', '#', '?'])
        || lower.contains("yy")
        || lower.contains("dd")
        || lower.contains("hh")
        || lower.contains("ss")
        || lower.contains("am/pm")
        || lower.contains("a/p")
        || lower.contains('%')
}

/// Scans a formula for A1 references, sheet qualifiers and range endpoints.
fn scan_references(source: &str) -> Vec<ScannedReference> {
    let characters: Vec<char> = source.chars().collect();
    let mut references = Vec::new();
    let mut index = 0usize;
    while index < characters.len() {
        let character = characters[index];
        if character == '"' {
            index += 1;
            while index < characters.len() {
                if characters[index] == '"' {
                    if index + 1 < characters.len() && characters[index + 1] == '"' {
                        index += 2;
                    }
                    else {
                        index += 1;
                        break;
                    }
                }
                else {
                    index += 1;
                }
            }
            continue;
        }
        if character == '\'' {
            let mut name = String::new();
            index += 1;
            while index < characters.len() {
                if characters[index] == '\'' {
                    if index + 1 < characters.len() && characters[index + 1] == '\'' {
                        name.push('\'');
                        index += 2;
                    }
                    else {
                        index += 1;
                        break;
                    }
                }
                else {
                    name.push(characters[index]);
                    index += 1;
                }
            }
            if index < characters.len() && characters[index] == '!' {
                index += 1;
                if let Some((row, column, next)) = parse_cell_at(&characters, index) {
                    references.push(ScannedReference { sheet: Some(name), row, column });
                    index = next;
                }
            }
            continue;
        }
        if character.is_ascii_alphabetic() || character == '$' {
            let (token, next) = read_token(&characters, index);
            if next < characters.len() && characters[next] == '(' {
                index = next;
                continue;
            }
            if next < characters.len() && characters[next] == '!' {
                let name = token.trim_start_matches('$').to_owned();
                if let Some((row, column, after)) = parse_cell_at(&characters, next + 1) {
                    references.push(ScannedReference { sheet: Some(name.clone()), row, column });
                    index = after;
                    if index < characters.len()
                        && characters[index] == ':'
                        && let Some((end_row, end_column, after)) = parse_cell_at(&characters, index + 1)
                    {
                        references.push(ScannedReference { sheet: Some(name), row: end_row, column: end_column });
                        index = after;
                    }
                }
                else {
                    index = next + 1;
                }
                continue;
            }
            if let Some((row, column)) = parse_reference(&token) {
                references.push(ScannedReference { sheet: None, row, column });
                index = next;
                if index < characters.len()
                    && characters[index] == ':'
                    && let Some((end_row, end_column, after)) = parse_cell_at(&characters, index + 1)
                {
                    references.push(ScannedReference { sheet: None, row: end_row, column: end_column });
                    index = after;
                }
                continue;
            }
            index = next.max(index + 1);
            continue;
        }
        index += 1;
    }
    references
}

fn read_token(characters: &[char], start: usize) -> (String, usize) {
    let mut index = start;
    let mut token = String::new();
    while index < characters.len() {
        let character = characters[index];
        if character.is_ascii_alphanumeric() || character == '$' || character == '_' || character == '.' {
            token.push(character);
            index += 1;
        }
        else {
            break;
        }
    }
    (token, index)
}

fn parse_cell_at(characters: &[char], start: usize) -> Option<(u32, u32, usize)> {
    let mut index = start;
    let mut token = String::new();
    while index < characters.len() {
        let character = characters[index];
        if character.is_ascii_alphanumeric() || character == '$' {
            token.push(character);
            index += 1;
        }
        else {
            break;
        }
    }
    let (row, column) = parse_reference(&token)?;
    Some((row, column, index))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{CellRange, Project};

    #[test]
    fn detects_direct_cycles() {
        let mut project = Project::blank("Cycle");
        let sheet = &mut project.sheets[0];
        sheet.write_cell(1, 1, CellValue::Empty, Some("=B1".to_owned()), None);
        sheet.write_cell(1, 2, CellValue::Empty, Some("=A1".to_owned()), None);
        sheet.dimension = sheet.actual_range();
        let report = run(&project, 0, Some(&["cycles".to_owned()])).expect("report");
        assert!(!report["passed"].as_bool().unwrap());
        let issues = report["issues"].as_array().unwrap();
        assert!(issues.iter().any(|issue| issue["code"] == "inspect.cycles.cycle"));
    }

    #[test]
    fn detects_references_outside_the_dimension() {
        let mut project = Project::blank("Refs");
        let sheet = &mut project.sheets[0];
        sheet.write_cell(1, 1, CellValue::number(2.0), Some("=C9*2".to_owned()), Some(CellValue::number(4.0)));
        sheet.dimension = Some(CellRange { start_row: 1, end_row: 1, start_column: 1, end_column: 1 });
        let report = run(&project, 0, Some(&["references".to_owned()])).expect("report");
        let issues = report["issues"].as_array().unwrap();
        assert!(issues.iter().any(|issue| issue["code"] == "inspect.references.out-of-range"));
    }

    #[test]
    fn detects_overlapping_merges() {
        let mut project = Project::blank("Merges");
        let sheet = &mut project.sheets[0];
        sheet.merged = vec![
            CellRange { start_row: 1, end_row: 2, start_column: 1, end_column: 2 },
            CellRange { start_row: 2, end_row: 3, start_column: 2, end_column: 3 },
        ];
        let report = run(&project, 0, Some(&["ranges".to_owned()])).expect("report");
        let issues = report["issues"].as_array().unwrap();
        assert!(issues.iter().any(|issue| issue["code"] == "inspect.ranges.merge-overlap"));
    }

    #[test]
    fn reports_formula_shape_problems() {
        let mut project = Project::blank("Formulas");
        let sheet = &mut project.sheets[0];
        sheet.write_cell(1, 1, CellValue::Empty, Some("SUM(A1:A2)".to_owned()), None);
        sheet.write_cell(1, 2, CellValue::Empty, Some("=SUM(".to_owned()), None);
        sheet.write_cell(1, 3, CellValue::Empty, Some("=()".to_owned()), None);
        let report = run(&project, 0, Some(&["formulas".to_owned()])).expect("report");
        let codes: Vec<_> =
            report["issues"].as_array().unwrap().iter().map(|issue| issue["code"].as_str().unwrap().to_owned()).collect();
        assert!(codes.contains(&"inspect.formulas.missing-equals".to_owned()));
        assert!(codes.contains(&"inspect.formulas.unbalanced-parentheses".to_owned()));
        assert!(codes.contains(&"inspect.formulas.empty-function".to_owned()));
    }

    #[test]
    fn reports_error_values_from_literals_and_evaluated_formulas() {
        let mut project = Project::blank("Types");
        let sheet = &mut project.sheets[0];
        sheet.write_cell(1, 1, CellValue::error("#DIV/0!"), None, None);
        sheet.write_cell(1, 2, CellValue::Empty, Some("=A1".to_owned()), None);
        let report = run(&project, 0, Some(&["types".to_owned()])).expect("report");
        let codes: Vec<_> =
            report["issues"].as_array().unwrap().iter().map(|issue| issue["code"].as_str().unwrap().to_owned()).collect();
        assert_eq!(codes.iter().filter(|code| *code == "inspect.types.error-value").count(), 2);
        assert!(!codes.contains(&"inspect.types.formula-without-cache".to_owned()));
    }
}
