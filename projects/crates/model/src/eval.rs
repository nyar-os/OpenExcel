//! Formula evaluation bound to the workbook project.
//!
//! The lexer, parser, dependency order and function library live in
//! `document-spreadsheet-formula`. This module maps the project onto the
//! engine's read-only workbook view and writes the results back onto the cells,
//! folding a handful of legacy and dotted Excel spellings onto engine functions
//! before the formula reaches the engine.

use document_spreadsheet_formula::{CellSnapshot, CellValue as SharedValue, FormulaCell, Workbook};

use crate::model::{Cell, CellValue, ComputedState, Project};

/// Identifies one cell across sheets.
pub(crate) type CellKey = document_spreadsheet_formula::CellKey;

/// Function names the engine understands, reported through capabilities.
pub(crate) use document_spreadsheet_formula::FUNCTIONS;

/// Function names this project folds onto an engine function before the formula
/// reaches the engine, so legacy and dotted Excel spellings compute.
pub(crate) const FUNCTION_ALIASES: &[(&str, &str)] = &[
    ("CONCATENATE", "CONCAT"),
    ("MODE.SNGL", "MODE"),
    ("STDEV.S", "STDEV"),
    ("STDEV.P", "STDEVP"),
    ("VAR.S", "VAR"),
    ("VAR.P", "VARP"),
    ("RANK.EQ", "RANK"),
];

/// Every function name a stored formula may call: engine functions plus aliases.
pub(crate) fn supported_functions() -> Vec<&'static str> {
    FUNCTIONS.iter().copied().chain(FUNCTION_ALIASES.iter().map(|(alias, _)| *alias)).collect()
}

/// Rewrites alias function calls to the engine spelling. String literals,
/// quoted sheet names and external-reference brackets are copied verbatim, so
/// only real call sites change; the stored formula keeps its original text.
pub(crate) fn expand_aliases(source: &str) -> String {
    let chars: Vec<char> = source.chars().collect();
    let mut output = String::with_capacity(source.len());
    let mut index = 0usize;
    let is_name_char = |value: char| value.is_alphanumeric() || matches!(value, '_' | '.' | '\\');
    while index < chars.len() {
        let current = chars[index];
        if current == '"' {
            output.push('"');
            index += 1;
            while index < chars.len() {
                if chars[index] == '"' {
                    output.push('"');
                    index += 1;
                    if index < chars.len() && chars[index] == '"' {
                        output.push('"');
                        index += 1;
                    }
                    else {
                        break;
                    }
                }
                else {
                    output.push(chars[index]);
                    index += 1;
                }
            }
        }
        else if current == '\'' || current == '[' {
            let closing = if current == '\'' { '\'' } else { ']' };
            output.push(current);
            index += 1;
            while index < chars.len() && chars[index] != closing {
                output.push(chars[index]);
                index += 1;
            }
            if index < chars.len() {
                output.push(chars[index]);
                index += 1;
            }
        }
        else if is_name_char(current) {
            let start = index;
            while index < chars.len() && is_name_char(chars[index]) {
                index += 1;
            }
            let token: String = chars[start..index].iter().collect();
            let mut probe = index;
            while probe < chars.len() && chars[probe].is_whitespace() {
                probe += 1;
            }
            if probe < chars.len() && chars[probe] == '(' {
                if let Some((_, target)) = FUNCTION_ALIASES.iter().find(|(alias, _)| alias.eq_ignore_ascii_case(&token)) {
                    output.push_str(target);
                    continue;
                }
            }
            output.push_str(&token);
        }
        else {
            output.push(current);
            index += 1;
        }
    }
    output
}

/// Reads the value a reader should show for one cell.
pub(crate) fn value_of(cell: &Cell) -> CellValue {
    if cell.formula.is_some() {
        cell.computed.clone().or_else(|| cell.cached.clone()).unwrap_or(CellValue::Empty)
    }
    else {
        cell.value.clone()
    }
}

/// Reports whether a formula body parses, without exposing the expression tree.
pub(crate) fn validate(source: &str) -> Result<(), String> {
    document_spreadsheet_formula::validate(&expand_aliases(source))
}

/// Recomputes every formula cell and returns how many computed values changed.
pub(crate) fn recalculate(project: &mut Project) -> usize {
    let report = document_spreadsheet_formula::evaluate(&ProjectView(project));
    let mut changed = 0usize;
    for (sheet_index, sheet) in project.sheets.iter_mut().enumerate() {
        for cell in &mut sheet.cells {
            let key = (sheet_index, cell.row, cell.column);
            let (next, state) = if cell.formula.is_some() {
                let value = report.values.get(&key).map(from_shared).unwrap_or_else(|| CellValue::error("#VALUE!"));
                let state =
                    if report.unsupported.contains(&key) { ComputedState::Unsupported } else { ComputedState::Computed };
                (Some(value), Some(state))
            }
            else {
                (None, None)
            };
            if cell.computed != next {
                changed += 1;
                cell.computed = next;
            }
            cell.computed_state = state;
        }
    }
    changed
}

/// Finds every formula cycle, including cycles that cross sheets.
pub(crate) fn find_cycles(project: &Project) -> Vec<Vec<CellKey>> {
    document_spreadsheet_formula::find_cycles(&ProjectView(project))
}

/// Read-only project view the engine reads through.
struct ProjectView<'a>(&'a Project);

impl Workbook for ProjectView<'_> {
    fn sheet_index(&self, name: &str) -> Option<usize> {
        self.0.sheet_index(name)
    }

    fn cell(&self, sheet: usize, row: u32, column: u32) -> Option<CellSnapshot> {
        let cell = self.0.sheets.get(sheet)?.cell(row, column)?;
        Some(CellSnapshot {
            value: to_shared(&cell.value),
            has_formula: cell.formula.is_some(),
            cached: cell.cached.as_ref().map(to_shared),
        })
    }

    fn formula_cells(&self) -> Vec<FormulaCell> {
        let mut cells = Vec::new();
        for (sheet_index, sheet) in self.0.sheets.iter().enumerate() {
            for cell in &sheet.cells {
                if let Some(source) = &cell.formula {
                    cells.push(FormulaCell {
                        sheet: sheet_index,
                        row: cell.row,
                        column: cell.column,
                        source: expand_aliases(source.trim().trim_start_matches('=')),
                    });
                }
            }
        }
        cells
    }
}

fn to_shared(value: &CellValue) -> SharedValue {
    match value {
        CellValue::Empty => SharedValue::Empty,
        CellValue::Number { number } => SharedValue::number(*number),
        CellValue::Text { text } => SharedValue::text(text),
        CellValue::Bool { value } => SharedValue::boolean(*value),
        CellValue::Error { code } => SharedValue::error(code),
    }
}

fn from_shared(value: &SharedValue) -> CellValue {
    match value {
        SharedValue::Empty => CellValue::Empty,
        SharedValue::Number { number } => CellValue::number(*number),
        SharedValue::Text { text } => CellValue::text(text),
        SharedValue::Bool { value } => CellValue::boolean(*value),
        SharedValue::Error { code } => CellValue::error(code),
    }
}
