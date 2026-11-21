//! Self-contained formula engine: lexer, parser, dependency graph and evaluator.
//!
//! The engine covers the arithmetic, comparison and text operators of the
//! SpreadsheetML grammar plus the function set the product exposes. Formulas are
//! parsed into an expression tree, evaluated in dependency order with a
//! depth-first walk and cached per cell in [`Cell::computed`]. Cycles are
//! detected across sheets and reported as the `#CYCLE!` error value.
//!
//! Numeric cells stay `f64`; integers lose their `.0` through
//! [`crate::model::format_number`] at every projection.

use std::{
    cmp::Ordering,
    collections::{HashMap, HashSet},
};

use crate::{
    model::{Cell, CellRange, CellValue, ComputedState, Project},
    reference::parse_reference,
};

/// Error code of a division by zero.
const DIV0: &str = "#DIV/0!";
/// Error code of an operand or result that cannot be coerced.
const VALUE: &str = "#VALUE!";
/// Error code of a reference that does not resolve.
const REF: &str = "#REF!";
/// Error code of an unknown function or name.
const NAME: &str = "#NAME?";
/// Error code of a lookup that found no match.
const NA: &str = "#N/A";
/// Error code of a formula that depends on itself, directly or through others.
const CYCLE: &str = "#CYCLE!";

/// Functions this engine evaluates, sorted for a stable capability list.
pub(crate) const FUNCTIONS: &[&str] = &[
    "ABS",
    "AND",
    "AVERAGE",
    "CONCAT",
    "COUNT",
    "COUNTA",
    "IF",
    "IFERROR",
    "INDEX",
    "INT",
    "LEFT",
    "LEN",
    "LOWER",
    "MATCH",
    "MAX",
    "MID",
    "MIN",
    "NOT",
    "OR",
    "RIGHT",
    "ROUND",
    "ROUNDDOWN",
    "ROUNDUP",
    "SUBSTITUTE",
    "SUM",
    "TEXTJOIN",
    "TRIM",
    "UPPER",
    "VLOOKUP",
];

const MAX_ROW: u32 = 1_048_576;
const MAX_COLUMN: u32 = 16_384;
/// Upper bound on the cells one range argument may expand into.
const MAX_RANGE_CELLS: u64 = 200_000;

/// One formula cell, identified by sheet index, row and column.
pub(crate) type CellKey = (usize, u32, u32);

// ---------------------------------------------------------------------------
// Expression tree
// ---------------------------------------------------------------------------

/// Binary operators of the supported grammar.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BinOp {
    Add,
    Subtract,
    Multiply,
    Divide,
    Power,
    Concat,
    Equal,
    NotEqual,
    Less,
    LessEqual,
    Greater,
    GreaterEqual,
}

/// One parsed formula body.
#[derive(Debug, Clone, PartialEq)]
enum Expr {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(String),
    Empty,
    /// A bare identifier that is neither a reference nor a boolean.
    Name(String),
    Cell {
        sheet: Option<String>,
        row: u32,
        column: u32,
    },
    Range {
        sheet: Option<String>,
        range: CellRange,
    },
    Negate(Box<Expr>),
    Percent(Box<Expr>),
    Binary {
        op: BinOp,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    Call {
        name: String,
        args: Vec<Expr>,
    },
}

// ---------------------------------------------------------------------------
// Lexer
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq)]
enum Token {
    Number(f64),
    Text(String),
    Ident(String),
    ErrorCode(String),
    Sheet(String),
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Amp,
    Percent,
    Equal,
    NotEqual,
    Less,
    LessEqual,
    Greater,
    GreaterEqual,
    LParen,
    RParen,
    Comma,
    Colon,
    Bang,
}

fn lex(source: &str) -> Result<Vec<Token>, String> {
    let characters: Vec<char> = source.chars().collect();
    let mut tokens = Vec::new();
    let mut index = 0usize;
    while index < characters.len() {
        let character = characters[index];
        if character.is_whitespace() {
            index += 1;
            continue;
        }
        match character {
            '+' => {
                tokens.push(Token::Plus);
                index += 1;
            }
            '-' => {
                tokens.push(Token::Minus);
                index += 1;
            }
            '*' => {
                tokens.push(Token::Star);
                index += 1;
            }
            '/' => {
                tokens.push(Token::Slash);
                index += 1;
            }
            '^' => {
                tokens.push(Token::Caret);
                index += 1;
            }
            '&' => {
                tokens.push(Token::Amp);
                index += 1;
            }
            '%' => {
                tokens.push(Token::Percent);
                index += 1;
            }
            '(' => {
                tokens.push(Token::LParen);
                index += 1;
            }
            ')' => {
                tokens.push(Token::RParen);
                index += 1;
            }
            ',' => {
                tokens.push(Token::Comma);
                index += 1;
            }
            ':' => {
                tokens.push(Token::Colon);
                index += 1;
            }
            '!' => {
                tokens.push(Token::Bang);
                index += 1;
            }
            '=' => {
                tokens.push(Token::Equal);
                index += 1;
            }
            '<' => match characters.get(index + 1) {
                Some('=') => {
                    tokens.push(Token::LessEqual);
                    index += 2;
                }
                Some('>') => {
                    tokens.push(Token::NotEqual);
                    index += 2;
                }
                _ => {
                    tokens.push(Token::Less);
                    index += 1;
                }
            },
            '>' => match characters.get(index + 1) {
                Some('=') => {
                    tokens.push(Token::GreaterEqual);
                    index += 2;
                }
                _ => {
                    tokens.push(Token::Greater);
                    index += 1;
                }
            },
            '"' => {
                let (text, next) = read_quoted(&characters, index + 1, '"', "string literal")?;
                tokens.push(Token::Text(text));
                index = next;
            }
            '\'' => {
                let (name, next) = read_quoted(&characters, index + 1, '\'', "sheet name")?;
                tokens.push(Token::Sheet(name));
                index = next;
            }
            '#' => {
                let start = index;
                index += 1;
                while index < characters.len()
                    && (characters[index].is_ascii_alphanumeric() || matches!(characters[index], '/' | '!' | '?' | '_' | '&'))
                {
                    index += 1;
                }
                tokens.push(Token::ErrorCode(characters[start..index].iter().collect()));
            }
            _ if character.is_ascii_digit()
                || (character == '.' && characters.get(index + 1).is_some_and(char::is_ascii_digit)) =>
            {
                let start = index;
                let mut seen_dot = false;
                while index < characters.len() {
                    let current = characters[index];
                    if current.is_ascii_digit() {
                        index += 1;
                    }
                    else if current == '.' && !seen_dot {
                        seen_dot = true;
                        index += 1;
                    }
                    else if (current == 'e' || current == 'E') && exponent_length(&characters, index + 1).is_some() {
                        index += 1 + exponent_length(&characters, index + 1).unwrap_or(0);
                    }
                    else {
                        break;
                    }
                }
                let literal: String = characters[start..index].iter().collect();
                let number = literal.parse::<f64>().map_err(|_| format!("`{literal}` is not a number"))?;
                tokens.push(Token::Number(number));
            }
            _ if character.is_alphabetic() || character == '_' || character == '$' => {
                let start = index;
                while index < characters.len()
                    && (characters[index].is_alphanumeric() || matches!(characters[index], '_' | '.' | '$'))
                {
                    index += 1;
                }
                tokens.push(Token::Ident(characters[start..index].iter().collect()));
            }
            other => return Err(format!("unexpected character `{other}`")),
        }
    }
    Ok(tokens)
}

/// Length of an exponent suffix (`3`, `+3`, `-3`) starting at `start`.
fn exponent_length(characters: &[char], start: usize) -> Option<usize> {
    let first = *characters.get(start)?;
    if first.is_ascii_digit() {
        return Some(1);
    }
    if (first == '+' || first == '-') && characters.get(start + 1).is_some_and(char::is_ascii_digit) {
        return Some(2);
    }
    None
}

/// Reads a delimited literal, honoring the doubled delimiter escape.
fn read_quoted(characters: &[char], start: usize, delimiter: char, label: &str) -> Result<(String, usize), String> {
    let mut index = start;
    let mut text = String::new();
    while index < characters.len() {
        if characters[index] == delimiter {
            if characters.get(index + 1) == Some(&delimiter) {
                text.push(delimiter);
                index += 2;
            }
            else {
                return Ok((text, index + 1));
            }
        }
        else {
            text.push(characters[index]);
            index += 1;
        }
    }
    Err(format!("unterminated {label}"))
}

// ---------------------------------------------------------------------------
// Parser
// ---------------------------------------------------------------------------

struct Parser {
    tokens: Vec<Token>,
    index: usize,
}

impl Parser {
    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.index)
    }

    fn advance(&mut self) -> Option<Token> {
        let token = self.tokens.get(self.index).cloned();
        if token.is_some() {
            self.index += 1;
        }
        token
    }

    fn eat(&mut self, token: &Token) -> bool {
        if self.peek() == Some(token) {
            self.index += 1;
            true
        }
        else {
            false
        }
    }

    fn expression(&mut self) -> Result<Expr, String> {
        self.comparison()
    }

    fn comparison(&mut self) -> Result<Expr, String> {
        let mut left = self.concat()?;
        loop {
            let op = match self.peek() {
                Some(Token::Equal) => BinOp::Equal,
                Some(Token::NotEqual) => BinOp::NotEqual,
                Some(Token::Less) => BinOp::Less,
                Some(Token::LessEqual) => BinOp::LessEqual,
                Some(Token::Greater) => BinOp::Greater,
                Some(Token::GreaterEqual) => BinOp::GreaterEqual,
                _ => break,
            };
            self.index += 1;
            let right = self.concat()?;
            left = Expr::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn concat(&mut self) -> Result<Expr, String> {
        let mut left = self.additive()?;
        while self.eat(&Token::Amp) {
            let right = self.additive()?;
            left = Expr::Binary { op: BinOp::Concat, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn additive(&mut self) -> Result<Expr, String> {
        let mut left = self.term()?;
        loop {
            let op = match self.peek() {
                Some(Token::Plus) => BinOp::Add,
                Some(Token::Minus) => BinOp::Subtract,
                _ => break,
            };
            self.index += 1;
            let right = self.term()?;
            left = Expr::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn term(&mut self) -> Result<Expr, String> {
        let mut left = self.power()?;
        loop {
            let op = match self.peek() {
                Some(Token::Star) => BinOp::Multiply,
                Some(Token::Slash) => BinOp::Divide,
                _ => break,
            };
            self.index += 1;
            let right = self.power()?;
            left = Expr::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn power(&mut self) -> Result<Expr, String> {
        let mut left = self.unary()?;
        while self.eat(&Token::Caret) {
            let right = self.unary()?;
            left = Expr::Binary { op: BinOp::Power, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn unary(&mut self) -> Result<Expr, String> {
        match self.peek() {
            Some(Token::Minus) => {
                self.index += 1;
                Ok(Expr::Negate(Box::new(self.unary()?)))
            }
            Some(Token::Plus) => {
                self.index += 1;
                self.unary()
            }
            _ => self.postfix(),
        }
    }

    fn postfix(&mut self) -> Result<Expr, String> {
        let mut expression = self.primary()?;
        while self.eat(&Token::Percent) {
            expression = Expr::Percent(Box::new(expression));
        }
        Ok(expression)
    }

    fn primary(&mut self) -> Result<Expr, String> {
        let token = self.advance().ok_or_else(|| "the formula ends unexpectedly".to_owned())?;
        match token {
            Token::Number(number) => Ok(Expr::Number(number)),
            Token::Text(text) => Ok(Expr::Text(text)),
            Token::ErrorCode(code) => Ok(Expr::Error(code)),
            Token::LParen => {
                let inner = self.expression()?;
                if !self.eat(&Token::RParen) {
                    return Err("a `(` is not closed".to_owned());
                }
                Ok(inner)
            }
            Token::Sheet(name) => {
                if !self.eat(&Token::Bang) {
                    return Err("a quoted sheet name needs `!`".to_owned());
                }
                self.reference(Some(name))
            }
            Token::Ident(name) => {
                if self.eat(&Token::LParen) {
                    return self.call(name);
                }
                if self.eat(&Token::Bang) {
                    return self.reference(Some(name));
                }
                if name.eq_ignore_ascii_case("TRUE") {
                    return Ok(Expr::Bool(true));
                }
                if name.eq_ignore_ascii_case("FALSE") {
                    return Ok(Expr::Bool(false));
                }
                self.reference_from(None, &name)
            }
            other => Err(format!("unexpected token `{}`", label(&other))),
        }
    }

    /// Reads a cell or range reference after an optional sheet qualifier.
    fn reference(&mut self, sheet: Option<String>) -> Result<Expr, String> {
        let token = self.advance().ok_or_else(|| "a sheet qualifier needs a cell reference".to_owned())?;
        match token {
            Token::Ident(name) => self.reference_from(sheet, &name),
            Token::Sheet(other) => {
                if !self.eat(&Token::Bang) {
                    return Err("a quoted sheet name needs `!`".to_owned());
                }
                let next = self.advance().ok_or_else(|| "a sheet qualifier needs a cell reference".to_owned())?;
                match next {
                    Token::Ident(name) => self.reference_from(sheet.or(Some(other)), &name),
                    other => Err(format!("unexpected token `{}` after a sheet qualifier", label(&other))),
                }
            }
            other => Err(format!("unexpected token `{}` after `!`", label(&other))),
        }
    }

    fn reference_from(&mut self, sheet: Option<String>, name: &str) -> Result<Expr, String> {
        let Some((row, column)) = parse_reference(name)
        else {
            // A bare identifier is a name, which evaluates to `#NAME?`.
            return Ok(Expr::Name(name.to_owned()));
        };
        if !self.eat(&Token::Colon) {
            return Ok(Expr::Cell { sheet, row, column });
        }
        let token = self.advance().ok_or_else(|| "a range needs an end reference".to_owned())?;
        let end = match token {
            Token::Ident(name) => name,
            Token::Sheet(_) => {
                if !self.eat(&Token::Bang) {
                    return Err("a quoted sheet name needs `!`".to_owned());
                }
                match self.advance() {
                    Some(Token::Ident(name)) => name,
                    other => {
                        return Err(format!(
                            "unexpected token `{}` at the end of a range",
                            other.as_ref().map_or_else(|| "end of formula".to_owned(), label)
                        ));
                    }
                }
            }
            other => return Err(format!("unexpected token `{}` at the end of a range", label(&other))),
        };
        let (end_row, end_column) = parse_reference(&end).ok_or_else(|| format!("`{end}` is not a cell reference"))?;
        Ok(Expr::Range { sheet, range: CellRange::between(row, column, end_row, end_column) })
    }

    fn call(&mut self, name: String) -> Result<Expr, String> {
        let mut args = Vec::new();
        if self.eat(&Token::RParen) {
            return Ok(Expr::Call { name: name.to_ascii_uppercase(), args });
        }
        loop {
            if self.eat(&Token::Comma) {
                args.push(Expr::Empty);
                continue;
            }
            if self.eat(&Token::RParen) {
                args.push(Expr::Empty);
                break;
            }
            args.push(self.expression()?);
            if self.eat(&Token::Comma) {
                continue;
            }
            if !self.eat(&Token::RParen) {
                return Err("a function call is not closed with `)`".to_owned());
            }
            break;
        }
        Ok(Expr::Call { name: name.to_ascii_uppercase(), args })
    }
}

fn label(token: &Token) -> String {
    match token {
        Token::Number(number) => number.to_string(),
        Token::Text(text) => format!("\"{text}\""),
        Token::Ident(name) => name.clone(),
        Token::ErrorCode(code) => code.clone(),
        Token::Sheet(name) => format!("'{name}'"),
        Token::Plus => "+".to_owned(),
        Token::Minus => "-".to_owned(),
        Token::Star => "*".to_owned(),
        Token::Slash => "/".to_owned(),
        Token::Caret => "^".to_owned(),
        Token::Amp => "&".to_owned(),
        Token::Percent => "%".to_owned(),
        Token::Equal => "=".to_owned(),
        Token::NotEqual => "<>".to_owned(),
        Token::Less => "<".to_owned(),
        Token::LessEqual => "<=".to_owned(),
        Token::Greater => ">".to_owned(),
        Token::GreaterEqual => ">=".to_owned(),
        Token::LParen => "(".to_owned(),
        Token::RParen => ")".to_owned(),
        Token::Comma => ",".to_owned(),
        Token::Colon => ":".to_owned(),
        Token::Bang => "!".to_owned(),
    }
}

/// Parses a formula body, with or without the leading `=`.
fn parse_formula(source: &str) -> Result<Expr, String> {
    let trimmed = source.trim();
    let body = trimmed.strip_prefix('=').unwrap_or(trimmed).trim();
    if body.is_empty() {
        return Err("the formula body is empty".to_owned());
    }
    let tokens = lex(body)?;
    let mut parser = Parser { tokens, index: 0 };
    let expression = parser.expression()?;
    if parser.peek().is_some() {
        return Err(format!("unexpected trailing input in `{body}`"));
    }
    Ok(expression)
}

// ---------------------------------------------------------------------------
// Evaluation
// ---------------------------------------------------------------------------

fn error(code: &str) -> CellValue {
    CellValue::error(code)
}

fn normalize_zero(number: f64) -> f64 {
    if number == 0.0 { 0.0 } else { number }
}

/// Coerces one value for arithmetic: empty is `0`, text is `#VALUE!`.
fn numeric(value: &CellValue) -> Result<f64, CellValue> {
    match value {
        CellValue::Number { number } => Ok(*number),
        CellValue::Bool { value } => Ok(if *value { 1.0 } else { 0.0 }),
        CellValue::Empty => Ok(0.0),
        CellValue::Error { .. } => Err(value.clone()),
        CellValue::Text { .. } => Err(error(VALUE)),
    }
}

/// Coerces one value for text concatenation.
fn as_text(value: &CellValue) -> Result<String, CellValue> {
    match value {
        CellValue::Text { text } => Ok(text.clone()),
        CellValue::Number { number } => Ok(crate::model::format_number(*number)),
        CellValue::Bool { value } => Ok(if *value { "TRUE".to_owned() } else { "FALSE".to_owned() }),
        CellValue::Empty => Ok(String::new()),
        CellValue::Error { .. } => Err(value.clone()),
    }
}

fn truthy(value: &CellValue) -> Result<bool, CellValue> {
    match value {
        CellValue::Bool { value } => Ok(*value),
        CellValue::Number { number } => Ok(*number != 0.0),
        CellValue::Empty => Ok(false),
        CellValue::Error { .. } => Err(value.clone()),
        CellValue::Text { .. } => Err(error(VALUE)),
    }
}

fn rank(value: &CellValue) -> u8 {
    match value {
        CellValue::Empty | CellValue::Number { .. } => 0,
        CellValue::Text { .. } => 1,
        CellValue::Bool { .. } => 2,
        CellValue::Error { .. } => 3,
    }
}

/// Orders two values the way a comparison operator does.
fn compare(left: &CellValue, right: &CellValue) -> Result<Ordering, CellValue> {
    match (left, right) {
        (CellValue::Error { .. }, _) => Err(left.clone()),
        (_, CellValue::Error { .. }) => Err(right.clone()),
        (CellValue::Empty, CellValue::Empty) => Ok(Ordering::Equal),
        (CellValue::Text { text: left }, CellValue::Text { text: right }) => Ok(left.to_lowercase().cmp(&right.to_lowercase())),
        (CellValue::Text { text }, CellValue::Empty) => Ok(text.to_lowercase().cmp(&String::new())),
        (CellValue::Empty, CellValue::Text { text }) => Ok(String::new().cmp(&text.to_lowercase())),
        (
            CellValue::Empty | CellValue::Number { .. } | CellValue::Bool { .. },
            CellValue::Empty | CellValue::Number { .. } | CellValue::Bool { .. },
        ) => {
            let left = numeric(left).unwrap_or(0.0);
            let right = numeric(right).unwrap_or(0.0);
            Ok(left.partial_cmp(&right).unwrap_or(Ordering::Equal))
        }
        _ => Ok(rank(left).cmp(&rank(right))),
    }
}

/// One evaluated function argument.
enum Argument {
    Value(CellValue),
    Grid(Grid),
    /// Resolution failed before any cell could be read.
    Invalid(CellValue),
}

/// A rectangular range flattened row by row, together with its shape.
struct Grid {
    values: Vec<CellValue>,
    rows: u32,
    columns: u32,
}

impl Grid {
    fn value(&self, row: u32, column: u32) -> Option<&CellValue> {
        if row == 0 || column == 0 || row > self.rows || column > self.columns {
            return None;
        }
        self.values.get(((row - 1) * self.columns + (column - 1)) as usize)
    }
}

struct Evaluator<'a> {
    project: &'a Project,
    forms: HashMap<CellKey, Option<Expr>>,
    values: HashMap<CellKey, CellValue>,
    states: HashMap<CellKey, u8>,
    stack: Vec<CellKey>,
    /// Cells whose formula used a function or construct the engine does not
    /// evaluate. Their imported cached value stays the effective value.
    unsupported: HashSet<CellKey>,
}

impl Evaluator<'_> {
    /// Marks the formula cell currently being evaluated as unsupported.
    fn mark_unsupported(&mut self) {
        if let Some(key) = self.stack.last().copied() {
            self.unsupported.insert(key);
        }
    }

    fn resolve_sheet(&self, sheet: Option<&str>, current: usize) -> Option<usize> {
        match sheet {
            None => Some(current),
            Some(name) => self.project.sheet_index(name),
        }
    }

    fn cell_value(&mut self, sheet: usize, row: u32, column: u32) -> CellValue {
        if row == 0 || row > MAX_ROW || column == 0 || column > MAX_COLUMN {
            return error(REF);
        }
        let Some(target) = self.project.sheets.get(sheet)
        else {
            return error(REF);
        };
        match target.cell(row, column) {
            None => CellValue::Empty,
            Some(cell) => match cell.formula {
                Some(_) => self.eval_cell(sheet, row, column),
                None => cell.value.clone(),
            },
        }
    }

    fn eval_cell(&mut self, sheet: usize, row: u32, column: u32) -> CellValue {
        let key = (sheet, row, column);
        match self.states.get(&key).copied() {
            Some(2) => return self.values.get(&key).cloned().unwrap_or(CellValue::Empty),
            Some(1) => {
                if let Some(position) = self.stack.iter().position(|entry| *entry == key) {
                    let members: Vec<CellKey> = self.stack[position..].to_vec();
                    for member in members {
                        self.states.insert(member, 2);
                        self.values.insert(member, error(CYCLE));
                    }
                }
                return error(CYCLE);
            }
            _ => {}
        }
        self.states.insert(key, 1);
        self.stack.push(key);
        let form = self.forms.get(&key).cloned().flatten();
        let evaluated = match form {
            Some(expression) => self.eval(&expression, sheet),
            None => error(VALUE),
        };
        // An unsupported formula keeps the value the source writer cached, so
        // opening and re-exporting a valid workbook never replaces a real
        // result with `#NAME?`.
        let value = if self.unsupported.contains(&key) {
            self.project
                .sheets
                .get(sheet)
                .and_then(|target| target.cell(row, column))
                .and_then(|cell| cell.cached.clone())
                .unwrap_or_else(|| error(NAME))
        }
        else {
            evaluated
        };
        self.stack.pop();
        if self.states.get(&key).copied() == Some(1) {
            self.states.insert(key, 2);
            self.values.insert(key, value.clone());
        }
        value
    }

    fn grid(&mut self, sheet: Option<&str>, range: &CellRange, current: usize) -> Result<Grid, CellValue> {
        let Some(index) = self.resolve_sheet(sheet, current)
        else {
            return Err(error(REF));
        };
        if !range.is_valid() || range.end_row > MAX_ROW || range.end_column > MAX_COLUMN {
            return Err(error(REF));
        }
        let area = u64::from(range.rows()) * u64::from(range.columns());
        if area > MAX_RANGE_CELLS {
            return Err(error(REF));
        }
        let mut values = Vec::with_capacity(area as usize);
        for row in range.start_row..=range.end_row {
            for column in range.start_column..=range.end_column {
                values.push(self.cell_value(index, row, column));
            }
        }
        Ok(Grid { values, rows: range.rows(), columns: range.columns() })
    }

    fn eval(&mut self, expression: &Expr, current: usize) -> CellValue {
        match expression {
            Expr::Number(number) => CellValue::number(*number),
            Expr::Text(text) => CellValue::text(text),
            Expr::Bool(value) => CellValue::boolean(*value),
            Expr::Error(code) => error(code),
            Expr::Empty => CellValue::Empty,
            Expr::Name(_) => {
                self.mark_unsupported();
                error(NAME)
            }
            Expr::Cell { sheet, row, column } => {
                let Some(index) = self.resolve_sheet(sheet.as_deref(), current)
                else {
                    return error(REF);
                };
                self.cell_value(index, *row, *column)
            }
            Expr::Range { .. } => {
                // A range used as one scalar is an array formula or an
                // unimplemented construct. Keep the imported cache instead of
                // replacing it with `#VALUE!`.
                self.mark_unsupported();
                error(VALUE)
            }
            Expr::Negate(inner) => match numeric(&self.eval(inner, current)) {
                Ok(number) => CellValue::number(normalize_zero(-number)),
                Err(error) => error,
            },
            Expr::Percent(inner) => match numeric(&self.eval(inner, current)) {
                Ok(number) => CellValue::number(normalize_zero(number / 100.0)),
                Err(error) => error,
            },
            Expr::Binary { op, left, right } => self.binary(*op, left, right, current),
            Expr::Call { name, args } => self.call(name, args, current),
        }
    }

    fn binary(&mut self, op: BinOp, left: &Expr, right: &Expr, current: usize) -> CellValue {
        let left = self.eval(left, current);
        if let CellValue::Error { .. } = left {
            return left;
        }
        let right = self.eval(right, current);
        if let CellValue::Error { .. } = right {
            return right;
        }
        match op {
            BinOp::Concat => match (as_text(&left), as_text(&right)) {
                (Ok(left), Ok(right)) => CellValue::text(format!("{left}{right}")),
                (Err(error), _) | (_, Err(error)) => error,
            },
            BinOp::Equal | BinOp::NotEqual | BinOp::Less | BinOp::LessEqual | BinOp::Greater | BinOp::GreaterEqual => {
                let ordering = match compare(&left, &right) {
                    Ok(ordering) => ordering,
                    Err(error) => return error,
                };
                let result = match op {
                    BinOp::Equal => ordering == Ordering::Equal,
                    BinOp::NotEqual => ordering != Ordering::Equal,
                    BinOp::Less => ordering == Ordering::Less,
                    BinOp::LessEqual => ordering != Ordering::Greater,
                    BinOp::Greater => ordering == Ordering::Greater,
                    _ => ordering != Ordering::Less,
                };
                CellValue::boolean(result)
            }
            _ => {
                let left = match numeric(&left) {
                    Ok(number) => number,
                    Err(error) => return error,
                };
                let right = match numeric(&right) {
                    Ok(number) => number,
                    Err(error) => return error,
                };
                let result = match op {
                    BinOp::Add => left + right,
                    BinOp::Subtract => left - right,
                    BinOp::Multiply => left * right,
                    BinOp::Divide => {
                        if right == 0.0 {
                            return error(DIV0);
                        }
                        left / right
                    }
                    BinOp::Power => left.powf(right),
                    _ => return error(VALUE),
                };
                if !result.is_finite() {
                    return error(VALUE);
                }
                CellValue::number(normalize_zero(result))
            }
        }
    }

    fn call(&mut self, name: &str, args: &[Expr], current: usize) -> CellValue {
        if !FUNCTIONS.contains(&name) {
            // The engine does not know this function. Mark the cell so the
            // imported cached value survives instead of `#NAME?`.
            self.mark_unsupported();
            return error(NAME);
        }
        match name {
            "IF" => return self.function_if(args, current),
            "IFERROR" => return self.function_iferror(args, current),
            "AND" | "OR" => return self.function_logical(name == "AND", args, current),
            _ => {}
        }
        let mut arguments = Vec::with_capacity(args.len());
        for argument in args {
            arguments.push(match argument {
                Expr::Range { sheet, range } => match self.grid(sheet.as_deref(), range, current) {
                    Ok(grid) => Argument::Grid(grid),
                    Err(error) => Argument::Invalid(error),
                },
                _ => Argument::Value(self.eval(argument, current)),
            });
        }
        apply_function(name, &arguments)
    }

    fn function_if(&mut self, args: &[Expr], current: usize) -> CellValue {
        let Some(condition) = args.first()
        else {
            return error(VALUE);
        };
        if args.len() > 3 {
            return error(VALUE);
        }
        let condition = self.eval(condition, current);
        let truth = match truthy(&condition) {
            Ok(truth) => truth,
            Err(error) => return error,
        };
        let chosen = if truth { args.get(1) } else { args.get(2) };
        match chosen {
            Some(expression) if !matches!(expression, Expr::Empty) => self.eval(expression, current),
            _ => CellValue::boolean(false),
        }
    }

    /// `IFERROR(value, fallback)` evaluates the fallback only for an error.
    fn function_iferror(&mut self, args: &[Expr], current: usize) -> CellValue {
        if args.len() != 2 {
            return error(VALUE);
        }
        let value = self.eval(&args[0], current);
        if matches!(value, CellValue::Error { .. }) { self.eval(&args[1], current) } else { value }
    }

    /// `AND`/`OR` over scalars and range members, short-circuiting on the
    /// deciding value. Text inside a range is ignored, matching a spreadsheet.
    fn function_logical(&mut self, and: bool, args: &[Expr], current: usize) -> CellValue {
        if args.is_empty() {
            return error(VALUE);
        }
        for argument in args {
            match argument {
                Expr::Range { sheet, range } => match self.grid(sheet.as_deref(), range, current) {
                    Ok(grid) => {
                        for value in &grid.values {
                            let truth = match value {
                                CellValue::Bool { value } => Some(*value),
                                CellValue::Number { number } => Some(*number != 0.0),
                                CellValue::Error { .. } => return value.clone(),
                                _ => None,
                            };
                            if let Some(truth) = truth
                                && truth != and
                            {
                                return CellValue::boolean(!and);
                            }
                        }
                    }
                    Err(error) => return error,
                },
                _ => {
                    let value = self.eval(argument, current);
                    match truthy(&value) {
                        Ok(truth) => {
                            if truth != and {
                                return CellValue::boolean(!and);
                            }
                        }
                        Err(error) => return error,
                    }
                }
            }
        }
        CellValue::boolean(and)
    }
}

fn scalar(args: &[Argument], index: usize) -> Option<CellValue> {
    match args.get(index)? {
        Argument::Value(value) => Some(value.clone()),
        Argument::Grid(grid) => grid.values.first().cloned(),
        Argument::Invalid(value) => Some(value.clone()),
    }
}

fn numbers(args: &[Argument]) -> Result<Vec<f64>, CellValue> {
    let mut values = Vec::new();
    for argument in args {
        match argument {
            Argument::Grid(grid) => {
                for item in &grid.values {
                    match item {
                        CellValue::Number { number } => values.push(*number),
                        CellValue::Error { .. } => return Err(item.clone()),
                        _ => {}
                    }
                }
            }
            Argument::Value(value) => match value {
                CellValue::Number { number } => values.push(*number),
                CellValue::Bool { value } => values.push(if *value { 1.0 } else { 0.0 }),
                CellValue::Empty => {}
                CellValue::Error { .. } => return Err(value.clone()),
                CellValue::Text { .. } => return Err(error(VALUE)),
            },
            Argument::Invalid(value) => return Err(value.clone()),
        }
    }
    Ok(values)
}

fn count_numbers(args: &[Argument]) -> usize {
    let mut count = 0;
    for argument in args {
        match argument {
            Argument::Value(CellValue::Number { .. }) => count += 1,
            Argument::Grid(grid) => count += grid.values.iter().filter(|item| matches!(item, CellValue::Number { .. })).count(),
            _ => {}
        }
    }
    count
}

fn count_non_empty(args: &[Argument]) -> usize {
    let mut count = 0;
    for argument in args {
        match argument {
            Argument::Value(value) => {
                if !value.is_empty() {
                    count += 1;
                }
            }
            Argument::Grid(grid) => count += grid.values.iter().filter(|item| !item.is_empty()).count(),
            Argument::Invalid(_) => {}
        }
    }
    count
}

fn single_text(args: &[Argument]) -> Result<String, CellValue> {
    if args.len() != 1 {
        return Err(error(VALUE));
    }
    match scalar(args, 0) {
        Some(value) => as_text(&value),
        None => Err(error(VALUE)),
    }
}

fn single_number(args: &[Argument]) -> Result<f64, CellValue> {
    if args.len() != 1 {
        return Err(error(VALUE));
    }
    match scalar(args, 0) {
        Some(value) => numeric(&value),
        None => Err(error(VALUE)),
    }
}

/// Rounds half away from zero, like a spreadsheet does.
fn round_half_away(number: f64, digits: f64) -> f64 {
    let digits = digits.clamp(-15.0, 15.0) as i32;
    let factor = 10f64.powi(digits);
    let scaled = number * factor;
    let rounded = if scaled >= 0.0 { (scaled + 0.5).floor() } else { (scaled - 0.5).ceil() };
    normalize_zero(rounded / factor)
}

/// Collapses inner whitespace runs to one space and trims both ends.
fn trim_text(text: &str) -> String {
    let mut result = String::new();
    let mut pending_space = false;
    for character in text.chars() {
        if character == ' ' {
            pending_space = !result.is_empty();
        }
        else {
            if pending_space {
                result.push(' ');
                pending_space = false;
            }
            result.push(character);
        }
    }
    result
}

fn text_argument(args: &[Argument], index: usize) -> Result<String, CellValue> {
    match scalar(args, index) {
        Some(value) => as_text(&value),
        None => Err(error(VALUE)),
    }
}

fn number_argument(args: &[Argument], index: usize) -> Result<f64, CellValue> {
    match scalar(args, index) {
        Some(value) => numeric(&value),
        None => Err(error(VALUE)),
    }
}

fn integer_argument(args: &[Argument], index: usize) -> Result<i64, CellValue> {
    number_argument(args, index).map(|number| number.trunc() as i64)
}

fn grid_argument(args: &[Argument], index: usize) -> Option<&Grid> {
    match args.get(index)? {
        Argument::Grid(grid) => Some(grid),
        _ => None,
    }
}

/// Counts one `num_chars` argument, rejecting negative widths.
fn character_count(number: f64) -> Result<usize, CellValue> {
    if number < 0.0 {
        return Err(error(VALUE));
    }
    Ok(number.trunc() as usize)
}

fn round_toward(number: f64, digits: f64, up: bool) -> f64 {
    let digits = digits.clamp(-15.0, 15.0) as i32;
    let factor = 10f64.powi(digits);
    let scaled = number * factor;
    let magnitude = scaled.abs();
    let rounded = if up { magnitude.ceil() } else { magnitude.floor() };
    normalize_zero(if scaled < 0.0 { -rounded / factor } else { rounded / factor })
}

/// Replaces the nth occurrence of a needle, leaving every other one intact.
fn replace_nth(text: &str, needle: &str, replacement: &str, instance: usize) -> String {
    let mut output = String::new();
    let mut last = 0usize;
    let mut count = 0usize;
    for (index, _) in text.match_indices(needle) {
        count += 1;
        if count == instance {
            output.push_str(&text[last..index]);
            output.push_str(replacement);
            last = index + needle.len();
        }
    }
    output.push_str(&text[last..]);
    output
}

fn function_vlookup(args: &[Argument]) -> CellValue {
    if !(3..=4).contains(&args.len()) {
        return error(VALUE);
    }
    let Some(grid) = grid_argument(args, 1)
    else {
        return error(VALUE);
    };
    let Some(lookup) = scalar(args, 0)
    else {
        return error(VALUE);
    };
    let column = match integer_argument(args, 2) {
        Ok(column) => column,
        Err(error) => return error,
    };
    if column < 1 {
        return error(VALUE);
    }
    if column as u32 > grid.columns {
        return error(REF);
    }
    let approximate = if args.len() == 4 {
        match scalar(args, 3).map(|value| truthy(&value)) {
            Some(Ok(approximate)) => approximate,
            Some(Err(error)) => return error,
            None => true,
        }
    }
    else {
        true
    };
    let mut found: Option<u32> = None;
    for row in 1..=grid.rows {
        let key = grid.value(row, 1).cloned().unwrap_or(CellValue::Empty);
        let ordering = match compare(&key, &lookup) {
            Ok(ordering) => ordering,
            Err(error) => return error,
        };
        if ordering == Ordering::Equal {
            found = Some(row);
            break;
        }
        if approximate {
            if ordering == Ordering::Less {
                found = Some(row);
            }
            else {
                break;
            }
        }
    }
    let Some(row) = found
    else {
        return error(NA);
    };
    grid.value(row, column as u32).cloned().unwrap_or(CellValue::Empty)
}

fn function_index(args: &[Argument]) -> CellValue {
    if !(2..=3).contains(&args.len()) {
        return error(VALUE);
    }
    let Some(grid) = grid_argument(args, 0)
    else {
        return error(VALUE);
    };
    let row = match integer_argument(args, 1) {
        Ok(row) => row,
        Err(error) => return error,
    };
    let column = if args.len() == 3 {
        match integer_argument(args, 2) {
            Ok(column) => column,
            Err(error) => return error,
        }
    }
    else {
        1
    };
    if row < 1 || column < 1 {
        return error(VALUE);
    }
    let (row, column) = (row as u32, column as u32);
    if row > grid.rows || column > grid.columns {
        return error(REF);
    }
    grid.value(row, column).cloned().unwrap_or(CellValue::Empty)
}

fn function_match(args: &[Argument]) -> CellValue {
    if !(2..=3).contains(&args.len()) {
        return error(VALUE);
    }
    let Some(lookup) = scalar(args, 0)
    else {
        return error(VALUE);
    };
    let Some(grid) = grid_argument(args, 1)
    else {
        return error(VALUE);
    };
    let match_type = if args.len() == 3 {
        match integer_argument(args, 2) {
            Ok(match_type) => match_type,
            Err(error) => return error,
        }
    }
    else {
        1
    };
    if !matches!(match_type, -1 | 0 | 1) {
        return error(VALUE);
    }
    let mut found: Option<u32> = None;
    for (index, value) in grid.values.iter().enumerate() {
        let ordering = match compare(value, &lookup) {
            Ok(ordering) => ordering,
            Err(error) => return error,
        };
        if match_type == 0 {
            if ordering == Ordering::Equal {
                return CellValue::number(index as f64 + 1.0);
            }
            continue;
        }
        if match_type == 1 {
            if ordering == Ordering::Greater {
                break;
            }
            found = Some(index as u32 + 1);
        }
        else {
            if ordering == Ordering::Less {
                break;
            }
            found = Some(index as u32 + 1);
        }
    }
    match found {
        Some(position) => CellValue::number(f64::from(position)),
        None => error(NA),
    }
}

fn function_textjoin(args: &[Argument]) -> CellValue {
    if args.len() < 3 {
        return error(VALUE);
    }
    let delimiter = match text_argument(args, 0) {
        Ok(delimiter) => delimiter,
        Err(error) => return error,
    };
    let ignore_empty = match scalar(args, 1).map(|value| truthy(&value)) {
        Some(Ok(ignore_empty)) => ignore_empty,
        Some(Err(error)) => return error,
        None => return error(VALUE),
    };
    let mut parts = Vec::new();
    for argument in &args[2..] {
        match argument {
            Argument::Value(value) => {
                let text = match as_text(value) {
                    Ok(text) => text,
                    Err(error) => return error,
                };
                if !(ignore_empty && text.is_empty()) {
                    parts.push(text);
                }
            }
            Argument::Grid(grid) => {
                for value in &grid.values {
                    let text = match as_text(value) {
                        Ok(text) => text,
                        Err(error) => return error,
                    };
                    if !(ignore_empty && text.is_empty()) {
                        parts.push(text);
                    }
                }
            }
            Argument::Invalid(value) => return value.clone(),
        }
    }
    CellValue::text(parts.join(&delimiter))
}

fn function_substitute(args: &[Argument]) -> CellValue {
    if !(3..=4).contains(&args.len()) {
        return error(VALUE);
    }
    let text = match text_argument(args, 0) {
        Ok(text) => text,
        Err(error) => return error,
    };
    let old = match text_argument(args, 1) {
        Ok(old) => old,
        Err(error) => return error,
    };
    let new = match text_argument(args, 2) {
        Ok(new) => new,
        Err(error) => return error,
    };
    if old.is_empty() {
        return CellValue::text(text);
    }
    if args.len() == 4 {
        let instance = match integer_argument(args, 3) {
            Ok(instance) => instance,
            Err(error) => return error,
        };
        if instance < 1 {
            return error(VALUE);
        }
        CellValue::text(replace_nth(&text, &old, &new, instance as usize))
    }
    else {
        CellValue::text(text.replace(&old, &new))
    }
}

fn function_left(args: &[Argument]) -> CellValue {
    if !(1..=2).contains(&args.len()) {
        return error(VALUE);
    }
    let text = match text_argument(args, 0) {
        Ok(text) => text,
        Err(error) => return error,
    };
    let count = if args.len() == 2 {
        match number_argument(args, 1).and_then(character_count) {
            Ok(count) => count,
            Err(error) => return error,
        }
    }
    else {
        1
    };
    CellValue::text(text.chars().take(count).collect::<String>())
}

fn function_right(args: &[Argument]) -> CellValue {
    if !(1..=2).contains(&args.len()) {
        return error(VALUE);
    }
    let text = match text_argument(args, 0) {
        Ok(text) => text,
        Err(error) => return error,
    };
    let count = if args.len() == 2 {
        match number_argument(args, 1).and_then(character_count) {
            Ok(count) => count,
            Err(error) => return error,
        }
    }
    else {
        1
    };
    let characters: Vec<char> = text.chars().collect();
    let start = characters.len().saturating_sub(count);
    CellValue::text(characters[start..].iter().collect::<String>())
}

fn function_mid(args: &[Argument]) -> CellValue {
    if args.len() != 3 {
        return error(VALUE);
    }
    let text = match text_argument(args, 0) {
        Ok(text) => text,
        Err(error) => return error,
    };
    let start = match number_argument(args, 1) {
        Ok(start) => start,
        Err(error) => return error,
    };
    if start < 1.0 {
        return error(VALUE);
    }
    let count = match number_argument(args, 2).and_then(character_count) {
        Ok(count) => count,
        Err(error) => return error,
    };
    let start = start.trunc() as usize - 1;
    CellValue::text(text.chars().skip(start).take(count).collect::<String>())
}

fn apply_function(name: &str, args: &[Argument]) -> CellValue {
    if let Some(error) = args.iter().find_map(|argument| match argument {
        Argument::Invalid(value) => Some(value.clone()),
        _ => None,
    }) {
        return error;
    }
    match name {
        "SUM" => match numbers(args) {
            Ok(values) => CellValue::number(normalize_zero(values.iter().sum())),
            Err(error) => error,
        },
        "AVERAGE" => match numbers(args) {
            Ok(values) if values.is_empty() => error(DIV0),
            Ok(values) => CellValue::number(normalize_zero(values.iter().sum::<f64>() / values.len() as f64)),
            Err(error) => error,
        },
        "MIN" => match numbers(args) {
            Ok(values) if values.is_empty() => CellValue::number(0.0),
            Ok(values) => CellValue::number(values.iter().copied().fold(f64::INFINITY, f64::min)),
            Err(error) => error,
        },
        "MAX" => match numbers(args) {
            Ok(values) if values.is_empty() => CellValue::number(0.0),
            Ok(values) => CellValue::number(values.iter().copied().fold(f64::NEG_INFINITY, f64::max)),
            Err(error) => error,
        },
        "COUNT" => CellValue::number(count_numbers(args) as f64),
        "COUNTA" => CellValue::number(count_non_empty(args) as f64),
        "IF" => error(VALUE),
        "ROUND" => {
            if args.len() != 2 {
                return error(VALUE);
            }
            let number = match scalar(args, 0).map(|value| numeric(&value)) {
                Some(Ok(number)) => number,
                Some(Err(error)) => return error,
                None => return error(VALUE),
            };
            let digits = match scalar(args, 1).map(|value| numeric(&value)) {
                Some(Ok(digits)) => digits,
                Some(Err(error)) => return error,
                None => return error(VALUE),
            };
            CellValue::number(round_half_away(number, digits))
        }
        "ABS" => match single_number(args) {
            Ok(number) => CellValue::number(number.abs()),
            Err(error) => error,
        },
        "CONCAT" => {
            let mut output = String::new();
            for argument in args {
                match argument {
                    Argument::Value(value) => match as_text(value) {
                        Ok(text) => output.push_str(&text),
                        Err(error) => return error,
                    },
                    Argument::Grid(grid) => {
                        for item in &grid.values {
                            match as_text(item) {
                                Ok(text) => output.push_str(&text),
                                Err(error) => return error,
                            }
                        }
                    }
                    Argument::Invalid(value) => return value.clone(),
                }
            }
            CellValue::text(output)
        }
        "LEN" => match single_text(args) {
            Ok(text) => CellValue::number(text.chars().count() as f64),
            Err(error) => error,
        },
        "UPPER" => match single_text(args) {
            Ok(text) => CellValue::text(text.to_uppercase()),
            Err(error) => error,
        },
        "LOWER" => match single_text(args) {
            Ok(text) => CellValue::text(text.to_lowercase()),
            Err(error) => error,
        },
        "TRIM" => match single_text(args) {
            Ok(text) => CellValue::text(trim_text(&text)),
            Err(error) => error,
        },
        "NOT" => {
            if args.len() != 1 {
                return error(VALUE);
            }
            match scalar(args, 0) {
                Some(value) => match truthy(&value) {
                    Ok(truth) => CellValue::boolean(!truth),
                    Err(error) => error,
                },
                None => error(VALUE),
            }
        }
        "VLOOKUP" => function_vlookup(args),
        "INDEX" => function_index(args),
        "MATCH" => function_match(args),
        "LEFT" => function_left(args),
        "RIGHT" => function_right(args),
        "MID" => function_mid(args),
        "SUBSTITUTE" => function_substitute(args),
        "TEXTJOIN" => function_textjoin(args),
        "INT" => match single_number(args) {
            Ok(number) => CellValue::number(normalize_zero(number.floor())),
            Err(error) => error,
        },
        "ROUNDUP" | "ROUNDDOWN" => {
            if args.len() != 2 {
                return error(VALUE);
            }
            let number = match number_argument(args, 0) {
                Ok(number) => number,
                Err(error) => return error,
            };
            let digits = match number_argument(args, 1) {
                Ok(digits) => digits,
                Err(error) => return error,
            };
            CellValue::number(round_toward(number, digits, name == "ROUNDUP"))
        }
        "IFERROR" | "AND" | "OR" => error(VALUE),
        _ => error(NAME),
    }
}

// ---------------------------------------------------------------------------
// Public surface used by the model crate
// ---------------------------------------------------------------------------

/// Reads the value a reader should show for one cell.
pub(crate) fn value_of(cell: &Cell) -> CellValue {
    if cell.formula.is_some() {
        cell.computed.clone().or_else(|| cell.cached.clone()).unwrap_or(CellValue::Empty)
    }
    else {
        cell.value.clone()
    }
}

fn collect_forms(project: &Project) -> HashMap<CellKey, Option<Expr>> {
    let mut forms = HashMap::new();
    for (sheet_index, sheet) in project.sheets.iter().enumerate() {
        for cell in &sheet.cells {
            if let Some(source) = &cell.formula {
                forms.insert((sheet_index, cell.row, cell.column), parse_formula(source).ok());
            }
        }
    }
    forms
}

/// Reports whether a formula body parses, without exposing the expression tree.
pub(crate) fn validate(source: &str) -> Result<(), String> {
    parse_formula(source).map(|_| ())
}

/// Recomputes every formula cell and returns how many computed values changed.
pub(crate) fn recalculate(project: &mut Project) -> usize {
    let forms = collect_forms(project);
    let mut keys: Vec<CellKey> = forms.keys().copied().collect();
    keys.sort_unstable();
    let (values, unsupported) = {
        let mut evaluator = Evaluator {
            project,
            forms,
            values: HashMap::new(),
            states: HashMap::new(),
            stack: Vec::new(),
            unsupported: HashSet::new(),
        };
        for (sheet, row, column) in keys {
            evaluator.eval_cell(sheet, row, column);
        }
        (std::mem::take(&mut evaluator.values), std::mem::take(&mut evaluator.unsupported))
    };

    let mut changed = 0usize;
    for (sheet_index, sheet) in project.sheets.iter_mut().enumerate() {
        for cell in &mut sheet.cells {
            let key = (sheet_index, cell.row, cell.column);
            let (next, state) = if cell.formula.is_some() {
                let value = values.get(&key).cloned().unwrap_or_else(|| error(VALUE));
                let state = if unsupported.contains(&key) { ComputedState::Unsupported } else { ComputedState::Computed };
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

/// Formula cells and the cells they read, including every range member.
fn dependencies(project: &Project) -> HashMap<CellKey, Vec<CellKey>> {
    let mut graph = HashMap::new();
    for (sheet_index, sheet) in project.sheets.iter().enumerate() {
        for cell in &sheet.cells {
            let Some(source) = &cell.formula
            else {
                continue;
            };
            let mut targets = Vec::new();
            if let Ok(expression) = parse_formula(source) {
                collect_dependencies(&expression, project, sheet_index, &mut targets);
            }
            graph.insert((sheet_index, cell.row, cell.column), targets);
        }
    }
    graph
}

fn collect_dependencies(expression: &Expr, project: &Project, current: usize, out: &mut Vec<CellKey>) {
    match expression {
        Expr::Cell { sheet, row, column } => {
            let index = match sheet {
                None => Some(current),
                Some(name) => project.sheet_index(name),
            };
            if let Some(index) = index
                && *row >= 1
                && *row <= MAX_ROW
                && *column >= 1
                && *column <= MAX_COLUMN
            {
                out.push((index, *row, *column));
            }
        }
        Expr::Range { sheet, range } => {
            let index = match sheet {
                None => Some(current),
                Some(name) => project.sheet_index(name),
            };
            if let Some(index) = index
                && range.is_valid()
                && u64::from(range.rows()) * u64::from(range.columns()) <= MAX_RANGE_CELLS
            {
                for row in range.start_row..=range.end_row {
                    for column in range.start_column..=range.end_column {
                        out.push((index, row, column));
                    }
                }
            }
        }
        Expr::Negate(inner) | Expr::Percent(inner) => collect_dependencies(inner, project, current, out),
        Expr::Binary { left, right, .. } => {
            collect_dependencies(left, project, current, out);
            collect_dependencies(right, project, current, out);
        }
        Expr::Call { args, .. } => {
            for argument in args {
                collect_dependencies(argument, project, current, out);
            }
        }
        _ => {}
    }
}

struct CycleScan<'a> {
    graph: &'a HashMap<CellKey, Vec<CellKey>>,
    colors: HashMap<CellKey, u8>,
    stack: Vec<CellKey>,
    cycles: Vec<Vec<CellKey>>,
    reported: HashSet<Vec<CellKey>>,
}

impl CycleScan<'_> {
    fn visit(&mut self, node: CellKey) {
        self.colors.insert(node, 1);
        self.stack.push(node);
        if let Some(targets) = self.graph.get(&node).cloned() {
            for target in targets {
                match self.colors.get(&target).copied().unwrap_or(0) {
                    0 => {
                        if self.graph.contains_key(&target) {
                            self.visit(target);
                        }
                    }
                    1 => {
                        if let Some(position) = self.stack.iter().position(|entry| *entry == target) {
                            let cycle = self.stack[position..].to_vec();
                            let mut key = cycle.clone();
                            key.sort_unstable();
                            if self.reported.insert(key) {
                                self.cycles.push(cycle);
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
        self.stack.pop();
        self.colors.insert(node, 2);
    }
}

/// Finds every formula cycle, including cycles that cross sheets.
pub(crate) fn find_cycles(project: &Project) -> Vec<Vec<CellKey>> {
    let graph = dependencies(project);
    let mut scan =
        CycleScan { graph: &graph, colors: HashMap::new(), stack: Vec::new(), cycles: Vec::new(), reported: HashSet::new() };
    let mut keys: Vec<CellKey> = graph.keys().copied().collect();
    keys.sort_unstable();
    for key in keys {
        if scan.colors.get(&key).copied().unwrap_or(0) == 0 {
            scan.visit(key);
        }
    }
    scan.cycles
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{CellValue, ComputedState, Project};

    fn project_with(cells: &[(&str, CellValue)]) -> Project {
        let mut project = Project::blank("Eval");
        for (reference, value) in cells {
            let (row, column) = parse_reference(reference).expect("reference");
            project.sheets[0].write_cell(row, column, value.clone(), None, None);
        }
        project
    }

    fn evaluate(project: &mut Project, source: &str) -> CellValue {
        project.sheets[0].write_cell(20, 20, CellValue::Empty, Some(source.to_owned()), None);
        recalculate(project);
        project.sheets[0].cell(20, 20).expect("cell").computed.clone().expect("computed")
    }

    #[test]
    fn arithmetic_follows_precedence() {
        let mut project = project_with(&[("A1", CellValue::number(2.0)), ("A2", CellValue::number(3.0))]);
        assert_eq!(evaluate(&mut project, "=A1+A2*4"), CellValue::number(14.0));
        assert_eq!(evaluate(&mut project, "=(A1+A2)*4"), CellValue::number(20.0));
        assert_eq!(evaluate(&mut project, "=2^3^2"), CellValue::number(64.0));
        assert_eq!(evaluate(&mut project, "=-2^2"), CellValue::number(4.0));
        assert_eq!(evaluate(&mut project, "=50%"), CellValue::number(0.5));
        assert_eq!(evaluate(&mut project, "=A2-10"), CellValue::number(-7.0));
    }

    #[test]
    fn absent_cells_are_zero_in_arithmetic() {
        let mut project = project_with(&[]);
        assert_eq!(evaluate(&mut project, "=Z99+1"), CellValue::number(1.0));
        assert_eq!(evaluate(&mut project, "=Z99*3"), CellValue::number(0.0));
    }

    #[test]
    fn text_in_arithmetic_is_a_value_error() {
        let mut project = project_with(&[("A1", CellValue::text("abc"))]);
        assert_eq!(evaluate(&mut project, "=A1+1"), error(VALUE));
        assert_eq!(evaluate(&mut project, "=1/0"), error(DIV0));
        assert_eq!(evaluate(&mut project, "=Missing!A1"), error(REF));
        assert_eq!(evaluate(&mut project, "=NOPE(1)"), error(NAME));
    }

    #[test]
    fn comparisons_and_concatenation() {
        let mut project = project_with(&[("A1", CellValue::number(2.0)), ("B1", CellValue::text("b"))]);
        assert_eq!(evaluate(&mut project, "=A1>1"), CellValue::boolean(true));
        assert_eq!(evaluate(&mut project, "=A1<=1"), CellValue::boolean(false));
        assert_eq!(evaluate(&mut project, "=A1<>2"), CellValue::boolean(false));
        assert_eq!(evaluate(&mut project, "=\"a\"&A1&B1"), CellValue::text("a2b"));
        assert_eq!(evaluate(&mut project, "=B1=\"B\""), CellValue::boolean(true));
        assert_eq!(evaluate(&mut project, "=\"x\"&TRUE"), CellValue::text("xTRUE"));
    }

    #[test]
    fn aggregate_functions_skip_text_inside_ranges() {
        let mut project = project_with(&[
            ("A1", CellValue::number(2.0)),
            ("A2", CellValue::number(4.0)),
            ("A3", CellValue::text("note")),
            ("A4", CellValue::Empty),
        ]);
        assert_eq!(evaluate(&mut project, "=SUM(A1:A4)"), CellValue::number(6.0));
        assert_eq!(evaluate(&mut project, "=AVERAGE(A1:A2)"), CellValue::number(3.0));
        assert_eq!(evaluate(&mut project, "=MIN(A1:A2)"), CellValue::number(2.0));
        assert_eq!(evaluate(&mut project, "=MAX(A1:A2)"), CellValue::number(4.0));
        assert_eq!(evaluate(&mut project, "=COUNT(A1:A4)"), CellValue::number(2.0));
        assert_eq!(evaluate(&mut project, "=COUNTA(A1:A4)"), CellValue::number(3.0));
    }

    #[test]
    fn conditional_and_text_functions() {
        let mut project = project_with(&[("A1", CellValue::number(7.0))]);
        assert_eq!(evaluate(&mut project, "=IF(A1>5,\"big\",\"small\")"), CellValue::text("big"));
        assert_eq!(evaluate(&mut project, "=IF(A1>9,1)"), CellValue::boolean(false));
        assert_eq!(evaluate(&mut project, "=ROUND(3.14159,2)"), CellValue::number(3.15 - 0.01));
        assert_eq!(evaluate(&mut project, "=ROUND(2.5,0)"), CellValue::number(3.0));
        assert_eq!(evaluate(&mut project, "=ROUND(-2.5,0)"), CellValue::number(-3.0));
        assert_eq!(evaluate(&mut project, "=ABS(-3)"), CellValue::number(3.0));
        assert_eq!(evaluate(&mut project, "=CONCAT(\"a\",1,TRUE)"), CellValue::text("a1TRUE"));
        assert_eq!(evaluate(&mut project, "=LEN(\"abcd\")"), CellValue::number(4.0));
        assert_eq!(evaluate(&mut project, "=UPPER(\"ab\")"), CellValue::text("AB"));
        assert_eq!(evaluate(&mut project, "=LOWER(\"AB\")"), CellValue::text("ab"));
        assert_eq!(evaluate(&mut project, "=TRIM(\"  a   b \")"), CellValue::text("a b"));
        assert_eq!(evaluate(&mut project, "=NOT(FALSE)"), CellValue::boolean(true));
    }

    #[test]
    fn ranges_cross_sheets() {
        let mut project = project_with(&[("A1", CellValue::number(1.0))]);
        project.sheets.push(crate::model::Sheet::new("sheet-2", "Data", 1));
        let data = &mut project.sheets[1];
        data.write_cell(1, 1, CellValue::number(10.0), None, None);
        data.write_cell(2, 1, CellValue::number(20.0), None, None);
        data.write_cell(2, 2, CellValue::number(2.0), None, None);
        assert_eq!(evaluate(&mut project, "=SUM(Data!A1:A2)"), CellValue::number(30.0));
        assert_eq!(evaluate(&mut project, "=SUM(Data!A1:B2)"), CellValue::number(32.0));
        assert_eq!(evaluate(&mut project, "=Data!A2*2"), CellValue::number(40.0));
        project.sheets[1].name = "My Data".to_owned();
        assert_eq!(evaluate(&mut project, "=SUM('My Data'!A1:A2)"), CellValue::number(30.0));
    }

    #[test]
    fn cycles_resolve_to_the_cycle_error() {
        let mut project = Project::blank("Cycle");
        project.sheets[0].write_cell(1, 1, CellValue::Empty, Some("=B1".to_owned()), None);
        project.sheets[0].write_cell(1, 2, CellValue::Empty, Some("=A1".to_owned()), None);
        recalculate(&mut project);
        assert_eq!(project.sheets[0].cell(1, 1).unwrap().computed, Some(error(CYCLE)));
        assert_eq!(project.sheets[0].cell(1, 2).unwrap().computed, Some(error(CYCLE)));
        let cycles = find_cycles(&project);
        assert_eq!(cycles.len(), 1);
        assert_eq!(cycles[0].len(), 2);
    }

    #[test]
    fn chains_evaluate_once_and_report_changes() {
        let mut project = Project::blank("Chain");
        project.sheets[0].write_cell(1, 1, CellValue::number(2.0), None, None);
        project.sheets[0].write_cell(1, 2, CellValue::Empty, Some("=A1*3".to_owned()), None);
        project.sheets[0].write_cell(1, 3, CellValue::Empty, Some("=B1+1".to_owned()), None);
        assert_eq!(recalculate(&mut project), 2);
        assert_eq!(project.sheets[0].cell(1, 3).unwrap().computed, Some(CellValue::number(7.0)));
        assert_eq!(recalculate(&mut project), 0);
    }

    fn lookup_project() -> Project {
        let mut project = Project::blank("Lookup");
        for (row, key, label, amount, descending) in
            [(1, 1.0, "one", 10.0, 30.0), (2, 2.0, "two", 20.0, 20.0), (3, 3.0, "three", 30.0, 10.0)]
        {
            project.sheets[0].write_cell(row, 1, CellValue::number(key), None, None);
            project.sheets[0].write_cell(row, 2, CellValue::text(label), None, None);
            project.sheets[0].write_cell(row, 3, CellValue::number(amount), None, None);
            project.sheets[0].write_cell(row, 4, CellValue::number(descending), None, None);
        }
        project
    }

    #[test]
    fn lookup_functions_cover_exact_and_approximate_matches() {
        let mut project = lookup_project();
        assert_eq!(evaluate(&mut project, "=VLOOKUP(2,A1:C3,2,FALSE)"), CellValue::text("two"));
        assert_eq!(evaluate(&mut project, "=VLOOKUP(2.5,A1:C3,3,TRUE)"), CellValue::number(20.0));
        assert_eq!(evaluate(&mut project, "=VLOOKUP(2.5,A1:C3,3)"), CellValue::number(20.0));
        assert_eq!(evaluate(&mut project, "=VLOOKUP(9,A1:C3,2,FALSE)"), error(NA));
        assert_eq!(evaluate(&mut project, "=VLOOKUP(1,A1:C3,4)"), error(REF));
        assert_eq!(evaluate(&mut project, "=INDEX(A1:C3,2,3)"), CellValue::number(20.0));
        assert_eq!(evaluate(&mut project, "=INDEX(A1:C3,4,1)"), error(REF));
        assert_eq!(evaluate(&mut project, "=MATCH(2,A1:A3,0)"), CellValue::number(2.0));
        assert_eq!(evaluate(&mut project, "=MATCH(2.5,A1:A3,1)"), CellValue::number(2.0));
        assert_eq!(evaluate(&mut project, "=MATCH(15,D1:D3,-1)"), CellValue::number(2.0));
        assert_eq!(evaluate(&mut project, "=MATCH(9,A1:A3,0)"), error(NA));
    }

    #[test]
    fn conditional_logical_and_text_functions_are_lazy() {
        let mut project = lookup_project();
        assert_eq!(evaluate(&mut project, "=IFERROR(1/0,\"safe\")"), CellValue::text("safe"));
        assert_eq!(evaluate(&mut project, "=IFERROR(5,\"safe\")"), CellValue::number(5.0));
        assert_eq!(evaluate(&mut project, "=IFERROR(1,1/0)"), CellValue::number(1.0));
        assert_eq!(evaluate(&mut project, "=AND(TRUE,1)"), CellValue::boolean(true));
        assert_eq!(evaluate(&mut project, "=AND(TRUE,0,1/0)"), CellValue::boolean(false));
        assert_eq!(evaluate(&mut project, "=OR(FALSE,0,2)"), CellValue::boolean(true));
        assert_eq!(evaluate(&mut project, "=OR(FALSE,0,1/0)"), error(DIV0));
        assert_eq!(evaluate(&mut project, "=AND(A1:A3)"), CellValue::boolean(true));
        assert_eq!(evaluate(&mut project, "=LEFT(\"hello\",2)"), CellValue::text("he"));
        assert_eq!(evaluate(&mut project, "=RIGHT(\"hello\",3)"), CellValue::text("llo"));
        assert_eq!(evaluate(&mut project, "=RIGHT(\"hello\")"), CellValue::text("o"));
        assert_eq!(evaluate(&mut project, "=MID(\"hello\",2,3)"), CellValue::text("ell"));
        assert_eq!(evaluate(&mut project, "=SUBSTITUTE(\"a-b-a\",\"a\",\"x\")"), CellValue::text("x-b-x"));
        assert_eq!(evaluate(&mut project, "=SUBSTITUTE(\"a-b-a\",\"a\",\"x\",2)"), CellValue::text("a-b-x"));
        assert_eq!(evaluate(&mut project, "=TEXTJOIN(\"-\",TRUE,A1:A2)"), CellValue::text("1-2"));
        assert_eq!(evaluate(&mut project, "=INT(-1.5)"), CellValue::number(-2.0));
        assert_eq!(evaluate(&mut project, "=ROUNDUP(2.1,0)"), CellValue::number(3.0));
        assert_eq!(evaluate(&mut project, "=ROUNDDOWN(2.9,0)"), CellValue::number(2.0));
        assert_eq!(evaluate(&mut project, "=ROUNDUP(-2.1,0)"), CellValue::number(-3.0));
        assert_eq!(evaluate(&mut project, "=ROUNDDOWN(-2.9,0)"), CellValue::number(-2.0));
        assert_eq!(evaluate(&mut project, "=ROUNDUP(2.9,-1)"), CellValue::number(10.0));
    }

    #[test]
    fn unsupported_functions_keep_the_imported_cache() {
        let mut project = Project::blank("Unsupported");
        project.sheets[0].write_cell(1, 1, CellValue::number(2.0), None, None);
        project.sheets[0].write_cell(1, 2, CellValue::Empty, Some("=FUTUREFN(A1)".to_owned()), Some(CellValue::number(7.0)));
        project.sheets[0].write_cell(1, 3, CellValue::Empty, Some("=B1*2".to_owned()), None);
        recalculate(&mut project);
        let cell = project.sheets[0].cell(1, 2).expect("B1");
        assert_eq!(cell.computed, Some(CellValue::number(7.0)));
        assert_eq!(cell.computed_state, Some(ComputedState::Unsupported));
        assert_eq!(project.sheets[0].cell(1, 3).expect("C1").computed, Some(CellValue::number(14.0)));
    }

    #[test]
    fn functions_without_a_cache_still_report_name() {
        let mut project = Project::blank("Missing");
        project.sheets[0].write_cell(1, 1, CellValue::Empty, Some("=FUTUREFN(1)".to_owned()), None);
        recalculate(&mut project);
        let cell = project.sheets[0].cell(1, 1).expect("A1");
        assert_eq!(cell.computed, Some(error(NAME)));
        assert_eq!(cell.computed_state, Some(ComputedState::Unsupported));
    }
}
