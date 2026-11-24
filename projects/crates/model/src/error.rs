//! One stable failure of a workbook operation.

use std::{fmt, io, path::Path};

use serde_json::{Value, json};

/// Stable failure of one binding operation.
///
/// Construction goes through the typed constructors below: each one fixes a
/// stable code and attaches the parameters that code's contract promises, so
/// the binding forwards `code`, `params` and the original OS error instead of
/// re-deriving them from prose. Callers compare `code` and read `params`;
/// they never parse the message.
#[derive(Debug)]
pub struct ServiceError {
    code: String,
    message: String,
    params: Value,
    source: Option<Box<dyn std::error::Error + Send + Sync>>,
}

impl ServiceError {
    fn build(code: &str, message: impl Into<String>) -> Self {
        Self { code: code.to_owned(), message: message.into(), params: Value::Null, source: None }
    }

    fn with_params(mut self, params: Value) -> Self {
        self.params = params;
        self
    }

    fn source(mut self, error: impl std::error::Error + Send + Sync + 'static) -> Self {
        self.source = Some(Box::new(error));
        self
    }

    /// `session/argument-invalid`: a request parameter was rejected.
    pub fn argument(message: impl Into<String>) -> Self {
        Self::build("session/argument-invalid", message)
    }

    /// `object/not-found`: the addressed object is absent from the workbook.
    pub fn not_found(kind: &str, id: &str) -> Self {
        Self::build("object/not-found", format!("`{kind}` `{id}` does not exist")).with_params(json!({ "kind": kind, "id": id }))
    }

    /// `object/not-found`: the workbook has no sheet to address at all.
    pub fn no_sheet() -> Self {
        Self::build("object/not-found", "the workbook has no sheet").with_params(json!({ "kind": "sheet" }))
    }

    /// `workbook/not-found`: no workbook is open for this handle.
    pub fn workbook_not_found() -> Self {
        Self::build("workbook/not-found", "no workbook is open for this handle")
    }

    /// `document/revision-conflict`: the workbook moved past the expected revision.
    pub fn revision_conflict(expected: u64, actual: u64) -> Self {
        Self::build(
            "document/revision-conflict",
            format!("expected revision {expected} but the workbook is at revision {actual}"),
        )
        .with_params(json!({ "expected": expected, "actual": actual }))
    }

    /// `object/unsupported`: the requested template is not available in this build.
    pub fn template_unavailable(template: &str) -> Self {
        Self::build(
            "object/unsupported",
            format!("template `{template}` is not available; this build starts from `blank`"),
        )
        .with_params(json!({ "operation": "template", "name": template }))
    }

    /// `object/unsupported`: the edit operation name is not implemented.
    pub fn operation_unsupported(operation: &str) -> Self {
        Self::build("object/unsupported", format!("edit operation `{operation}` is not supported"))
            .with_params(json!({ "operation": "edit operation", "name": operation }))
    }

    /// `format/invalid`: an input failed validation for the claimed format.
    pub fn invalid(format: &str, reason: impl Into<String>) -> Self {
        let reason = reason.into();
        Self::build("format/invalid", format!("`{format}` input is invalid: {reason}"))
            .with_params(json!({ "format": format, "reason": reason }))
    }

    /// `format/unsupported`: the requested format is not offered by this build.
    pub fn unsupported(format: &str) -> Self {
        Self::build("format/unsupported", format!("`{format}` is not supported by this build"))
            .with_params(json!({ "format": format }))
    }

    /// `io/read`: reading a file failed; the OS error kind is preserved.
    pub fn io_read(path: &Path, error: io::Error) -> Self {
        let kind = format!("{:?}", error.kind());
        Self::build("io/read", format!("could not read `{}`: {error}", path.display()))
            .with_params(json!({ "path": path.display().to_string(), "kind": kind }))
            .source(error)
    }

    /// `io/write`: publishing a file failed.
    pub fn io_write(path: &Path, reason: impl Into<String>) -> Self {
        Self::build("io/write", format!("could not publish `{}`: {}", path.display(), reason.into()))
            .with_params(json!({ "path": path.display().to_string() }))
    }

    /// `io/exists`: the target is present and overwrite was not granted.
    pub fn io_exists(path: &Path) -> Self {
        Self::build("io/exists", format!("`{}` already exists; pass overwrite to replace it", path.display()))
            .with_params(json!({ "path": path.display().to_string() }))
    }

    /// `io/failure`: an engine-level package failure with no OS error behind it.
    pub fn io(message: impl Into<String>) -> Self {
        Self::build("io/failure", message)
    }

    /// Stable, comparable code.
    pub fn code(&self) -> &str {
        &self.code
    }

    /// Human-readable message; never carries cell data.
    pub fn message(&self) -> &str {
        &self.message
    }

    /// Typed per-code parameters, or `Value::Null` when the code carries none.
    pub fn params(&self) -> &Value {
        &self.params
    }
}

impl fmt::Display for ServiceError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}: {}", self.code, self.message)
    }
}

impl std::error::Error for ServiceError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        self.source.as_ref().map(|error| error.as_ref() as &(dyn std::error::Error + 'static))
    }
}
