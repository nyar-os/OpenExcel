//! One stable failure of a workbook operation.

use std::{fmt, io, path::Path};

use document_spreadsheet_formula::error::params::{ErrorParams, ObjectKind};
use serde_json::Value;

/// Stable failure of one binding operation.
///
/// Construction goes through the typed constructors below: each one fixes a
/// stable code and attaches the frozen parameters that code's contract
/// promises, so the binding forwards `code`, `params` and the original OS error
/// instead of re-deriving them from prose. Callers compare `code` and read
/// `params`; they never parse the message.
#[derive(Debug)]
pub struct ServiceError {
    message: String,
    params: ErrorParams,
    source: Option<Box<dyn std::error::Error + Send + Sync>>,
}

impl ServiceError {
    fn build(params: ErrorParams, message: impl Into<String>) -> Self {
        Self { message: message.into(), params, source: None }
    }

    fn source(mut self, error: impl std::error::Error + Send + Sync + 'static) -> Self {
        self.source = Some(Box::new(error));
        self
    }

    /// `session/argument-invalid`: a request parameter was rejected.
    pub fn argument(reason: impl Into<String>) -> Self {
        let reason = reason.into();
        Self::build(ErrorParams::ArgumentInvalid { reason: reason.clone() }, reason)
    }

    /// `object/not-found`: the addressed object is absent from the workbook.
    pub fn not_found(kind: ObjectKind, id: &str) -> Self {
        Self::build(
            ErrorParams::NotFound { kind, id: Some(id.to_owned()) },
            format!("`{}` `{id}` does not exist", kind.as_str()),
        )
    }

    /// `object/not-found`: the workbook has no sheet to address at all.
    pub fn no_sheet() -> Self {
        Self::build(ErrorParams::NotFound { kind: ObjectKind::Sheet, id: None }, "the workbook has no sheet")
    }

    /// `workbook/not-found`: no workbook is open for this handle.
    pub fn workbook_not_found() -> Self {
        Self::build(ErrorParams::WorkbookNotFound, "no workbook is open for this handle")
    }

    /// `document/revision-conflict`: the workbook moved past the expected revision.
    pub fn revision_conflict(expected: u64, actual: u64) -> Self {
        Self::build(
            ErrorParams::RevisionConflict { actual, expected: Some(expected) },
            format!("expected revision {expected} but the workbook is at revision {actual}"),
        )
    }

    /// `object/unsupported`: the requested template is not available in this build.
    pub fn template_unavailable(template: &str) -> Self {
        Self::build(
            ErrorParams::Unsupported { operation: "template".to_owned(), name: template.to_owned() },
            format!("template `{template}` is not available; this build starts from `blank`"),
        )
    }

    /// `object/unsupported`: the edit operation name is not implemented.
    pub fn operation_unsupported(operation: &str) -> Self {
        Self::build(
            ErrorParams::Unsupported { operation: "edit operation".to_owned(), name: operation.to_owned() },
            format!("edit operation `{operation}` is not supported"),
        )
    }

    /// `format/invalid`: an input failed validation for the claimed format.
    pub fn invalid(format: &str, reason: impl Into<String>) -> Self {
        let reason = reason.into();
        Self::build(
            ErrorParams::FormatInvalid { format: format.to_owned(), reason: reason.clone() },
            format!("`{format}` input is invalid: {reason}"),
        )
    }

    /// `format/unsupported`: the requested format is not offered by this build.
    pub fn unsupported(format: &str) -> Self {
        Self::build(
            ErrorParams::FormatUnsupported { format: format.to_owned() },
            format!("`{format}` is not supported by this build"),
        )
    }

    /// `io/read`: reading a file failed; the OS error kind is preserved.
    pub fn io_read(path: &Path, error: io::Error) -> Self {
        let kind = format!("{:?}", error.kind());
        Self::build(
            ErrorParams::IoRead { path: path.display().to_string(), kind },
            format!("could not read `{}`: {error}", path.display()),
        )
        .source(error)
    }

    /// `io/write`: publishing a file failed.
    pub fn io_write(path: &Path, reason: impl Into<String>) -> Self {
        Self::build(
            ErrorParams::IoWrite { path: path.display().to_string() },
            format!("could not publish `{}`: {}", path.display(), reason.into()),
        )
    }

    /// `io/exists`: the target is present and overwrite was not granted.
    pub fn io_exists(path: &Path) -> Self {
        Self::build(
            ErrorParams::IoExists { path: path.display().to_string() },
            format!("`{}` already exists; pass overwrite to replace it", path.display()),
        )
    }

    /// `io/failure`: an engine-level package failure with no OS error behind it.
    pub fn io(reason: impl Into<String>) -> Self {
        let reason = reason.into();
        Self::build(ErrorParams::IoFailure { reason: reason.clone() }, reason)
    }

    /// Stable, comparable code.
    pub fn code(&self) -> &'static str {
        self.params.code()
    }

    /// Human-readable message; never carries cell data.
    pub fn message(&self) -> &str {
        &self.message
    }

    /// Typed per-code parameters, frozen by [`ErrorParams`].
    pub fn params(&self) -> &ErrorParams {
        &self.params
    }

    /// Serialises the frozen parameters for a binding envelope.
    pub fn params_json(&self) -> Value {
        self.params.to_json()
    }
}

impl fmt::Display for ServiceError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}: {}", self.code(), self.message)
    }
}

impl std::error::Error for ServiceError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        self.source.as_ref().map(|error| error.as_ref() as &(dyn std::error::Error + 'static))
    }
}
