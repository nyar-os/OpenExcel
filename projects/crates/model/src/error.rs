//! One stable failure of a workbook operation.

use std::fmt;

/// Stable failure of one binding operation.
///
/// The code is contract: callers compare it and never parse the message.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ServiceError {
    /// Stable, comparable code such as `format/unsupported`.
    pub code: String,
    /// Human-readable message; never carries cell data.
    pub message: String,
}

impl ServiceError {
    /// Builds a failure from a stable code and a message.
    pub fn new(code: impl Into<String>, message: impl Into<String>) -> Self {
        Self { code: code.into(), message: message.into() }
    }

    /// Reports an input that does not satisfy the operation contract.
    pub fn argument(message: impl Into<String>) -> Self {
        Self::new("session/argument-invalid", message)
    }

    /// Reports an input file that cannot be parsed as the claimed format.
    pub fn invalid(format: &str, message: impl Into<String>) -> Self {
        Self::new("format/invalid", format!("`{format}` input is invalid: {}", message.into()))
    }

    /// Reports a format the installed build does not read or write.
    pub fn unsupported(format: &str) -> Self {
        Self::new("format/unsupported", format!("`{format}` is not supported by this build"))
    }

    /// Reports an operating-system failure while reading or writing a file.
    pub fn io(message: impl Into<String>) -> Self {
        Self::new("io/failure", message)
    }
}

impl fmt::Display for ServiceError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}: {}", self.code, self.message)
    }
}

impl std::error::Error for ServiceError {}
