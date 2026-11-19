//! Declared workbook model contracts for OpenExcel.
//!
//! The spreadsheet engine — workbook parsing, formula evaluation, pivot
//! aggregation, rendering and export — does not exist in this repository yet.
//! This crate therefore publishes one honest capability matrix and rejects every
//! domain operation with the stable `capability/not-implemented` code instead of
//! fabricating a workbook, a summary or an inspection report.

#![deny(missing_docs)]

use std::{fmt, path::Path};

use serde_json::{Map, Value, json};

/// Workbook formats the product must classify, even while none is implemented.
pub const WORKBOOK_FORMATS: &[&str] = &["xlsx", "xlsm", "xltx", "xls", "csv", "tsv", "ods", "json", "xml", "pdf"];

/// Rule checks the inspector will expose once the engine lands.
pub const INSPECT_CHECKS: &[&str] = &["formulas", "references", "cycles", "types", "ranges", "styles"];

/// Render targets the studio will expose once the engine lands.
pub const RENDER_FORMATS: &[&str] = &["svg", "png", "pdf"];

/// Export targets the SDK will expose once the engine lands.
pub const EXPORT_TARGETS: &[&str] = &["xlsx", "xlsm", "csv", "json", "pdf"];

/// Stable failure of one binding operation.
#[derive(Debug)]
pub struct ServiceError {
    /// Stable, comparable code.
    pub code: String,
    /// Human-readable message; never carries cell data.
    pub message: String,
}

impl ServiceError {
    /// Builds a failure from a stable code and a message.
    pub fn new(code: impl Into<String>, message: impl Into<String>) -> Self {
        Self { code: code.into(), message: message.into() }
    }
}

impl fmt::Display for ServiceError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}: {}", self.code, self.message)
    }
}

impl std::error::Error for ServiceError {}

/// Every domain operation reports this until the engine exists.
fn not_implemented(operation: &str) -> ServiceError {
    ServiceError::new(
        "capability/not-implemented",
        format!("`{operation}` needs the spreadsheet engine, which is not implemented yet"),
    )
}

/// A `{ name: "unsupported" }` object for one capability family.
fn unsupported_map(names: &[&str]) -> Value {
    Value::Object(names.iter().map(|name| ((*name).to_owned(), Value::String("unsupported".to_owned()))).collect())
}

/// What this build can actually do.
///
/// The matrix only reports what is wired. No workbook engine ships yet, so the
/// project container, every classified format, every inspection check, every
/// renderer and every export target is `unsupported`. Callers read this instead
/// of treating a method name or a file extension as a capability.
pub fn capabilities() -> Value {
    let mut formats = Map::new();
    for format in WORKBOOK_FORMATS {
        formats.insert(
            (*format).to_owned(),
            json!({
                "import": "unsupported",
                "edit": "unsupported",
                "export": "unsupported",
                "inspect": "unsupported",
                "lossCodes": [],
            }),
        );
    }

    json!({
        "engine": { "package": env!("CARGO_PKG_NAME"), "version": env!("CARGO_PKG_VERSION") },
        "project": "unsupported",
        "formats": Value::Object(formats),
        "checks": unsupported_map(INSPECT_CHECKS),
        "renderers": unsupported_map(RENDER_FORMATS),
        "exportTargets": unsupported_map(EXPORT_TARGETS),
    })
}

/// Creates a workbook from a JSON request.
///
/// Always answers `capability/not-implemented`; no workbook handle can exist
/// before the engine lands.
pub fn create(_request: &Value) -> Result<Value, ServiceError> {
    Err(not_implemented("workbook.create"))
}

/// Opens a workbook project from a path.
///
/// Always answers `capability/not-implemented`.
pub fn open(_path: &Path) -> Result<Value, ServiceError> {
    Err(not_implemented("workbook.open"))
}

/// Writes one open workbook through a temporary file and a rename.
///
/// Always answers `capability/not-implemented`.
pub fn save(_handle: u64, _path: &Path, _overwrite: bool) -> Result<Value, ServiceError> {
    Err(not_implemented("workbook.save"))
}

/// Runs the requested rule checks and reports findings with their locations.
///
/// Always answers `capability/not-implemented`.
pub fn inspect(_handle: u64, _checks: Option<Vec<String>>) -> Result<Value, ServiceError> {
    Err(not_implemented("workbook.inspect"))
}

/// Returns the summary of one open workbook.
///
/// Always answers `capability/not-implemented`.
pub fn summary(_handle: u64) -> Result<Value, ServiceError> {
    Err(not_implemented("workbook.summary"))
}

/// Writes one open workbook to an export target.
///
/// Always answers `capability/not-implemented`.
pub fn export(_handle: u64, _format: &str, _target: &Path, _overwrite: bool) -> Result<Value, ServiceError> {
    Err(not_implemented("workbook.export"))
}

/// Releases one open workbook handle.
///
/// Always answers `capability/not-implemented`.
pub fn close(_handle: u64) -> Result<(), ServiceError> {
    Err(not_implemented("workbook.close"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn capability_matrix_discloses_no_engine() {
        let matrix = capabilities();

        assert_eq!(matrix["project"], json!("unsupported"));
        assert_eq!(matrix["formats"].as_object().map(Map::len), Some(WORKBOOK_FORMATS.len()));

        for format in WORKBOOK_FORMATS {
            assert_eq!(matrix["formats"][*format]["import"], json!("unsupported"));
            assert_eq!(matrix["formats"][*format]["edit"], json!("unsupported"));
            assert_eq!(matrix["formats"][*format]["export"], json!("unsupported"));
            assert_eq!(matrix["formats"][*format]["inspect"], json!("unsupported"));
        }
        for check in INSPECT_CHECKS {
            assert_eq!(matrix["checks"][*check], json!("unsupported"));
        }
        for renderer in RENDER_FORMATS {
            assert_eq!(matrix["renderers"][*renderer], json!("unsupported"));
        }
        for target in EXPORT_TARGETS {
            assert_eq!(matrix["exportTargets"][*target], json!("unsupported"));
        }
    }

    #[test]
    fn domain_operations_report_one_stable_code() {
        let errors = [
            create(&json!({ "title": "Ledger" })).unwrap_err(),
            open(Path::new("ledger.oexcel")).unwrap_err(),
            save(1, Path::new("ledger.oexcel"), false).unwrap_err(),
            inspect(1, None).unwrap_err(),
            summary(1).unwrap_err(),
            export(1, "csv", Path::new("ledger.csv"), false).unwrap_err(),
            close(1).unwrap_err(),
        ];

        for error in &errors {
            assert_eq!(error.code, "capability/not-implemented");
            assert!(!error.message.is_empty());
        }
    }
}
