//! Process-wide handle store.
//!
//! Handles, revisions and the undo stack live in the shared `document-session`
//! store. This module keeps only the workbook state a handle points at.

use std::{path::PathBuf, sync::LazyLock};

use document_session::Store;

use crate::{error::ServiceError, model::Project};

/// One open workbook.
pub struct Entry {
    /// Editable project.
    pub project: Project,
    /// Monotonic revision, bumped by every committed edit.
    pub revision: u64,
    /// Last path the project was opened from or saved to.
    pub path: Option<PathBuf>,
}

impl document_session::Revisioned for Entry {
    fn revision(&self) -> u64 {
        self.revision
    }

    fn set_revision(&mut self, revision: u64) {
        self.revision = revision;
    }
}

static STORE: LazyLock<Store<Entry>> = LazyLock::new(Store::new);

/// Inserts an entry and returns its new handle.
pub fn insert(project: Project, path: Option<PathBuf>) -> u64 {
    STORE.insert(Entry { project, revision: 0, path })
}

/// Runs `operation` against one entry.
pub fn with<R>(handle: u64, operation: impl FnOnce(&Entry) -> R) -> Result<R, ServiceError> {
    STORE.with(handle, operation).ok_or_else(not_found)
}

/// Runs `operation` against one entry mutably.
pub fn with_mut<R>(handle: u64, operation: impl FnOnce(&mut Entry) -> R) -> Result<R, ServiceError> {
    STORE.with_mut(handle, |record| operation(record.state_mut())).ok_or_else(not_found)
}

/// Removes one entry; closing an unknown handle is a no-op.
pub fn remove(handle: u64) -> bool {
    STORE.remove(handle)
}

fn not_found() -> ServiceError {
    ServiceError::new("workbook/not-found", "no workbook is open for this handle")
}
