//! Process-wide handle store.
//!
//! Handles are opaque monotonic integers. Every operation locks the store,
//! borrows one entry and releases the lock before returning, so a caller can
//! never observe a half-applied edit.

use std::{
    collections::HashMap,
    path::PathBuf,
    sync::{
        LazyLock, Mutex,
        atomic::{AtomicU64, Ordering},
    },
};

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

static STORE: LazyLock<Mutex<HashMap<u64, Entry>>> = LazyLock::new(|| Mutex::new(HashMap::new()));
static NEXT: AtomicU64 = AtomicU64::new(1);

fn store() -> std::sync::MutexGuard<'static, HashMap<u64, Entry>> {
    STORE.lock().unwrap_or_else(|poisoned| poisoned.into_inner())
}

/// Inserts an entry and returns its new handle.
pub fn insert(project: Project, path: Option<PathBuf>) -> u64 {
    let handle = NEXT.fetch_add(1, Ordering::Relaxed);
    store().insert(handle, Entry { project, revision: 0, path });
    handle
}

/// Runs `operation` against one entry.
pub fn with<R>(handle: u64, operation: impl FnOnce(&Entry) -> R) -> Result<R, ServiceError> {
    let guard = store();
    let entry = guard.get(&handle).ok_or_else(not_found)?;
    Ok(operation(entry))
}

/// Runs `operation` against one entry mutably.
pub fn with_mut<R>(handle: u64, operation: impl FnOnce(&mut Entry) -> R) -> Result<R, ServiceError> {
    let mut guard = store();
    let entry = guard.get_mut(&handle).ok_or_else(not_found)?;
    Ok(operation(entry))
}

/// Removes one entry; closing an unknown handle is a no-op.
pub fn remove(handle: u64) -> bool {
    store().remove(&handle).is_some()
}

fn not_found() -> ServiceError {
    ServiceError::new("workbook/not-found", "no workbook is open for this handle")
}
