//! Thin Node-API surface for OpenExcel.
//!
//! Every export does one orthogonal operation and returns a JSON string, so the
//! TypeScript layer owns argument parsing, help text and exit codes through
//! `@wae/commander`. Workbook semantics stay in the model crate; this layer only
//! maps values and hosts the WebView window.
//!
//! Functions that are still unwired are forwarded to the model crate and return
//! the stable `capability/not-implemented` code instead of a fabricated workbook
//! or report.

#![deny(missing_docs)]

use std::{
    cell::RefCell,
    path::{Path, PathBuf},
    sync::{
        Arc,
        atomic::{AtomicBool, Ordering},
    },
    time::Instant,
};

use open_excel_workbook::ServiceError;
use serde_json::{Value, json};
use wae::{
    CallContext, Error, JsValue, Result, function, optional_string_prop, register_module,
    webview::{IpcHandler, IpcOutcome, IpcSink, OpenOptions, WebViewError, WebViewHost},
};

fn fail(error: ServiceError) -> Error {
    Error::new(format!("{}: {}", error.code, error.message))
}

fn argument(ctx: &CallContext, index: usize) -> Result<JsValue> {
    ctx.args().get(index).copied().ok_or_else(|| Error::new("session/argument-invalid: a required argument is missing"))
}

fn text(ctx: &CallContext, index: usize) -> Result<String> {
    argument(ctx, index)?.as_string()
}

fn handle(ctx: &CallContext, index: usize) -> Result<u64> {
    let raw = argument(ctx, index)?.as_i64()?;
    u64::try_from(raw).map_err(|_| Error::new("session/argument-invalid: the workbook handle must be a non-negative integer"))
}

fn respond(ctx: &CallContext, value: &Value) -> Result<JsValue> {
    let text = value.to_string();
    ctx.env().string(&text)
}

fn parse_json(raw: &str) -> Result<Value> {
    serde_json::from_str(raw).map_err(|error| Error::new(format!("session/argument-invalid: {error}")))
}

/// Library version of the native binding.
#[function]
fn version(ctx: CallContext) -> Result<JsValue> {
    ctx.env().string(env!("CARGO_PKG_VERSION"))
}

/// Reports formats, checks, renderers and export targets of the engine.
///
/// The matrix comes from the same model crate the CLI uses, so neither entry
/// point can report different capabilities.
#[function]
fn capabilities(ctx: CallContext) -> Result<JsValue> {
    let matrix = open_excel_workbook::capabilities().to_string();
    ctx.env().string(&matrix)
}

/// Creates a workbook from a JSON request and returns `{ handle, workbook }`.
#[function]
fn workbook_create(ctx: CallContext) -> Result<JsValue> {
    let raw = text(&ctx, 0)?;
    let request = parse_json(&raw)?;
    let opened = open_excel_workbook::create(&request).map_err(fail)?;
    respond(&ctx, &opened)
}

/// Opens a workbook project from a path and returns `{ handle, workbook }`.
#[function]
fn workbook_open(ctx: CallContext) -> Result<JsValue> {
    let path = text(&ctx, 0)?;
    let opened = open_excel_workbook::open(Path::new(&path)).map_err(fail)?;
    respond(&ctx, &opened)
}

/// Writes one open workbook to a project path.
#[function]
fn workbook_save(ctx: CallContext) -> Result<JsValue> {
    let handle = handle(&ctx, 0)?;
    let path = text(&ctx, 1)?;
    let overwrite = ctx.args().get(2).map(|value| value.as_bool()).transpose()?.unwrap_or(false);
    let report = open_excel_workbook::save(handle, Path::new(&path), overwrite).map_err(fail)?;
    respond(&ctx, &report)
}

/// Runs the rule checks of one open workbook.
#[function]
fn workbook_inspect(ctx: CallContext) -> Result<JsValue> {
    let handle = handle(&ctx, 0)?;
    let checks = match ctx.args().get(1) {
        Some(value) => {
            let raw = value.as_string()?;
            let names: Vec<String> = parse_json(&raw).and_then(|parsed| {
                parsed
                    .as_array()
                    .map(|items| items.iter().filter_map(Value::as_str).map(str::to_string).collect())
                    .ok_or_else(|| Error::new("session/argument-invalid: `checks` must be a JSON array"))
            })?;
            Some(names)
        }
        None => None,
    };
    let report = open_excel_workbook::inspect(handle, checks).map_err(fail)?;
    respond(&ctx, &report)
}

/// Returns the summary of one open workbook.
#[function]
fn workbook_summary(ctx: CallContext) -> Result<JsValue> {
    let handle = handle(&ctx, 0)?;
    let report = open_excel_workbook::summary(handle).map_err(fail)?;
    respond(&ctx, &report)
}

/// Applies a JSON edit request to one open workbook.
#[function]
fn workbook_edit(ctx: CallContext) -> Result<JsValue> {
    let handle = handle(&ctx, 0)?;
    let raw = text(&ctx, 1)?;
    let request = parse_json(&raw)?;
    let report = open_excel_workbook::edit(handle, &request).map_err(fail)?;
    respond(&ctx, &report)
}

/// Reads one cell with its evaluated value.
#[function]
fn workbook_cell(ctx: CallContext) -> Result<JsValue> {
    let handle = handle(&ctx, 0)?;
    let raw = text(&ctx, 1)?;
    let request = parse_json(&raw)?;
    let detail = open_excel_workbook::cell(handle, &request).map_err(fail)?;
    respond(&ctx, &detail)
}

/// Reads a rectangle of cells with their evaluated values.
#[function]
fn workbook_range(ctx: CallContext) -> Result<JsValue> {
    let handle = handle(&ctx, 0)?;
    let raw = text(&ctx, 1)?;
    let request = parse_json(&raw)?;
    let matrix = open_excel_workbook::range(handle, &request).map_err(fail)?;
    respond(&ctx, &matrix)
}

/// Recomputes every formula of one open workbook.
#[function]
fn workbook_recalc(ctx: CallContext) -> Result<JsValue> {
    let handle = handle(&ctx, 0)?;
    let report = open_excel_workbook::recalculate(handle).map_err(fail)?;
    respond(&ctx, &report)
}

/// Searches evaluated values and formula sources of one open workbook.
#[function]
fn workbook_find(ctx: CallContext) -> Result<JsValue> {
    let handle = handle(&ctx, 0)?;
    let raw = text(&ctx, 1)?;
    let request = parse_json(&raw)?;
    let hits = open_excel_workbook::find(handle, &request).map_err(fail)?;
    respond(&ctx, &hits)
}

/// Renders one sheet of an open workbook as SVG.
#[function]
fn workbook_render(ctx: CallContext) -> Result<JsValue> {
    let handle = handle(&ctx, 0)?;
    let request = match ctx.args().get(1) {
        Some(value) if !value.is_undefined()? => {
            let raw = value.as_string()?;
            parse_json(&raw)?
        }
        _ => json!({}),
    };
    let report = open_excel_workbook::render(handle, &request).map_err(fail)?;
    respond(&ctx, &report)
}

/// Writes one open workbook to an export target.
#[function]
fn workbook_export(ctx: CallContext) -> Result<JsValue> {
    let handle = handle(&ctx, 0)?;
    let format = text(&ctx, 1)?;
    let target = text(&ctx, 2)?;
    let overwrite = ctx.args().get(3).map(|value| value.as_bool()).transpose()?.unwrap_or(false);
    let report = open_excel_workbook::export(handle, &format, Path::new(&target), overwrite).map_err(fail)?;
    respond(&ctx, &report)
}

/// Releases one open workbook handle.
#[function]
fn workbook_close(ctx: CallContext) -> Result<JsValue> {
    let handle = handle(&ctx, 0)?;
    open_excel_workbook::close(handle).map_err(fail)?;
    ctx.env().boolean(true)
}

// One WebView host per process, owned by the thread that opened it.
thread_local! {
    static GUI: RefCell<Option<GuiState>> = const { RefCell::new(None) };
}

const GUI_CLOSED: &str = "gui_closed";
const GUI_HOST_EXITED: &str = "gui_host_exited";
const GUI_NOT_OPEN: &str = "gui_not_open";

struct GuiState {
    host: WebViewHost,
    backend: &'static str,
    opened_at: Instant,
    /// Set when the page asks the host to close, so a dead host is
    /// distinguishable from one that never processed the request.
    close_requested: Arc<AtomicBool>,
}

struct WorkbookIpcHandler {
    close_requested: Arc<AtomicBool>,
}

impl IpcHandler for WorkbookIpcHandler {
    fn on_ipc(&mut self, message: &str, _sink: &IpcSink) -> IpcOutcome {
        let Ok(value) = serde_json::from_str::<Value>(message)
        else {
            return IpcOutcome::Ignored;
        };
        let ty = value.get("type").and_then(Value::as_str).unwrap_or("");
        let id = value.get("id").cloned().unwrap_or(Value::Null);
        match ty {
            "close" => {
                self.close_requested.store(true, Ordering::SeqCst);
                IpcOutcome::CloseWindow
            }
            "ready" => reply(id, "readyResult", json!({ "version": env!("CARGO_PKG_VERSION"), "platform": platform_name() })),
            "capabilities" => reply(id, "capabilitiesResult", json!({ "capabilities": open_excel_workbook::capabilities() })),
            "workbook.create" => {
                let request = value.get("request").cloned().unwrap_or(Value::Null);
                match call_create(&request) {
                    Ok(result) => reply(id, "createResult", result),
                    Err(message) => failure(id, "createError", message),
                }
            }
            "workbook.open" => {
                let path = value.get("path").and_then(Value::as_str).unwrap_or("").trim().to_string();
                if path.is_empty() {
                    return failure(id, "openError", "an open request needs a project path".to_string());
                }
                match call_open(&path) {
                    Ok(result) => reply(id, "openResult", result),
                    Err(message) => failure(id, "openError", message),
                }
            }
            "workbook.save" => {
                let path = value.get("path").and_then(Value::as_str).unwrap_or("").trim().to_string();
                if path.is_empty() {
                    return failure(id, "saveError", "a save request needs a project path".to_string());
                }
                let overwrite = value.get("overwrite").and_then(Value::as_bool).unwrap_or(true);
                match call_save(handle_of(&value), &path, overwrite) {
                    Ok(report) => reply(id, "saveResult", json!({ "report": report })),
                    Err(message) => failure(id, "saveError", message),
                }
            }
            "workbook.inspect" => {
                let checks = value
                    .get("checks")
                    .and_then(Value::as_array)
                    .map(|items| items.iter().filter_map(Value::as_str).map(str::to_string).collect());
                match call_inspect(handle_of(&value), checks) {
                    Ok(report) => reply(id, "inspectResult", json!({ "report": report })),
                    Err(message) => failure(id, "inspectError", message),
                }
            }
            "workbook.summary" => match call_summary(handle_of(&value)) {
                Ok(summary) => reply(id, "summaryResult", json!({ "workbook": summary })),
                Err(message) => failure(id, "summaryError", message),
            },
            "workbook.edit" => {
                let request = value.get("request").cloned().unwrap_or(Value::Null);
                match call_edit(handle_of(&value), &request) {
                    Ok(report) => reply(id, "editResult", report),
                    Err(message) => failure(id, "editError", message),
                }
            }
            "workbook.cell" => {
                let request = value.get("request").cloned().unwrap_or(Value::Null);
                match call_cell(handle_of(&value), &request) {
                    Ok(detail) => reply(id, "cellResult", json!({ "cell": detail })),
                    Err(message) => failure(id, "cellError", message),
                }
            }
            "workbook.range" => {
                let request = value.get("request").cloned().unwrap_or(Value::Null);
                match call_range(handle_of(&value), &request) {
                    Ok(matrix) => reply(id, "rangeResult", json!({ "matrix": matrix })),
                    Err(message) => failure(id, "rangeError", message),
                }
            }
            "workbook.recalc" => match call_recalc(handle_of(&value)) {
                Ok(report) => reply(id, "recalcResult", json!({ "report": report })),
                Err(message) => failure(id, "recalcError", message),
            },
            "workbook.find" => {
                let request = value.get("request").cloned().unwrap_or(Value::Null);
                match call_find(handle_of(&value), &request) {
                    Ok(hits) => reply(id, "findResult", json!({ "hits": hits })),
                    Err(message) => failure(id, "findError", message),
                }
            }
            "workbook.render" => {
                let request = value.get("request").cloned().unwrap_or_else(|| json!({}));
                match call_render(handle_of(&value), &request) {
                    Ok(report) => reply(id, "renderResult", report),
                    Err(message) => failure(id, "renderError", message),
                }
            }
            "workbook.export" => {
                let format = value.get("format").and_then(Value::as_str).unwrap_or("").trim().to_string();
                let target = value.get("target").and_then(Value::as_str).unwrap_or("").trim().to_string();
                if format.is_empty() || target.is_empty() {
                    return failure(id, "exportError", "an export request needs a format and a target path".to_string());
                }
                let overwrite = value.get("overwrite").and_then(Value::as_bool).unwrap_or(true);
                match call_export(handle_of(&value), &format, &target, overwrite) {
                    Ok(report) => reply(id, "exportResult", json!({ "report": report })),
                    Err(message) => failure(id, "exportError", message),
                }
            }
            "workbook.close" => match call_close(handle_of(&value)) {
                Ok(()) => reply(id, "closeResult", json!({ "closed": true })),
                Err(message) => failure(id, "closeError", message),
            },
            _ => IpcOutcome::Ignored,
        }
    }
}

fn reply(id: Value, kind: &str, payload: Value) -> IpcOutcome {
    let mut body = payload;
    body["type"] = Value::String(kind.to_string());
    body["id"] = id;
    IpcOutcome::ReplyToPage(body.to_string())
}

fn failure(id: Value, kind: &str, message: String) -> IpcOutcome {
    reply(id, kind, json!({ "message": message }))
}

fn handle_of(value: &Value) -> u64 {
    value.get("handle").and_then(Value::as_u64).unwrap_or(0)
}

fn service_failure(error: ServiceError) -> String {
    format!("{}: {}", error.code, error.message)
}

fn call_create(request: &Value) -> std::result::Result<Value, String> {
    open_excel_workbook::create(request).map_err(service_failure)
}

fn call_open(path: &str) -> std::result::Result<Value, String> {
    open_excel_workbook::open(Path::new(path)).map_err(service_failure)
}

fn call_save(handle: u64, path: &str, overwrite: bool) -> std::result::Result<Value, String> {
    open_excel_workbook::save(handle, Path::new(path), overwrite).map_err(service_failure)
}

fn call_inspect(handle: u64, checks: Option<Vec<String>>) -> std::result::Result<Value, String> {
    open_excel_workbook::inspect(handle, checks).map_err(service_failure)
}

fn call_export(handle: u64, format: &str, target: &str, overwrite: bool) -> std::result::Result<Value, String> {
    open_excel_workbook::export(handle, format, Path::new(target), overwrite).map_err(service_failure)
}

fn call_summary(handle: u64) -> std::result::Result<Value, String> {
    open_excel_workbook::summary(handle).map_err(service_failure)
}

fn call_edit(handle: u64, request: &Value) -> std::result::Result<Value, String> {
    open_excel_workbook::edit(handle, request).map_err(service_failure)
}

fn call_cell(handle: u64, request: &Value) -> std::result::Result<Value, String> {
    open_excel_workbook::cell(handle, request).map_err(service_failure)
}

fn call_range(handle: u64, request: &Value) -> std::result::Result<Value, String> {
    open_excel_workbook::range(handle, request).map_err(service_failure)
}

fn call_recalc(handle: u64) -> std::result::Result<Value, String> {
    open_excel_workbook::recalculate(handle).map_err(service_failure)
}

fn call_find(handle: u64, request: &Value) -> std::result::Result<Value, String> {
    open_excel_workbook::find(handle, request).map_err(service_failure)
}

fn call_render(handle: u64, request: &Value) -> std::result::Result<Value, String> {
    open_excel_workbook::render(handle, request).map_err(service_failure)
}

fn call_close(handle: u64) -> std::result::Result<(), String> {
    open_excel_workbook::close(handle).map_err(service_failure)
}

fn platform_name() -> &'static str {
    if cfg!(windows) {
        "win32"
    }
    else if cfg!(target_os = "macos") {
        "darwin"
    }
    else {
        "linux"
    }
}

fn gui_status(
    state: &'static str,
    code: Option<&'static str>,
    backend: Option<&'static str>,
    uptime_ms: u64,
    message: Option<String>,
) -> Value {
    json!({
        "state": state,
        "code": code,
        "backend": backend,
        "platform": platform_name(),
        "uptimeMs": uptime_ms,
        "message": message,
    })
}

fn gui_open_result(ok: bool, status: Value, error: Option<&WebViewError>) -> Value {
    let (stage, hresult, attempted) = match error {
        Some(error) => {
            let attempted = match error {
                WebViewError::Host(host) => host.attempted.clone(),
                _ => Vec::new(),
            };
            (error.stage(), error.hresult(), attempted)
        }
        None => (None, None, Vec::new()),
    };
    json!({
        "ok": ok,
        "status": status,
        "stage": stage,
        "hresult": hresult,
        "attempted": attempted,
    })
}

/// Resolves the GUI content directory: explicit override, then the built tree.
fn resolve_gui_dir() -> Option<PathBuf> {
    if let Ok(value) = std::env::var("OPEN_EXCEL_GUI") {
        let path = PathBuf::from(value);
        if path.is_dir() {
            return Some(path);
        }
        if path.is_file() {
            return path.parent().map(Path::to_path_buf);
        }
    }

    let manifest_relative = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("..")
        .join("projects")
        .join("packages")
        .join("open-excel")
        .join("dist")
        .join("gui");
    if manifest_relative.join("index.html").is_file() {
        return Some(manifest_relative);
    }

    if let Ok(cwd) = std::env::current_dir() {
        for candidate in [cwd.join("projects/packages/open-excel/dist/gui"), cwd.join("dist/gui"), cwd.join("gui")] {
            if candidate.join("index.html").is_file() {
                return Some(candidate);
            }
        }
    }
    None
}

/// Opens the OpenExcel window, or focuses the live one.
#[function(js_name = "openGui")]
fn open_gui(ctx: CallContext) -> Result<JsValue> {
    let mut title = "OpenExcel".to_string();
    let mut content_dir: Option<String> = None;
    let mut width = 1280i32;
    let mut height = 820i32;

    if let Some(arg) = ctx.args().first().copied()
        && !arg.is_undefined()?
    {
        if let Some(value) = optional_string_prop(arg, "title")? {
            title = value;
        }
        if let Some(value) = optional_string_prop(arg, "contentDir")? {
            content_dir = Some(value);
        }
        if let Ok(value) = arg.get_named("width")
            && let Ok(number) = value.as_i64()
        {
            width = number as i32;
        }
        if let Ok(value) = arg.get_named("height")
            && let Ok(number) = value.as_i64()
        {
            height = number as i32;
        }
    }

    let live = GUI.with(|slot| {
        let guard = slot.borrow();
        guard
            .as_ref()
            .filter(|state| state.host.is_alive())
            .map(|state| (state.backend, state.opened_at.elapsed().as_millis() as u64))
    });
    if let Some((backend, uptime_ms)) = live {
        GUI.with(|slot| {
            if let Some(state) = slot.borrow_mut().as_mut() {
                let _ = state.host.focus();
            }
        });
        let status = gui_status("open", None, Some(backend), uptime_ms, None);
        return respond(&ctx, &gui_open_result(true, status, None));
    }
    GUI.with(|slot| {
        *slot.borrow_mut() = None;
    });

    let dir = match content_dir {
        Some(dir) => PathBuf::from(dir),
        None => match resolve_gui_dir() {
            Some(dir) => dir,
            None => {
                let status = gui_status(
                    "failed",
                    Some("gui_content_missing"),
                    None,
                    0,
                    Some("the WebView content directory is missing; build the GUI first".to_string()),
                );
                return respond(&ctx, &gui_open_result(false, status, None));
            }
        },
    };

    let close_requested = Arc::new(AtomicBool::new(false));
    let mut options = OpenOptions::from_content_dir(dir.display().to_string(), title);
    options.width = width;
    options.height = height;
    options.accept_file_drop = true;

    match WebViewHost::open(options, Box::new(WorkbookIpcHandler { close_requested: Arc::clone(&close_requested) })) {
        Ok(host) => {
            let backend = host.backend_name();
            GUI.with(|slot| {
                *slot.borrow_mut() = Some(GuiState { host, backend, opened_at: Instant::now(), close_requested });
            });
            let status = gui_status("open", None, Some(backend), 0, None);
            respond(&ctx, &gui_open_result(true, status, None))
        }
        Err(error) => {
            let status = gui_status("failed", error.code(), None, 0, Some(error.to_string()));
            respond(&ctx, &gui_open_result(false, status, Some(&error)))
        }
    }
}

/// Pumps the host message loop once and reports the observed lifecycle state.
#[function(js_name = "pumpEvents")]
fn pump_events(ctx: CallContext) -> Result<JsValue> {
    let status = GUI.with(|slot| {
        let mut guard = slot.borrow_mut();
        let Some(state) = guard.as_mut()
        else {
            return gui_status("closed", Some(GUI_NOT_OPEN), None, 0, Some("no OpenExcel window is open".to_string()));
        };
        let backend = state.backend;
        let uptime_ms = state.opened_at.elapsed().as_millis() as u64;
        if state.host.pump() {
            return gui_status("open", None, Some(backend), uptime_ms, None);
        }
        let requested = state.close_requested.swap(false, Ordering::SeqCst);
        *guard = None;
        if requested {
            gui_status("closed", Some(GUI_CLOSED), Some(backend), uptime_ms, None)
        }
        else {
            gui_status(
                "closed",
                Some(GUI_HOST_EXITED),
                Some(backend),
                uptime_ms,
                Some(format!("the {backend} webview host stopped without a page close request")),
            )
        }
    });
    respond(&ctx, &status)
}

/// Closes the OpenExcel window; the result distinguishes "already closed".
#[function(js_name = "closeGui")]
fn close_gui(ctx: CallContext) -> Result<JsValue> {
    let result = GUI.with(|slot| {
        let mut guard = slot.borrow_mut();
        let Some(state) = guard.as_mut()
        else {
            return json!({
                "closed": false,
                "status": gui_status("not_open", Some(GUI_NOT_OPEN), None, 0, Some("no OpenExcel window is open".to_string())),
            });
        };
        let backend = state.backend;
        let uptime_ms = state.opened_at.elapsed().as_millis() as u64;
        match state.host.close() {
            Ok(()) => {
                *guard = None;
                json!({ "closed": true, "status": gui_status("closed", None, Some(backend), uptime_ms, None) })
            }
            Err(error) => json!({
                "closed": false,
                "status": gui_status("open", Some("gui_close_failed"), Some(backend), uptime_ms, Some(error.to_string())),
            }),
        }
    });
    respond(&ctx, &result)
}

/// Reports whether the OpenExcel window is alive.
#[function(js_name = "isGuiOpen")]
fn is_gui_open(ctx: CallContext) -> Result<JsValue> {
    let open = GUI.with(|slot| slot.borrow().as_ref().is_some_and(|state| state.host.is_alive()));
    ctx.env().boolean(open)
}

register_module![
    version,
    capabilities,
    workbook_create,
    workbook_open,
    workbook_save,
    workbook_inspect,
    workbook_summary,
    workbook_edit,
    workbook_cell,
    workbook_range,
    workbook_recalc,
    workbook_find,
    workbook_render,
    workbook_export,
    workbook_close,
    open_gui,
    pump_events,
    close_gui,
    is_gui_open
];
