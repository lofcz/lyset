//! word-verify: open DOCX files in a hidden Microsoft Word and report whether
//! Word accepts them, and if not, exactly what Word said.
//!
//!   word-verify once [--timeout-ms N] [--no-repair] [--pdf] <file.docx>...
//!   word-verify serve [--listen 0.0.0.0:47400] [--token-file <path>]
//!   word-verify version
//!
//! `once` prints a JSON array of results. `serve` answers HTTP requests (see
//! README.md) so a host outside the VM can submit files over a forwarded port.

mod diff;
mod dispatch;
mod isolate;
mod mute;
mod serve;
mod watch;

use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use serde::Serialize;
use windows::core::{w, GUID};
use windows::Win32::System::Com::{CoCreateInstance, CoInitializeEx, CoUninitialize, IDispatch, CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED};
use windows::Win32::System::Registry::{RegDeleteTreeW, RegSetKeyValueW, HKEY_CURRENT_USER, REG_DWORD};

use crate::diff::{diff_docx, PackageDiff};
use crate::dispatch::{
    as_dispatch, as_i32, as_string, get_i32, method, prop_get, prop_put, vt_bool, vt_bstr, vt_i4, vt_missing, ComFailure,
};
use crate::isolate::Isolation;
use crate::mute::AudioMute;
use crate::watch::{watch_dialogs, Answer, DialogHit, WatchConfig};

/// Word.Application
const WORD_CLSID: GUID = GUID::from_u128(0x0002_09FF_0000_0000_C000_0000_0000_0046);
const WD_ALERTS_NONE: i32 = 0;
const WD_DO_NOT_SAVE_CHANGES: i32 = 0;
const WD_FORMAT_XML_DOCUMENT: i32 = 12;
const WD_EXPORT_FORMAT_PDF: i32 = 17;
const WD_STATISTIC_PAGES: i32 = 2;
const WD_STATISTIC_CHARACTERS: i32 = 3;
const MSO_AUTOMATION_SECURITY_FORCE_DISABLE: i32 = 3;
pub(crate) const VERSION: &str = env!("CARGO_PKG_VERSION");

// ---------------------------------------------------------------------------
// Result shapes (the JSON contract consumed by the Linux driver)
// ---------------------------------------------------------------------------

#[derive(Serialize, Clone, Debug, Default)]
struct DocStats {
    pages: Option<i32>,
    characters: Option<i32>,
    paragraphs: Option<i32>,
    tables: Option<i32>,
    #[serde(rename = "inlineShapes")]
    inline_shapes: Option<i32>,
    shapes: Option<i32>,
    #[serde(rename = "oMaths")]
    omaths: Option<i32>,
    fields: Option<i32>,
    sections: Option<i32>,
    /// First COM failure while reading statistics (a modal dialog blocking the
    /// object model shows up here instead of as silently missing numbers).
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<ComFailure>,
}

#[derive(Serialize, Clone, Debug, Default)]
struct PassResult {
    opened: bool,
    /// Protected View only: whether "Enable Editing" succeeded.
    edited: Option<bool>,
    error: Option<ComFailure>,
    name: Option<String>,
    stats: Option<DocStats>,
    /// Closing the document failed even after retries: the next pass would
    /// then see "already open" (6302) instead of the file's own behaviour.
    #[serde(rename = "closeError", skip_serializing_if = "Option::is_none")]
    close_error: Option<ComFailure>,
    #[serde(rename = "timedOut")]
    timed_out: bool,
    ms: u128,
}

#[derive(Serialize, Clone, Debug, Default)]
pub(crate) struct FileResult {
    pub(crate) file: String,
    /// ok | repair | reject | timeout | error
    verdict: String,
    strict: PassResult,
    /// Protected View open: what a teacher gets for a downloaded file.
    #[serde(rename = "protectedView")]
    protected_view: Option<PassResult>,
    repair: Option<PassResult>,
    dialogs: Vec<DialogHit>,
    signals: Vec<String>,
    #[serde(rename = "packageDiff")]
    package_diff: Option<PackageDiff>,
    /// Word's own PDF rendering (strict pass), when requested.
    pub(crate) pdf: Option<String>,
    /// Word's repaired save, when the repair pass recovered the file.
    #[serde(rename = "repairedCopy")]
    pub(crate) repaired_copy: Option<String>,
    ms: u128,
}

#[derive(Serialize, Clone, Debug, Default)]
pub(crate) struct WordInfo {
    version: Option<String>,
    build: Option<String>,
}

pub(crate) fn default_timeout_ms() -> u64 {
    45_000
}

#[derive(Clone, Copy)]
pub(crate) struct Options {
    pub(crate) timeout: Duration,
    pub(crate) repair: bool,
    pub(crate) pdf: bool,
    pub(crate) protected_view: bool,
}

// ---------------------------------------------------------------------------
// Word session
// ---------------------------------------------------------------------------

struct WordSession {
    app: IDispatch,
    spawned: Vec<u32>,
    /// WINWORD pids that existed before this session (never killed).
    preexisting: Vec<u32>,
    info: WordInfo,
}

static LOG_TAIL: Mutex<std::collections::VecDeque<String>> = Mutex::new(std::collections::VecDeque::new());

pub(crate) fn log(msg: impl AsRef<str>) {
    let ms = now_ms();
    let secs = (ms / 1000) % 86_400;
    let line = format!("[{:02}:{:02}:{:02}.{:03}Z] {}", secs / 3600, (secs / 60) % 60, secs % 60, ms % 1000, msg.as_ref());
    eprintln!("word-verify {line}");
    if let Ok(mut tail) = LOG_TAIL.lock() {
        if tail.len() >= 400 {
            tail.pop_front();
        }
        tail.push_back(line);
    }
}

pub(crate) fn log_tail() -> Vec<String> {
    LOG_TAIL.lock().map(|t| t.iter().cloned().collect()).unwrap_or_default()
}

impl WordSession {
    fn start(isolation: &mut Isolation) -> Result<Self, String> {
        clear_resiliency();
        let before = toolhelp_pids("WINWORD.EXE");

        // Watch the hidden desktop while Word boots: a safe-mode prompt there
        // would block CoCreateInstance indefinitely.
        let stop = Arc::new(AtomicBool::new(false));
        let hits = Arc::new(Mutex::new(Vec::new()));
        let watcher = spawn_watcher(0, Duration::from_secs(90), Answer::Decline, stop.clone(), hits.clone());
        let killer = spawn_watchdog(Duration::from_secs(75), stop.clone(), Arc::new(Mutex::new(Vec::new())), Some(before.clone()));

        let created = unsafe { CoCreateInstance::<_, IDispatch>(&WORD_CLSID, None, CLSCTX_LOCAL_SERVER) };
        stop.store(true, Ordering::Relaxed);
        let _ = watcher.join();
        let _ = killer.join();
        for hit in hits.lock().map(|h| h.clone()).unwrap_or_default() {
            log(format!("startup dialog [{}] {}", hit.kind, hit.text));
        }
        let app = created.map_err(|err| format!("CoCreateInstance Word.Application: {err}"))?;

        let mut spawned = Vec::new();
        for _ in 0..40 {
            spawned = toolhelp_pids("WINWORD.EXE").into_iter().filter(|p| !before.contains(p)).collect();
            if !spawned.is_empty() {
                break;
            }
            std::thread::sleep(Duration::from_millis(100));
        }
        if let Some(&pid) = spawned.first() {
            if let Err(err) = isolation.adopt_pid(pid) {
                log(format!("adopt pid {pid} skipped: {err}"));
            }
        }

        // Visible on the private desktop so prompts render for UIA; never on screen.
        let _ = prop_put(&app, "Visible", vt_bool(true));
        let _ = prop_put(&app, "DisplayAlerts", vt_i4(WD_ALERTS_NONE));
        let _ = prop_put(&app, "AutomationSecurity", vt_i4(MSO_AUTOMATION_SECURITY_FORCE_DISABLE));
        if let Ok(options) = prop_get(&app, "Options").and_then(|v| as_dispatch(&v)) {
            let _ = prop_put(&options, "ConfirmConversions", vt_bool(false));
            let _ = prop_put(&options, "UpdateLinksAtOpen", vt_bool(false));
            let _ = prop_put(&options, "CheckGrammarAsYouType", vt_bool(false));
            let _ = prop_put(&options, "CheckSpellingAsYouType", vt_bool(false));
            let _ = prop_put(&options, "SaveNormalPrompt", vt_bool(false));
        }
        let info = WordInfo {
            version: prop_get(&app, "Version").ok().and_then(|v| as_string(&v)),
            build: prop_get(&app, "Build").ok().and_then(|v| as_string(&v)),
        };
        log(format!("Word {} build {} pid={spawned:?}", info.version.as_deref().unwrap_or("?"), info.build.as_deref().unwrap_or("?")));
        Ok(Self { app, spawned, preexisting: before, info })
    }

    fn shutdown(self, isolation: &mut Isolation) {
        let _ = method(&self.app, "Quit", vec![vt_i4(WD_DO_NOT_SAVE_CHANGES)]);
        drop(self.app);
        let deadline = Instant::now() + Duration::from_secs(8);
        while Instant::now() < deadline && self.spawned.iter().any(|p| toolhelp_pids("WINWORD.EXE").contains(p)) {
            std::thread::sleep(Duration::from_millis(150));
        }
        for pid in toolhelp_pids("WINWORD.EXE").into_iter().filter(|p| !self.preexisting.contains(p)) {
            terminate_pid(pid);
        }
        isolation.release();
    }
}

/// Word keeps crash bookkeeping here. A leftover entry makes the next start
/// offer safe mode (a blocking prompt) or disable add-ins.
fn clear_resiliency() {
    unsafe {
        let _ = RegDeleteTreeW(HKEY_CURRENT_USER, w!("Software\\Microsoft\\Office\\16.0\\Word\\Resiliency"));
        // Office asks once which default file types to use (OOXML vs ODF),
        // e.g. after a language pack install. The prompt is modal and blocks
        // every automation call ("Call was rejected by callee") until answered.
        let shown: u32 = 1;
        let _ = RegSetKeyValueW(
            HKEY_CURRENT_USER,
            w!("Software\\Microsoft\\Office\\16.0\\Common\\General"),
            w!("ShownFileFmtPrompt"),
            REG_DWORD.0,
            Some(&shown as *const u32 as *const core::ffi::c_void),
            4,
        );
    }
}

fn spawn_watcher(
    pid: u32,
    timeout: Duration,
    answer: Answer,
    stop: Arc<AtomicBool>,
    hits: Arc<Mutex<Vec<DialogHit>>>,
) -> std::thread::JoinHandle<()> {
    std::thread::spawn(move || {
        let _ = Isolation::bind_current_thread();
        unsafe {
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        }
        watch_dialogs(WatchConfig { pid, timeout, answer }, stop, hits);
        unsafe { CoUninitialize() };
    })
}

/// Kill Word if the blocking call is still running at the deadline. With
/// `new_since` set, kills every WINWORD that was not running before (startup).
fn spawn_watchdog(
    timeout: Duration,
    stop: Arc<AtomicBool>,
    pids: Arc<Mutex<Vec<u32>>>,
    new_since: Option<Vec<u32>>,
) -> std::thread::JoinHandle<bool> {
    std::thread::spawn(move || {
        let deadline = Instant::now() + timeout;
        while Instant::now() < deadline {
            if stop.load(Ordering::Relaxed) {
                return false;
            }
            std::thread::sleep(Duration::from_millis(100));
        }
        if stop.load(Ordering::Relaxed) {
            return false;
        }
        let targets: Vec<u32> = match &new_since {
            Some(before) => toolhelp_pids("WINWORD.EXE").into_iter().filter(|p| !before.contains(p)).collect(),
            None => pids.lock().map(|p| p.clone()).unwrap_or_default(),
        };
        for pid in targets {
            terminate_pid(pid);
        }
        true
    })
}

// ---------------------------------------------------------------------------
// One file
// ---------------------------------------------------------------------------

fn open_document(app: &IDispatch, path: &str, repair: bool) -> Result<IDispatch, ComFailure> {
    let documents = as_dispatch(&prop_get(app, "Documents")?)?;
    // Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
    //   PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument,
    //   WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair,
    //   DocumentDirection, NoEncodingDialog)
    let doc = method(
        &documents,
        "Open",
        vec![
            vt_bstr(path),
            vt_bool(false),
            // Read-write, like a double-click. The service only ever opens its
            // own temp copy, so nothing user-owned can be modified.
            vt_bool(false),
            vt_bool(false),
            vt_missing(),
            vt_missing(),
            vt_bool(false),
            vt_missing(),
            vt_missing(),
            vt_i4(0),
            vt_missing(),
            vt_bool(true),
            vt_bool(repair),
            vt_missing(),
            vt_bool(true),
        ],
    )?;
    as_dispatch(&doc)
}

/// Open in Protected View, the sandboxed read-only path Word uses for files
/// with a mark of the web (every browser download). Returns the PV window.
fn open_protected(app: &IDispatch, path: &str) -> Result<IDispatch, ComFailure> {
    let windows = as_dispatch(&prop_get(app, "ProtectedViewWindows")?)?;
    // ProtectedViewWindows.Open(FileName, AddToRecentFiles, PasswordDocument, Visible, OpenAndRepair)
    let window = method(&windows, "Open", vec![vt_bstr(path), vt_bool(false), vt_missing(), vt_bool(true), vt_bool(false)])?;
    as_dispatch(&window)
}

fn collect_stats(doc: &IDispatch) -> DocStats {
    let statistic = |kind: i32| method(doc, "ComputeStatistics", vec![vt_i4(kind)]).ok().and_then(|v| as_i32(&v));
    let probe = method(doc, "ComputeStatistics", vec![vt_i4(WD_STATISTIC_PAGES)]).err();
    DocStats {
        error: probe,
        pages: statistic(WD_STATISTIC_PAGES),
        characters: statistic(WD_STATISTIC_CHARACTERS),
        paragraphs: get_i32(doc, &["Paragraphs", "Count"]),
        tables: get_i32(doc, &["Tables", "Count"]),
        inline_shapes: get_i32(doc, &["InlineShapes", "Count"]),
        shapes: get_i32(doc, &["Shapes", "Count"]),
        omaths: get_i32(doc, &["OMaths", "Count"]),
        fields: get_i32(doc, &["Fields", "Count"]),
        sections: get_i32(doc, &["Sections", "Count"]),
    }
}

struct PassOutcome {
    pass: PassResult,
    dialogs: Vec<DialogHit>,
    disconnected: bool,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum PassKind {
    Strict,
    Protected,
    Repair,
}

fn run_pass(
    session: &WordSession,
    path: &Path,
    kind: PassKind,
    timeout: Duration,
    after_open: &mut dyn FnMut(&IDispatch),
) -> PassOutcome {
    let started = Instant::now();
    let stop = Arc::new(AtomicBool::new(false));
    let hits = Arc::new(Mutex::new(Vec::new()));
    // Everything on the private desktop is ours. Protected View parses in a
    // separate sandboxed WINWORD process, so do not filter by pid.
    let answer = if kind == PassKind::Repair { Answer::Accept } else { Answer::Decline };
    let watcher = spawn_watcher(0, timeout + Duration::from_secs(5), answer, stop.clone(), hits.clone());
    std::thread::sleep(Duration::from_millis(150));
    let watchdog = spawn_watchdog(timeout, stop.clone(), Arc::new(Mutex::new(Vec::new())), Some(session.preexisting.clone()));

    let mut pass = PassResult::default();
    let path_str = path.to_string_lossy().to_string();
    let opened = match kind {
        PassKind::Protected => open_protected(&session.app, &path_str).and_then(|window| {
            let doc = as_dispatch(&prop_get(&window, "Document")?)?;
            Ok((doc, Some(window)))
        }),
        _ => open_document(&session.app, &path_str, kind == PassKind::Repair).map(|doc| (doc, None)),
    };
    match opened {
        Ok((doc, window)) => {
            pass.opened = true;
            pass.name = prop_get(&doc, "Name").ok().and_then(|v| as_string(&v));
            pass.stats = Some(collect_stats(&doc));
            after_open(&doc);
            match window {
                Some(window) => {
                    // "Enable Editing": Word re-opens the file outside the
                    // sandbox. Teachers click it right after opening a download.
                    match method(&window, "Edit", vec![vt_missing(), vt_missing()]).and_then(|d| as_dispatch(&d)) {
                        Ok(edited) => {
                            pass.edited = Some(true);
                            pass.close_error = close_retrying(&edited);
                        }
                        Err(err) => {
                            pass.edited = Some(false);
                            pass.error = Some(err);
                            let _ = method(&window, "Close", Vec::new());
                        }
                    }
                }
                None => {
                    pass.close_error = close_retrying(&doc);
                }
            }
        }
        Err(err) => pass.error = Some(err),
    }
    stop.store(true, Ordering::Relaxed);
    let timed_out = watchdog.join().unwrap_or(false);
    let _ = watcher.join();
    pass.timed_out = timed_out;
    pass.ms = started.elapsed().as_millis();
    let disconnected = timed_out || pass.error.as_ref().is_some_and(|e| e.is_disconnect());
    let dialogs = hits.lock().map(|h| h.clone()).unwrap_or_default();
    PassOutcome { pass, dialogs, disconnected }
}

/// Close without saving. Word rejects calls while it is busy (background
/// proofing, layout), so retry briefly before reporting the failure.
fn close_retrying(doc: &IDispatch) -> Option<ComFailure> {
    let mut last = None;
    for attempt in 0..6 {
        match method(doc, "Close", vec![vt_i4(WD_DO_NOT_SAVE_CHANGES)]) {
            Ok(_) => return None,
            Err(err) => {
                if err.is_disconnect() {
                    return Some(err);
                }
                last = Some(err);
                std::thread::sleep(Duration::from_millis(250 * (attempt + 1)));
            }
        }
    }
    if let Some(err) = last.as_mut() {
        let windows = isolate::desktop_windows();
        if !windows.is_empty() {
            err.description = format!("{} | windows: {}", err.description, windows.join("; "));
        }
    }
    last
}

/// Verify one local file. `outputs` is where the PDF and repaired copy go.
fn verify_file(
    session: &mut Option<WordSession>,
    isolation: &mut Isolation,
    file: &Path,
    outputs: &Path,
    opts: Options,
    info: &mut WordInfo,
) -> FileResult {
    let started = Instant::now();
    let mut result = FileResult { file: file.to_string_lossy().to_string(), ..Default::default() };
    let stem = file.file_stem().map(|s| s.to_string_lossy().to_string()).unwrap_or_else(|| "document".into());

    if session.is_none() {
        match WordSession::start(isolation) {
            Ok(s) => {
                *info = s.info.clone();
                *session = Some(s);
            }
            Err(err) => {
                result.verdict = "error".into();
                result.signals.push(format!("word-start-failed: {err}"));
                return result;
            }
        }
    }

    // Strict pass: what a teacher double-clicking the file gets.
    let pdf_path = outputs.join(format!("{stem}.word.pdf"));
    let mut pdf_written = false;
    let strict = {
        let s = session.as_ref().expect("session");
        run_pass(s, file, PassKind::Strict, opts.timeout, &mut |doc| {
            if opts.pdf {
                let target = pdf_path.to_string_lossy().to_string();
                pdf_written = method(doc, "ExportAsFixedFormat", vec![vt_bstr(&target), vt_i4(WD_EXPORT_FORMAT_PDF), vt_bool(false)]).is_ok();
            }
        })
    };
    if strict.disconnected {
        restart(session, isolation);
    }
    result.dialogs.extend(strict.dialogs.iter().cloned());
    let strict_problem = strict.dialogs.iter().any(|d| d.kind == "problem");
    result.strict = strict.pass;
    if pdf_written {
        result.pdf = Some(pdf_path.to_string_lossy().to_string());
    }

    // Protected View pass: teachers open downloads, which Word sandboxes.
    let mut protected_ok = true;
    if opts.protected_view {
        if session.is_none() {
            if let Ok(s) = WordSession::start(isolation) {
                *session = Some(s);
            }
        }
        if let Some(s) = session.as_ref() {
            let outcome = run_pass(s, file, PassKind::Protected, opts.timeout, &mut |_| {});
            if outcome.disconnected {
                restart(session, isolation);
            }
            let problem = outcome.dialogs.iter().any(|d| d.kind == "problem");
            protected_ok = outcome.pass.opened && outcome.pass.edited != Some(false) && !problem;
            if outcome.pass.edited == Some(false) {
                result.signals.push("protected-view-enable-editing-failed".into());
            }
            if problem {
                result.signals.push("protected-view-problem-dialog".into());
            }
            if outcome.pass.error.is_some() {
                result.signals.push("protected-view-open-failed".into());
            }
            if outcome.pass.timed_out {
                result.signals.push("protected-view-timeout".into());
            }
            result.dialogs.extend(outcome.dialogs);
            result.protected_view = Some(outcome.pass);
        }
    }

    let strict_ok = result.strict.opened && !strict_problem;
    if !(strict_ok && protected_ok) && opts.repair {
        if session.is_none() {
            if let Ok(s) = WordSession::start(isolation) {
                *session = Some(s);
            }
        }
        if let Some(s) = session.as_ref() {
            let repaired_path = outputs.join(format!("{stem}.word-repaired.docx"));
            let mut saved = false;
            let outcome = run_pass(s, file, PassKind::Repair, opts.timeout, &mut |doc| {
                let target = repaired_path.to_string_lossy().to_string();
                saved = method(doc, "SaveAs2", vec![vt_bstr(&target), vt_i4(WD_FORMAT_XML_DOCUMENT)]).is_ok();
            });
            if outcome.disconnected {
                restart(session, isolation);
            }
            result.dialogs.extend(outcome.dialogs);
            if saved {
                result.package_diff = diff_docx(file, &repaired_path).ok();
                result.repaired_copy = Some(repaired_path.to_string_lossy().to_string());
            }
            result.repair = Some(outcome.pass);
        }
    }

    // Signals and verdict.
    if let Some(err) = &result.strict.error {
        result.signals.push("strict-open-failed".into());
        if err.is_disconnect() {
            result.signals.push("word-disconnected".into());
        }
    }
    if strict_problem {
        result.signals.push("problem-dialog".into());
    }
    if result.dialogs.iter().any(|d| d.kind == "safe-mode") {
        result.signals.push("safe-mode-dialog".into());
    }
    if result.dialogs.iter().any(|d| d.kind == "other") {
        result.signals.push("unclassified-dialog".into());
    }
    if result.strict.timed_out {
        result.signals.push("strict-timeout".into());
    }
    let repaired_ok = result.repair.as_ref().is_some_and(|r| r.opened);
    let any_timeout = result.strict.timed_out || result.protected_view.as_ref().is_some_and(|p| p.timed_out);
    result.verdict = if strict_ok && protected_ok {
        "ok"
    } else if any_timeout && !repaired_ok {
        "timeout"
    } else if repaired_ok {
        "repair"
    } else {
        "reject"
    }
    .into();
    result.ms = started.elapsed().as_millis();
    result
}

fn restart(session: &mut Option<WordSession>, isolation: &mut Isolation) {
    if let Some(s) = session.take() {
        log("Word disconnected or timed out; restarting it");
        s.shutdown(isolation);
    }
    // Give COM a moment to notice the dead server before the next CoCreateInstance.
    std::thread::sleep(Duration::from_millis(500));
}

// ---------------------------------------------------------------------------
// Modes
// ---------------------------------------------------------------------------

pub(crate) struct Runtime {
    isolation: Isolation,
    mute: AudioMute,
    session: Option<WordSession>,
    pub(crate) info: WordInfo,
}

impl Runtime {
    pub(crate) fn new() -> Result<Self, String> {
        let isolation = Isolation::new().map_err(|err| format!("create hidden desktop: {err}"))?;
        Isolation::bind_current_thread().map_err(|err| format!("bind hidden desktop: {err}"))?;
        let mute = AudioMute::start();
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok() }.map_err(|err| format!("CoInitializeEx: {err}"))?;
        Ok(Self { isolation, mute, session: None, info: WordInfo::default() })
    }

    pub(crate) fn run_files(&mut self, files: &[PathBuf], outputs: &Path, opts: Options) -> Vec<FileResult> {
        let mut out = Vec::new();
        for file in files {
            let r = verify_file(&mut self.session, &mut self.isolation, file, outputs, opts, &mut self.info);
            if let Some(s) = &self.session {
                self.mute.track_pids(s.spawned.iter().copied());
            }
            log(format!("{} -> {} ({} ms)", file.display(), r.verdict, r.ms));
            out.push(r);
        }
        out
    }

    pub(crate) fn close(&mut self) {
        if let Some(s) = self.session.take() {
            s.shutdown(&mut self.isolation);
        }
    }
}

impl Drop for Runtime {
    fn drop(&mut self) {
        self.close();
        unsafe { CoUninitialize() };
    }
}

fn run_once(args: &[String]) -> Result<(), String> {
    let mut opts = Options { timeout: Duration::from_millis(default_timeout_ms()), repair: true, pdf: false, protected_view: true };
    let mut files = Vec::new();
    let mut it = args.iter();
    while let Some(arg) = it.next() {
        match arg.as_str() {
            "--timeout-ms" => opts.timeout = Duration::from_millis(it.next().ok_or("--timeout-ms needs a value")?.parse().map_err(|e| format!("{e}"))?),
            "--no-repair" => opts.repair = false,
            "--no-protected-view" => opts.protected_view = false,
            "--pdf" => opts.pdf = true,
            flag if flag.starts_with('-') => return Err(format!("unknown flag {flag}")),
            file => files.push(absolute(Path::new(file))),
        }
    }
    if files.is_empty() {
        return Err("pass one or more .docx paths".into());
    }
    let outputs = files[0].parent().map(Path::to_path_buf).unwrap_or_else(std::env::temp_dir);
    let mut runtime = Runtime::new()?;
    let results = runtime.run_files(&files, &outputs, opts);
    runtime.close();
    println!("{}", serde_json::to_string(&results).map_err(|e| e.to_string())?);
    Ok(())
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

pub(crate) fn now_ms() -> u128 {
    SystemTime::now().duration_since(UNIX_EPOCH).map(|d| d.as_millis()).unwrap_or(0)
}

fn absolute(path: &Path) -> PathBuf {
    let abs = std::fs::canonicalize(path).unwrap_or_else(|_| path.to_path_buf());
    PathBuf::from(abs.to_string_lossy().trim_start_matches(r"\\?\").to_string())
}

fn toolhelp_pids(name: &str) -> Vec<u32> {
    use windows::Win32::Foundation::CloseHandle;
    use windows::Win32::System::Diagnostics::ToolHelp::{
        CreateToolhelp32Snapshot, Process32FirstW, Process32NextW, PROCESSENTRY32W, TH32CS_SNAPPROCESS,
    };
    let mut pids = Vec::new();
    unsafe {
        let Ok(snap) = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) else { return pids };
        let mut entry = PROCESSENTRY32W { dwSize: std::mem::size_of::<PROCESSENTRY32W>() as u32, ..Default::default() };
        if Process32FirstW(snap, &mut entry).is_ok() {
            loop {
                let exe = String::from_utf16_lossy(entry.szExeFile.split(|c| *c == 0).next().unwrap_or(&[]));
                if exe.eq_ignore_ascii_case(name) {
                    pids.push(entry.th32ProcessID);
                }
                if Process32NextW(snap, &mut entry).is_err() {
                    break;
                }
            }
        }
        let _ = CloseHandle(snap);
    }
    pids
}

fn terminate_pid(pid: u32) {
    use windows::Win32::Foundation::CloseHandle;
    use windows::Win32::System::Threading::{OpenProcess, TerminateProcess, PROCESS_TERMINATE};
    unsafe {
        if let Ok(handle) = OpenProcess(PROCESS_TERMINATE, false, pid) {
            let _ = TerminateProcess(handle, 1);
            let _ = CloseHandle(handle);
        }
    }
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let (mode, rest) = args.split_first().map(|(m, r)| (m.as_str(), r)).unwrap_or(("", &[]));
    let outcome = match mode {
        "once" => run_once(rest),
        "serve" => serve::run(rest),
        "version" => {
            println!("word-verify {VERSION}");
            Ok(())
        }
        _ => Err("usage: word-verify once [--timeout-ms N] [--no-repair] [--pdf] <file.docx>... | serve [--listen ADDR] [--token-file PATH] | version".into()),
    };
    if let Err(err) = outcome {
        log(err);
        std::process::exit(2);
    }
}
