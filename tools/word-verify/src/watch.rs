//! UI Automation watcher for Word's modal dialogs on the hidden desktop.
//!
//! With `DisplayAlerts = wdAlertsNone` most open failures come back as COM
//! exceptions, but a few prompts still appear (unreadable-content recovery,
//! the safe-mode offer after a crash, conversion prompts). Any of them would
//! block the STA call forever, so every non-document window owned by Word is
//! recorded and dismissed.

use std::collections::HashSet;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

use serde::Serialize;
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows::Win32::System::Variant::VARIANT;
use windows::core::Interface;
use windows::Win32::UI::Accessibility::{
    CUIAutomation, IUIAutomation, IUIAutomation2, IUIAutomationElement, IUIAutomationInvokePattern, TreeScope_Children,
    TreeScope_Descendants, UIA_InvokePatternId, UIA_ProcessIdPropertyId,
};

/// Text that marks a content problem (English and Czech Office builds).
const PROBLEM_TEXT: &[&str] = &[
    "unreadable content",
    "problems with the contents",
    "problem with the contents",
    "experienced an error trying to open",
    "error trying to open the file",
    "cannot be opened",
    "can't be opened",
    "cannot open",
    "recover the contents",
    "nečitelný obsah",
    "problémy s obsahem",
    "potíže s obsahem",
    "došlo k chybě",
    "nelze otevřít",
    "obnovit obsah",
];
/// The transient progress window Word shows while it opens a file.
const PROGRESS_TEXT: &[&str] = &["opening in protected view", "otevírání v chráněném zobrazení", "opening - microsoft word", "otevírání - microsoft word"];
const SAFE_MODE_TEXT: &[&str] = &["safe mode", "nouzovém režimu", "nouzový režim"];
/// Main document frame and Office chrome that is not a prompt.
const IGNORED_CLASSES: &[&str] = &["OpusApp", "MsoCommandBar", "Net UI Tool Window", "MSO_BORDEREFFECT_WINDOW_CLASS"];

const NO_BUTTONS: &[&str] = &["no", "ne", "nein", "non"];
const YES_BUTTONS: &[&str] = &["yes", "ano", "ja", "oui"];
const OK_BUTTONS: &[&str] = &["ok"];
const CANCEL_BUTTONS: &[&str] = &["cancel", "storno", "abbrechen", "annuler", "close", "zavřít"];

#[derive(Serialize, Clone, Debug)]
pub struct DialogHit {
    /// `problem`, `safe-mode` or `other`.
    pub kind: &'static str,
    pub class: String,
    pub text: String,
    pub clicked: Option<String>,
    #[serde(rename = "atMs")]
    pub at_ms: u128,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Answer {
    /// Decline recovery so the strict open fails with Word's own error.
    Decline,
    /// Accept recovery (used for the repair pass).
    Accept,
}

pub struct WatchConfig {
    pub pid: u32,
    pub timeout: Duration,
    pub answer: Answer,
}

pub fn watch_dialogs(cfg: WatchConfig, stop: Arc<AtomicBool>, out: Arc<Mutex<Vec<DialogHit>>>) {
    if let Err(err) = watch_inner(&cfg, &stop, &out) {
        if let Ok(mut hits) = out.lock() {
            hits.push(DialogHit {
                kind: "other",
                class: String::new(),
                text: format!("uia watcher failed: {err}"),
                clicked: None,
                at_ms: 0,
            });
        }
    }
}

fn watch_inner(cfg: &WatchConfig, stop: &AtomicBool, out: &Mutex<Vec<DialogHit>>) -> windows::core::Result<()> {
    let uia: IUIAutomation = unsafe { CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER)? };
    // A window that is closing can stall a UIA call for the default ~60 s,
    // which then holds up the whole pass. Fail fast and rescan instead.
    if let Ok(uia2) = uia.cast::<IUIAutomation2>() {
        unsafe {
            let _ = uia2.SetConnectionTimeout(2_000);
            let _ = uia2.SetTransactionTimeout(2_000);
        }
    }
    let root = unsafe { uia.GetRootElement()? };
    let cond = if cfg.pid == 0 {
        unsafe { uia.CreateTrueCondition()? }
    } else {
        unsafe { uia.CreatePropertyCondition(UIA_ProcessIdPropertyId, &VARIANT::from(cfg.pid as i32))? }
    };
    let started = Instant::now();
    let deadline = started + cfg.timeout;
    let mut seen: HashSet<String> = HashSet::new();

    while Instant::now() < deadline && !stop.load(Ordering::Relaxed) {
        // Windows appear and vanish mid-scan; UIA then fails for that element
        // (ELEMENTNOTAVAILABLE, TIMEOUT). Skip it and look again next tick.
        if let Ok(windows) = unsafe { root.FindAll(TreeScope_Children, &cond) } {
            let count = unsafe { windows.Length() }.unwrap_or(0);
            for i in 0..count {
                let Ok(window) = (unsafe { windows.GetElement(i) }) else { continue };
                if let Ok(Some(hit)) = inspect_window(&uia, &window, cfg.answer, started) {
                    if seen.insert(format!("{}\n{}", hit.class, hit.text)) {
                        if let Ok(mut hits) = out.lock() {
                            hits.push(hit);
                        }
                    }
                }
            }
        }
        std::thread::sleep(Duration::from_millis(150));
    }
    Ok(())
}

fn inspect_window(
    uia: &IUIAutomation,
    window: &IUIAutomationElement,
    answer: Answer,
    started: Instant,
) -> windows::core::Result<Option<DialogHit>> {
    let class = unsafe { window.CurrentClassName() }.map(|s| s.to_string()).unwrap_or_default();
    if IGNORED_CLASSES.iter().any(|c| class.eq_ignore_ascii_case(c)) {
        return Ok(None);
    }
    let names = descendant_names(uia, window)?;
    if names.is_empty() {
        return Ok(None);
    }
    let text = names.join(" | ");
    let lower = text.to_lowercase();
    let kind = if SAFE_MODE_TEXT.iter().any(|n| lower.contains(n)) {
        "safe-mode"
    } else if PROBLEM_TEXT.iter().any(|n| lower.contains(n)) {
        "problem"
    } else if PROGRESS_TEXT.iter().any(|n| lower.contains(n)) {
        // Word's own "Opening…" progress window: not a prompt, never click it.
        return Ok(None);
    } else {
        "other"
    };
    let order: &[&[&str]] = match (kind, answer) {
        ("safe-mode", _) => &[NO_BUTTONS, CANCEL_BUTTONS],
        ("problem", Answer::Accept) => &[YES_BUTTONS, OK_BUTTONS, CANCEL_BUTTONS],
        _ => &[NO_BUTTONS, OK_BUTTONS, CANCEL_BUTTONS],
    };
    let clicked = dismiss(uia, window, order)?;
    Ok(Some(DialogHit { kind, class, text, clicked, at_ms: started.elapsed().as_millis() }))
}

fn descendant_names(uia: &IUIAutomation, window: &IUIAutomationElement) -> windows::core::Result<Vec<String>> {
    let mut names = Vec::new();
    if let Ok(name) = unsafe { window.CurrentName() } {
        let text = name.to_string();
        if !text.trim().is_empty() {
            names.push(text.trim().to_string());
        }
    }
    let true_cond = unsafe { uia.CreateTrueCondition()? };
    let all = unsafe { window.FindAll(TreeScope_Descendants, &true_cond)? };
    let count = unsafe { all.Length()? };
    for i in 0..count.min(400) {
        let el = unsafe { all.GetElement(i)? };
        if let Ok(name) = unsafe { el.CurrentName() } {
            let text = name.to_string();
            let text = text.trim();
            if !text.is_empty() && !names.iter().any(|n| n == text) {
                names.push(text.to_string());
            }
        }
    }
    Ok(names)
}

fn dismiss(uia: &IUIAutomation, window: &IUIAutomationElement, order: &[&[&str]]) -> windows::core::Result<Option<String>> {
    let true_cond = unsafe { uia.CreateTrueCondition()? };
    let all = unsafe { window.FindAll(TreeScope_Descendants, &true_cond)? };
    let count = unsafe { all.Length()? };
    for group in order {
        for i in 0..count {
            let el = unsafe { all.GetElement(i)? };
            let Ok(current) = (unsafe { el.CurrentName() }) else { continue };
            let text = current.to_string();
            let label = text.trim().trim_start_matches('&').replace('&', "").to_lowercase();
            if group.iter().any(|name| label == *name) {
                if let Ok(pattern) = unsafe { el.GetCurrentPatternAs::<IUIAutomationInvokePattern>(UIA_InvokePatternId) } {
                    unsafe { pattern.Invoke()? };
                    return Ok(Some(text));
                }
            }
        }
    }
    Ok(None)
}
