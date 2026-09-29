//! HTTP front end for a host outside the VM.
//!
//! The VM forwards a port to the host's loopback, so the host connects in.
//! A shared folder would be simpler, but the Windows SMB client caches the
//! listing and metadata of a share whose files change behind Samba's back,
//! which delayed job pickup by minutes.
//!
//!   GET  /health   status, Word version, sha256 of this exe
//!   GET  /log      recent log lines
//!   POST /verify   {"files":[{"name","data"(base64)}],"repair","pdf","protectedView","timeoutMs"}
//!   POST /update   raw bytes of a new word-verify.exe; the launcher swaps it in
//!
//! Every request needs `Authorization: Bearer <token>` when a token file is set.
//! COM work happens on the main (STA) thread, one request at a time.

use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::{mpsc, Arc, Mutex};
use std::time::Duration;

use base64::engine::general_purpose::STANDARD as B64;
use base64::Engine;
use serde::Deserialize;
use sha2::{Digest, Sha256};
use tiny_http::{Header, Method, Request, Response, Server};

use crate::{default_timeout_ms, log, log_tail, now_ms, Options, Runtime, WordInfo, VERSION};

/// Exit code telling the launcher to swap in `word-verify.new.exe`.
const EXIT_UPDATE: i32 = 75;
const MAX_BODY: usize = 256 * 1024 * 1024;

#[derive(Deserialize)]
struct VerifyRequest {
    files: Vec<InputFile>,
    #[serde(default = "yes")]
    repair: bool,
    #[serde(default)]
    pdf: bool,
    #[serde(rename = "timeoutMs", default = "default_timeout_ms")]
    timeout_ms: u64,
    #[serde(rename = "protectedView", default = "yes")]
    protected_view: bool,
}

#[derive(Deserialize)]
struct InputFile {
    name: String,
    data: String,
}

fn yes() -> bool {
    true
}

enum Work {
    Verify(Request, Vec<u8>),
    Update(Request, Vec<u8>),
}

struct Status {
    busy: bool,
    word: WordInfo,
    served: u64,
}

pub fn run(args: &[String]) -> Result<(), String> {
    let mut listen = "0.0.0.0:47400".to_string();
    let mut token_file: Option<PathBuf> = None;
    let mut it = args.iter();
    while let Some(arg) = it.next() {
        match arg.as_str() {
            "--listen" => listen = it.next().ok_or("--listen needs a value")?.clone(),
            "--token-file" => token_file = Some(PathBuf::from(it.next().ok_or("--token-file needs a value")?)),
            other => return Err(format!("unknown argument {other}")),
        }
    }
    let token = match &token_file {
        Some(path) => Some(std::fs::read_to_string(path).map_err(|e| format!("read {}: {e}", path.display()))?.trim().to_string()),
        None => None,
    }
    .filter(|t| !t.is_empty());
    if token.is_none() {
        log("no token file: requests are not authenticated");
    }
    let exe_sha = std::env::current_exe().ok().and_then(|p| std::fs::read(p).ok()).map(|b| hex(&Sha256::digest(&b))).unwrap_or_default();
    let started_at = now_ms();
    let server = Server::http(&listen).map_err(|e| format!("listen {listen}: {e}"))?;
    log(format!("serving on {listen}, exe sha256 {}", &exe_sha[..12.min(exe_sha.len())]));

    let status = Arc::new(Mutex::new(Status { busy: false, word: WordInfo::default(), served: 0 }));
    let (tx, rx) = mpsc::channel::<Work>();
    {
        let status = status.clone();
        let exe_sha = exe_sha.clone();
        std::thread::spawn(move || accept_loop(server, token, exe_sha, started_at, status, tx));
    }

    let mut runtime = Runtime::new()?;
    for work in rx {
        match work {
            Work::Verify(req, body) => {
                set_busy(&status, true);
                let reply = verify(&mut runtime, &body, &exe_sha);
                if let Ok(mut s) = status.lock() {
                    s.busy = false;
                    s.word = runtime.info.clone();
                    s.served += 1;
                }
                match reply {
                    Ok(json) => respond_json(req, 200, &json),
                    Err(err) => respond_json(req, 400, &serde_json::json!({ "error": err })),
                }
            }
            Work::Update(req, body) => {
                let target = std::env::current_exe().map(|p| p.with_file_name("word-verify.new.exe"));
                let written = target.map_err(|e| e.to_string()).and_then(|t| std::fs::write(&t, &body).map(|_| t).map_err(|e| e.to_string()));
                match written {
                    Ok(path) => {
                        log(format!("update staged at {} ({} bytes); restarting", path.display(), body.len()));
                        respond_json(req, 200, &serde_json::json!({ "staged": true, "sha256": hex(&Sha256::digest(&body)) }));
                        runtime.close();
                        std::thread::sleep(Duration::from_millis(300));
                        std::process::exit(EXIT_UPDATE);
                    }
                    Err(err) => respond_json(req, 500, &serde_json::json!({ "error": err })),
                }
            }
        }
    }
    Ok(())
}

fn accept_loop(server: Server, token: Option<String>, exe_sha: String, started_at: u128, status: Arc<Mutex<Status>>, tx: mpsc::Sender<Work>) {
    for mut req in server.incoming_requests() {
        if let Some(expected) = &token {
            let ok = req
                .headers()
                .iter()
                .any(|h| h.field.equiv("Authorization") && h.value.as_str().strip_prefix("Bearer ").is_some_and(|t| t.trim() == expected));
            if !ok {
                respond_json(req, 401, &serde_json::json!({ "error": "missing or wrong bearer token" }));
                continue;
            }
        }
        let path = req.url().split('?').next().unwrap_or("").to_string();
        match (req.method().clone(), path.as_str()) {
            (Method::Get, "/health") => {
                let body = {
                    let s = status.lock().ok();
                    serde_json::json!({
                        "tool": format!("word-verify {VERSION}"),
                        "exeSha256": exe_sha,
                        "startedAt": started_at,
                        "now": now_ms(),
                        "busy": s.as_ref().map(|s| s.busy),
                        "served": s.as_ref().map(|s| s.served),
                        "word": s.as_ref().map(|s| s.word.clone()),
                    })
                };
                respond_json(req, 200, &body);
            }
            (Method::Get, "/log") => respond_json(req, 200, &serde_json::json!({ "lines": log_tail() })),
            (Method::Post, "/verify") | (Method::Post, "/update") => {
                let mut body = Vec::new();
                if let Err(err) = req.as_reader().take(MAX_BODY as u64).read_to_end(&mut body) {
                    respond_json(req, 400, &serde_json::json!({ "error": format!("read body: {err}") }));
                    continue;
                }
                let work = if path == "/verify" { Work::Verify(req, body) } else { Work::Update(req, body) };
                if let Err(mpsc::SendError(work)) = tx.send(work) {
                    let req = match work {
                        Work::Verify(r, _) | Work::Update(r, _) => r,
                    };
                    respond_json(req, 503, &serde_json::json!({ "error": "worker stopped" }));
                }
            }
            _ => respond_json(req, 404, &serde_json::json!({ "error": "not found" })),
        }
    }
}

fn verify(runtime: &mut Runtime, body: &[u8], exe_sha: &str) -> Result<serde_json::Value, String> {
    let spec: VerifyRequest = serde_json::from_slice(body).map_err(|e| format!("bad request: {e}"))?;
    if spec.files.is_empty() {
        return Err("no files".into());
    }
    let started_at = now_ms();
    let work = std::env::temp_dir().join("lyset-word-verify").join(format!("{started_at}"));
    std::fs::create_dir_all(&work).map_err(|e| format!("create {}: {e}", work.display()))?;
    let mut locals = Vec::new();
    for (i, file) in spec.files.iter().enumerate() {
        let bytes = B64.decode(file.data.as_bytes()).map_err(|e| format!("{}: bad base64: {e}", file.name))?;
        let local = work.join(format!("{i:03}-{}", safe_name(&file.name)));
        std::fs::write(&local, bytes).map_err(|e| format!("write {}: {e}", local.display()))?;
        locals.push(local);
    }
    let opts = Options {
        timeout: Duration::from_millis(spec.timeout_ms),
        repair: spec.repair,
        pdf: spec.pdf,
        protected_view: spec.protected_view,
    };
    let results = runtime.run_files(&locals, &work, opts);
    let mut out = Vec::new();
    for (mut r, file) in results.into_iter().zip(&spec.files) {
        r.file = file.name.clone();
        let pdf = take_artifact(&mut r.pdf);
        let repaired = take_artifact(&mut r.repaired_copy);
        let mut value = serde_json::to_value(&r).map_err(|e| e.to_string())?;
        if let Some(obj) = value.as_object_mut() {
            if let Some(bytes) = pdf {
                obj.insert("pdfBase64".into(), B64.encode(bytes).into());
            }
            if let Some(bytes) = repaired {
                obj.insert("repairedBase64".into(), B64.encode(bytes).into());
            }
        }
        out.push(value);
    }
    let _ = std::fs::remove_dir_all(&work);
    Ok(serde_json::json!({
        "tool": format!("word-verify {VERSION}"),
        "exeSha256": exe_sha,
        "word": runtime.info,
        "startedAt": started_at,
        "finishedAt": now_ms(),
        "results": out,
    }))
}

/// Read an artifact written into the work dir and keep only its file name.
fn take_artifact(slot: &mut Option<String>) -> Option<Vec<u8>> {
    let path = slot.clone()?;
    let bytes = std::fs::read(&path).ok()?;
    *slot = Path::new(&path).file_name().map(|n| n.to_string_lossy().to_string());
    Some(bytes)
}

fn safe_name(name: &str) -> String {
    let base = Path::new(name).file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default();
    let cleaned: String = base.chars().map(|c| if c.is_ascii_alphanumeric() || "._-".contains(c) { c } else { '_' }).collect();
    if cleaned.is_empty() { "document.docx".into() } else { cleaned }
}

fn set_busy(status: &Mutex<Status>, busy: bool) {
    if let Ok(mut s) = status.lock() {
        s.busy = busy;
    }
}

fn respond_json(req: Request, code: u16, body: &serde_json::Value) {
    let bytes = serde_json::to_vec(body).unwrap_or_default();
    let header = Header::from_bytes(&b"Content-Type"[..], &b"application/json"[..]).expect("static header");
    let _ = req.respond(Response::from_data(bytes).with_status_code(code).with_header(header));
}

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}
