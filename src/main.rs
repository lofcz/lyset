//! lyset (`rdocx-sidecar`): CLI around the sibling `rdocx` crate.
//!
//! Renders the ScioBot print-document IR (see `sciobot-next/src/lib/print-ir`)
//! to DOCX and, through rdocx's layout engine, to PDF and per-page PNGs.
//!
//! ```text
//! rdocx-sidecar render <ir.json> --docx <out.docx> [--pdf <out.pdf>] [--png-dir <dir>] [--dpi <n>]
//! rdocx-sidecar convert <source.docx> --to <pdf|html|md> -o <output>   (reserved, exit 2)
//! ```
//!
//! stdout: one JSON object `{ "pages": n|null, "warnings": [...] }`.
//! Exit 0 = success, 1 = render failure, 2 = usage / not implemented.

mod ir;
mod render;
mod theme;

use std::path::{Path, PathBuf};
use std::process::ExitCode;

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 || args.iter().any(|a| a == "--help" || a == "-h") {
        eprint_usage(&args[0]);
        return ExitCode::from(2);
    }
    match args[1].as_str() {
        "render" => match run_render(&args[2..]) {
            Ok(code) => code,
            Err(message) => {
                eprintln!("error: {message}");
                ExitCode::from(1)
            }
        },
        "convert" => {
            eprintln!("error: rdocx-sidecar convert is not implemented");
            ExitCode::from(2)
        }
        _ => {
            eprint_usage(&args[0]);
            ExitCode::from(2)
        }
    }
}

struct RenderArgs {
    input: PathBuf,
    docx: Option<PathBuf>,
    pdf: Option<PathBuf>,
    png_dir: Option<PathBuf>,
    dpi: f64,
}

fn parse_render_args(args: &[String]) -> Result<RenderArgs, String> {
    let mut input = None;
    let mut docx = None;
    let mut pdf = None;
    let mut png_dir = None;
    let mut dpi = 110.0;
    let mut i = 0;
    while i < args.len() {
        let a = args[i].as_str();
        let value = |i: &mut usize| -> Result<String, String> {
            *i += 1;
            args.get(*i).cloned().ok_or_else(|| format!("missing value for {a}"))
        };
        match a {
            "--docx" => docx = Some(PathBuf::from(value(&mut i)?)),
            "--pdf" => pdf = Some(PathBuf::from(value(&mut i)?)),
            "--png-dir" => png_dir = Some(PathBuf::from(value(&mut i)?)),
            "--dpi" => dpi = value(&mut i)?.parse::<f64>().map_err(|e| format!("bad --dpi: {e}"))?,
            _ if a.starts_with("--") => return Err(format!("unknown option {a}")),
            _ if input.is_none() => input = Some(PathBuf::from(a)),
            _ => return Err(format!("unexpected argument {a}")),
        }
        i += 1;
    }
    let input = input.ok_or("missing <ir.json>")?;
    if docx.is_none() && pdf.is_none() && png_dir.is_none() {
        return Err("nothing to do: pass --docx, --pdf and/or --png-dir".into());
    }
    Ok(RenderArgs { input, docx, pdf, png_dir, dpi })
}

fn run_render(args: &[String]) -> Result<ExitCode, String> {
    let args = match parse_render_args(args) {
        Ok(a) => a,
        Err(e) => {
            eprintln!("error: {e}");
            eprint_usage("rdocx-sidecar");
            return Ok(ExitCode::from(2));
        }
    };
    let json = std::fs::read_to_string(&args.input).map_err(|e| format!("read {}: {e}", args.input.display()))?;
    let ir: ir::PrintDocument = serde_json::from_str(&json).map_err(|e| format!("invalid IR: {e}"))?;

    let (mut doc, report) = render::render(&ir)?;
    let mut warnings = report.warnings;

    if let Some(path) = &args.docx {
        ensure_parent(path)?;
        doc.save(path).map_err(|e| format!("save docx: {e}"))?;
    }
    if let Some(path) = &args.pdf {
        ensure_parent(path)?;
        doc.save_pdf(path).map_err(|e| format!("render pdf: {e}"))?;
    }
    let mut pages: Option<usize> = None;
    if let Some(dir) = &args.png_dir {
        std::fs::create_dir_all(dir).map_err(|e| format!("mkdir {}: {e}", dir.display()))?;
        let rendered = doc.render_all_pages(args.dpi).map_err(|e| format!("render png: {e}"))?;
        for (idx, bytes) in rendered.iter().enumerate() {
            let path = dir.join(format!("page-{:02}.png", idx + 1));
            std::fs::write(&path, bytes).map_err(|e| format!("write {}: {e}", path.display()))?;
        }
        pages = Some(rendered.len());
    }

    warnings.sort();
    warnings.dedup();
    println!(
        "{}",
        serde_json::json!({
            "pages": pages,
            "warnings": warnings,
        })
    );
    Ok(ExitCode::SUCCESS)
}

fn ensure_parent(path: &Path) -> Result<(), String> {
    if let Some(parent) = path.parent() {
        if !parent.as_os_str().is_empty() {
            std::fs::create_dir_all(parent).map_err(|e| format!("mkdir {}: {e}", parent.display()))?;
        }
    }
    Ok(())
}

fn eprint_usage(argv0: &str) {
    eprintln!("usage: {argv0} render <ir.json> --docx <out.docx> [--pdf <out.pdf>] [--png-dir <dir>] [--dpi <n>]");
    eprintln!("       {argv0} convert <source.docx> --to <pdf|html|md> -o <output>");
}
