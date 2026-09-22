# lyset

CLI that renders a print-document IR (JSON) to DOCX, PDF, and page PNGs
through [`rdocx`](https://github.com/lofcz/rdocx). Word and PDF share one
layout engine, so they look the same.

The IR is defined in `src/ir.rs`. Unknown fields are rejected. Formulas
may carry presentation MathML; otherwise the authored LaTeX is parsed by
rdocx's subset.

`rdocx` comes from the exact Git revision pinned in `Cargo.toml`, including
editable boxed math and correct PDF text extraction after ligatures. System
fonts are disabled for deterministic output. No sibling checkout is required.

Unsupported or malformed formulas are preserved as literal text and reported
in `warnings`. They do not abort the rest of the export.

```
lyset render <ir.json> --docx <out.docx> [--pdf <out.pdf>] [--png-dir <dir>] [--dpi <n>]
```

stdout is one JSON line: `{ "pages": n | null, "warnings": [...] }`.
Exit 0 success, 1 render failure, 2 usage.

```bash
cargo build --release --locked # target/release/lyset
cargo test --locked
# Windows MSVC cross-build from Linux (requires cargo-xwin):
cargo xwin build --release --locked --target x86_64-pc-windows-msvc
```
