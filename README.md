# lyset

CLI that renders a print-document IR (JSON) to DOCX, PDF, and page PNGs
through [`rdocx`](https://github.com/lofcz/rdocx). Word and PDF share one
layout engine, so they look the same.

The IR is defined in `src/ir.rs`. Unknown fields are rejected. Formulas
may carry presentation MathML; otherwise the authored LaTeX is parsed by
rdocx's subset.

`rdocx` comes from the git dependency (system fonts off, so output is
deterministic — bundled Carlito / Caladea / Liberation).

```
lyset render <ir.json> --docx <out.docx> [--pdf <out.pdf>] [--png-dir <dir>] [--dpi <n>]
```

stdout is one JSON line: `{ "pages": n | null, "warnings": [...] }`.
Exit 0 success, 1 render failure, 2 usage.

```bash
cargo build          # target/debug/lyset
cargo test
```
