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

Noto Emoji is embedded in generated documents for deterministic monochrome
emoji fallback, including flags and joined sequences. The outline font works
in both PDF and DOCX without installing fonts on the server. Its SIL Open Font
License is included in `fonts/OFL-NotoEmoji.txt`. `fonts/NotoEmoji-Regular.ttf`
is the weight-400 static instance of Google Fonts' `ofl/notoemoji/NotoEmoji[wght].ttf`,
generated with fontTools `instantiateVariableFont(font, {"wght": 400})`.

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

Code blocks use a `paragraph` with `style: "code"` and literal text/break
inlines. They render as shaded, regular monospace paragraphs. Four-column
tabs, indentation and empty lines are preserved. Long blocks may span pages.
Older task prompts containing code-only lines are split into code blocks
without treating inline identifiers inside prose as blocks.

GitHub Actions builds `lyset-linux-x86_64` and `lyset-windows-x86_64`
artifacts on pushes. The rdocx revision is pinned in both Cargo files.

Highlighted code tokens can set `color` to a six-digit hexadecimal foreground (with an optional `#`) and use `bold`, `italic`, `underline`, or `strike`. These runs retain their formatting through DOCX and PDF export. Producers should supply Shiki light-theme token colours for the light code panel. Plain code remains regular monospace.
