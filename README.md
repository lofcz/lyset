# lyset

CLI around the sibling [`rdocx`](https://github.com/tensorbee/rdocx) crate
(cloned to `GitHub/rdocx`, path dependency). It renders the ScioBot
print-document IR to DOCX and, through rdocx's own layout engine, to PDF and
per-page PNGs — one renderer, one look, for Word and PDF export alike.

Priprava ships the built binary as `rdocx-sidecar` under gitignored
`src/Tools/Rdocx/` (`rdocx-sidecar.exe` on Windows, `linux/rdocx-sidecar` on
Linux) and resolves it via `RdocxService` / `ToolBinary`.

## Pipeline

```
material (lesson / worksheet / test / activity / print)
  └─ sciobot-next  src/lib/print-ir/*.ts      → PrintDocument JSON (zod schema in _shared/print-ir)
       └─ rdocx-sidecar render                → DOCX (+ PDF, + PNG pages)
            └─ rdocx (crates/rdocx, rdocx-layout, oxml-layout)
```

* `src/ir.rs` is the serde mirror of `sciobot-next/supabase/functions/_shared/print-ir/schema.ts`.
  Field names, enum values and defaults are 1:1 and unknown fields are
  rejected, so schema drift fails loudly at the boundary.
* `src/render.rs` owns typography, spacing, colour and pagination hints
  (keep-with-next chains for short tasks, repeated table header rows, ruled
  answer lines, fields, panels, two-column options…).
* `src/theme.rs` holds the print theme: A4, margins, Carlito/Caladea (metric
  clones of Calibri/Cambria, bundled — `system-fonts` is off so output is
  identical on every machine), sizes, accent, `atLeast` line pitch.
* Formulas arrive as presentation MathML (`Inline::Math.mathml`), rendered
  from the authored LaTeX by KaTeX on the client (`src/lib/print-ir/mathml.ts`
  in sciobot-next); `rdocx::equation_from_mathml` turns that into OMML and the
  layout engine handles math italic, operator spacing and font fallback. IR
  without `mathml` (hand-written fixtures) goes through `rdocx`'s LaTeX
  subset with its diagnostics surfaced as warnings.

## CLI

```
rdocx-sidecar render <ir.json> --docx <out.docx> [--pdf <out.pdf>] [--png-dir <dir>] [--dpi <n>]
rdocx-sidecar convert <source.docx> --to <pdf|html|md> -o <output>     # reserved, exit 2
```

stdout is one JSON line: `{ "pages": n | null, "warnings": [...] }` (`pages`
is set when PNGs were requested). Exit 0 = success, 1 = render failure
(message on stderr), 2 = usage / not implemented.

## Fidelity loop

`sciobot-next/scripts/print-export-fidelity.ts` pulls every preparation with
ready linked materials from the local Supabase DB, lowers lesson / worksheet /
test / activity / print (student + teacher, every A/B/C variant) to IR, runs
this sidecar and writes `DOCX`, `PDF`, `page-NN.png` and a `sheet.png`
contact sheet per document plus an `index.md` summary under
`/tmp/export-fidelity/out/<entity8>/`. The IR JSON sits next to each output
(`<doc>.ir.json`), so renderer-only changes can be re-checked by pointing
`rdocx-sidecar render` at it without touching the database.

```bash
cargo build                                     # target/debug/rdocx-sidecar
cd ../sciobot-next
bun scripts/print-export-fidelity.ts            # everything
bun scripts/print-export-fidelity.ts --entity 02d8d438 --only test
```

Compare against Word/LibreOffice when in doubt: `soffice --headless
--convert-to pdf out.docx` renders the same DOCX with a different engine;
page counts and breaks should agree.

## Tests

```bash
cargo test            # end-to-end render smoke tests
```

The TypeScript side has `src/lib/print-ir/*.test.ts` (`bunx rstest run
src/lib/print-ir`) covering text/HTML → inline lowering, worksheet/test and
lesson conversion.

## Build (into priprava)

```bash
# Linux → priprava/src/Tools/Rdocx/linux/rdocx-sidecar
priprava/tools/fetch-linux-sidecars.sh rdocx

# Windows-from-Linux → priprava/src/Tools/Rdocx/rdocx-sidecar.exe
priprava/tools/cross-build-win-sidecars.sh rdocx
```
