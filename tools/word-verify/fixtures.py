#!/usr/bin/env python3
"""Known-bad DOCX fixtures, each checked against real Word.

Takes a DOCX that Word opens cleanly, applies one mutation per fixture, runs
every result through Word (word-verify service) and docxlint, and writes a
manifest recording both verdicts. A fixture where the linter and Word
disagree is either a missing lint rule or a rule whose severity is wrong.

    fixtures.py BASE.docx [--out DIR]

The manifest is the evidence behind docxlint's severities: re-run it after a
Word update to see which behaviours changed.
"""
from __future__ import annotations

import argparse
import io
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Callable

sys.path.insert(0, str(Path(__file__).resolve().parent))
import docxlint  # noqa: E402
import wordcheck  # noqa: E402
from docxbisect import Pkg, all_elements, child_spans, inner, put, read_pkg, splice, text, write_pkg  # noqa: E402

CT = "[Content_Types].xml"
DOC = "word/document.xml"
RELS = "word/_rels/document.xml.rels"


def sub_first(pkg: Pkg, part: str, pattern: str, repl: str, flags: int = 0) -> None:
    xml = text(pkg, part)
    new, n = re.subn(pattern, repl, xml, count=1, flags=flags)
    if n == 0:
        raise ValueError(f"{part}: pattern {pattern!r} not found")
    put(pkg, part, new)


def first_span(pkg: Pkg, part: str, qname: str, nth: int = 0) -> tuple[str, tuple[int, int, str]]:
    xml = text(pkg, part)
    spans = all_elements(xml, {qname})
    if len(spans) <= nth:
        raise ValueError(f"{part}: no {qname}")
    return xml, spans[nth]


def m_ct_duplicate_override(p: Pkg) -> None:
    sub_first(p, CT, r'(<Override\b[^>]*PartName="/word/document.xml"[^>]*/>)', r"\1\1")


def m_ct_duplicate_default(p: Pkg) -> None:
    sub_first(p, CT, r'(<Default\b[^>]*Extension="xml"[^>]*/>)', r"\1\1")


def m_ct_untyped_font(p: Pkg) -> None:
    sub_first(p, CT, r'<Override\b[^>]*PartName="/word/fonts/[^"]*"[^>]*/>', "")


def m_rel_duplicate_id(p: Pkg) -> None:
    sub_first(p, RELS, r'Id="rdocxTheme"', 'Id="rId0"')


def m_rel_missing_target(p: Pkg) -> None:
    sub_first(p, RELS, r'Target="footer1.xml"', 'Target="footer9.xml"')


def m_rel_dangling_rid(p: Pkg) -> None:
    sub_first(p, DOC, r'(<w:footerReference\b[^>]*r:id=")[^"]*"', r'\1rId99"')


def m_xml_control_char(p: Pkg) -> None:
    sub_first(p, DOC, r"(<w:t(?:\s[^>]*)?>)([^<]{3})", "\\1\\2\x01")


def m_xml_control_char_ref(p: Pkg) -> None:
    sub_first(p, DOC, r"(<w:t(?:\s[^>]*)?>)([^<]{3})", r"\1\2&#1;")


def m_xml_vertical_tab(p: Pkg) -> None:
    sub_first(p, DOC, r"(<w:t(?:\s[^>]*)?>)([^<]{3})", "\\1\\2\x0b")


def m_xml_malformed(p: Pkg) -> None:
    sub_first(p, DOC, r"</w:p>", "</w:pp>")


def m_cell_without_paragraph(p: Pkg) -> None:
    xml, tc = first_span(p, DOC, "w:tc")
    rng = inner(xml, tc)
    kids = child_spans(xml, rng[0], rng[1])
    drop = [(s, e) for s, e, q in kids if q != "w:tcPr"]
    put(p, DOC, splice(xml, drop))


def m_cell_ends_with_table(p: Pkg) -> None:
    xml, tc = first_span(p, DOC, "w:tc")
    rng = inner(xml, tc)
    kids = child_spans(xml, rng[0], rng[1])
    last = [k for k in kids if k[2] == "w:p"][-1]
    nested = '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
    put(p, DOC, xml[:last[1]] + nested + xml[last[1]:])


def m_table_without_rows(p: Pkg) -> None:
    xml, tbl = first_span(p, DOC, "w:tbl")
    rng = inner(xml, tbl)
    put(p, DOC, splice(xml, [(s, e) for s, e, q in child_spans(xml, rng[0], rng[1]) if q == "w:tr"]))


def m_row_without_cells(p: Pkg) -> None:
    xml, tr = first_span(p, DOC, "w:tr")
    rng = inner(xml, tr)
    put(p, DOC, splice(xml, [(s, e) for s, e, q in child_spans(xml, rng[0], rng[1]) if q == "w:tc"]))


def m_sectpr_not_last(p: Pkg) -> None:
    xml = text(p, DOC)
    body = all_elements(xml, {"w:body"})[0]
    rng = inner(xml, body)
    kids = child_spans(xml, rng[0], rng[1])
    sect = kids[-1]
    assert sect[2] == "w:sectPr"
    last_block = kids[-2]
    moved = xml[:last_block[0]] + xml[sect[0]:sect[1]] + xml[last_block[0]:last_block[1]] + xml[sect[1]:]
    put(p, DOC, moved)


def m_font_key_wrong(p: Pkg) -> None:
    sub_first(p, "word/fontTable.xml", r'w:fontKey="\{[0-9A-F]', 'w:fontKey="{0')


def m_font_key_unbraced(p: Pkg) -> None:
    sub_first(p, "word/fontTable.xml", r'w:fontKey="\{([^}]*)\}"', r'w:fontKey="\1"')


def m_mc_ignorable_undeclared(p: Pkg) -> None:
    sub_first(p, DOC, r"<w:document\b", '<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14x"')


def m_nan_grid_col(p: Pkg) -> None:
    sub_first(p, DOC, r'(<w:gridCol w:w=")\d+"', r'\1NaN"')


def m_nan_font_size(p: Pkg) -> None:
    sub_first(p, DOC, r'(<w:sz w:val=")\d+"', r'\1NaN"')


def m_decimal_twips(p: Pkg) -> None:
    sub_first(p, DOC, r'(<w:gridCol w:w=")(\d+)"', r'\g<1>\2.5"')


def m_decimal_spacing(p: Pkg) -> None:
    sub_first(p, DOC, r'(<w:spacing w:before=")(\d+)"', r'\g<1>\2.5"')


def m_negative_font_size(p: Pkg) -> None:
    sub_first(p, DOC, r'(<w:sz w:val=")\d+"', r'\1-4"')


def m_unknown_w_element(p: Pkg) -> None:
    sub_first(p, DOC, r"(<w:p>)", r"\1<w:bogus/>")


def m_rpr_wrong_order(p: Pkg) -> None:
    sub_first(p, DOC, r"<w:rPr>(\s*)(<w:rFonts[^>]*/>)(.*?)(<w:sz [^>]*/>)", r"<w:rPr>\1\4\3\2", re.S)


def m_empty_run_text(p: Pkg) -> None:
    sub_first(p, DOC, r"(<w:r>)", r"\1<w:t></w:t>")


def m_text_outside_run(p: Pkg) -> None:
    sub_first(p, DOC, r"(<w:p>)", r"\1<w:t>loose</w:t>")


def m_absolute_targets_everywhere(p: Pkg) -> None:
    put(p, RELS, re.sub(r'Target="(?!/)([^"]*)"', r'Target="/word/\1"', text(p, RELS)))


def m_foreign_attribute_unignorable(p: Pkg) -> None:
    sub_first(p, DOC, r"<w:document\b", '<w:document xmlns:x="urn:example:x"')
    sub_first(p, DOC, r"<w:p>", '<w:p x:note="1">')


def m_foreign_element_unignorable(p: Pkg) -> None:
    sub_first(p, DOC, r"<w:document\b", '<w:document xmlns:x="urn:example:x"')
    sub_first(p, DOC, r"(<w:p>)", r"\1<x:note/>")


def zip_bytes(entries: list[tuple[str, bytes]]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries:
            z.writestr(zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0)), data, compress_type=zipfile.ZIP_DEFLATED)
    return buf.getvalue()


def ordered(p: Pkg) -> list[tuple[str, bytes]]:
    return sorted(p.items(), key=lambda kv: (kv[0] != CT, kv[0]))


MUTATIONS: list[tuple[str, Callable[[Pkg], None] | None]] = [
    ("control", lambda p: None),
    ("ct-duplicate-override", m_ct_duplicate_override),
    ("ct-duplicate-default", m_ct_duplicate_default),
    ("ct-untyped-font-part", m_ct_untyped_font),
    ("rel-duplicate-id", m_rel_duplicate_id),
    ("rel-missing-target", m_rel_missing_target),
    ("rel-dangling-rid", m_rel_dangling_rid),
    ("xml-control-char", m_xml_control_char),
    ("xml-control-char-reference", m_xml_control_char_ref),
    ("xml-vertical-tab", m_xml_vertical_tab),
    ("xml-malformed", m_xml_malformed),
    ("wml-cell-without-paragraph", m_cell_without_paragraph),
    ("wml-cell-ends-with-table", m_cell_ends_with_table),
    ("wml-table-without-rows", m_table_without_rows),
    ("wml-row-without-cells", m_row_without_cells),
    ("wml-sectpr-not-last", m_sectpr_not_last),
    ("font-key-wrong", m_font_key_wrong),
    ("font-key-unbraced", m_font_key_unbraced),
    ("mc-ignorable-undeclared", m_mc_ignorable_undeclared),
    ("nan-grid-col", m_nan_grid_col),
    ("nan-font-size", m_nan_font_size),
    ("decimal-twips-grid-col", m_decimal_twips),
    ("decimal-twips-spacing", m_decimal_spacing),
    ("negative-font-size", m_negative_font_size),
    ("unknown-w-element", m_unknown_w_element),
    ("rpr-wrong-child-order", m_rpr_wrong_order),
    ("empty-run-text", m_empty_run_text),
    ("text-outside-run", m_text_outside_run),
    ("absolute-targets-everywhere", m_absolute_targets_everywhere),
    ("foreign-attribute-not-ignorable", m_foreign_attribute_unignorable),
    ("foreign-element-not-ignorable", m_foreign_element_unignorable),
    ("zip-duplicate-entry", None),
    ("zip-directory-entries", None),
    ("zip-stored-no-compression", None),
    ("zip-content-types-last", None),
]


def build(name: str, fn: Callable[[Pkg], None] | None, base: Pkg) -> bytes:
    p = dict(base)
    if fn is not None:
        fn(p)
        return write_pkg(p)
    entries = ordered(p)
    if name == "zip-duplicate-entry":
        import warnings
        warnings.simplefilter("ignore")
        return zip_bytes(entries + [(DOC, p[DOC])])
    if name == "zip-directory-entries":
        return zip_bytes([("word/", b""), ("_rels/", b"")] + entries)
    if name == "zip-stored-no-compression":
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as z:
            for n, d in entries:
                z.writestr(n, d)
        return buf.getvalue()
    if name == "zip-content-types-last":
        return zip_bytes([e for e in entries if e[0] != CT] + [(CT, p[CT])])
    raise ValueError(name)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("base")
    ap.add_argument("--out", default=str(Path(__file__).resolve().parent / "fixtures"))
    args = ap.parse_args()
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    base = read_pkg(Path(args.base).read_bytes())
    built: list[tuple[str, bytes]] = []
    for name, fn in MUTATIONS:
        try:
            built.append((name, build(name, fn, base)))
        except Exception as err:
            print(f"skip {name}: {err}")
    for name, data in built:
        (out / f"{name}.docx").write_bytes(data)
    results = wordcheck.verify([(f"{n}.docx", d) for n, d in built], repair=True, timeout_ms=30_000)["results"]
    manifest = []
    print(f"{'fixture':34} {'word':8} {'lint':6} {'agree':8}  word error / lint rules")
    for (name, data), r in zip(built, results):
        findings = docxlint.lint_bytes(data)
        errors = sorted({f.rule for f in findings if f.severity == "error"})
        word_bad = r["verdict"] != "ok"
        err = (r.get("strict") or {}).get("error") or (r.get("protectedView") or {}).get("error") or {}
        # "missed": Word fails and lint is silent (a lint gap, always wrong).
        # "stricter": lint errors on something this Word build tolerates (fine
        # for values that are generator bugs anyway, e.g. NaN).
        agree = "yes" if word_bad == bool(errors) else ("missed" if word_bad else "stricter")
        manifest.append({
            "fixture": name,
            "word": r["verdict"],
            "wordError": err.get("wordError"),
            "wordMessage": (err.get("description") or "").splitlines()[0] if err else None,
            "protectedView": (r.get("protectedView") or {}).get("opened"),
            "dialogs": [d["text"][:200] for d in r.get("dialogs", []) if d["kind"] != "other"],
            "lintErrors": errors,
            "lintWarnings": sorted({f.rule for f in findings if f.severity == "warning"}),
            "agree": agree,
        })
        detail = f"{err.get('wordError') or ''} {((err.get('description') or '').splitlines() or [''])[0][:50]}" if err else ""
        print(f"{name:34} {r['verdict']:8} {'ERR' if errors else 'ok':6} {agree:8}  {detail} {','.join(errors)}")
    word = results[0].get("word") if results else None
    (out / "manifest.json").write_text(json.dumps({"base": Path(args.base).name, "results": manifest}, indent=2, ensure_ascii=False))
    missed = [m["fixture"] for m in manifest if m["agree"] == "missed"]
    stricter = [m["fixture"] for m in manifest if m["agree"] == "stricter"]
    print(f"\n{len(manifest)} fixtures; lint missed {len(missed)} {missed}; lint stricter than Word on {len(stricter)} {stricter}")
    print(f"manifest in {out / 'manifest.json'}")
    return 1 if missed else 0


if __name__ == "__main__":
    sys.exit(main())
