#!/usr/bin/env python3
"""Shrink a DOCX that Word rejects to the smallest package that still fails.

    docxbisect.py FILE.docx [--out DIR] [--fail-on any|strict|protected] [--loose]

Needs the word-verify service (see wordcheck.py). Phases:

  1. confirm   The input fails, and so does a plain re-zip of it. If the re-zip
               opens, the ZIP container itself is the trigger and we stop.
  2. ablate    Package-level edits one at a time (drop embedded fonts, strip
               foreign namespaces, relativize targets, drop math, ...). An edit
               that makes Word accept the file names a feature involved.
  3. minimize  Delta debugging (ddmin) over the body blocks, then inside each
               surviving table (rows, cells) and paragraph (runs).

A candidate only counts as failing when Word's error signature matches the
original, so the search cannot drift into a different, self-inflicted error.
XML is edited as text by element spans: untouched bytes stay identical and
namespace prefixes (which mc:Ignorable names) are never rewritten.

Writes <out>/minimal.docx, <out>/fragment.xml and <out>/report.json.
"""
from __future__ import annotations

import argparse
import io
import json
import posixpath
import re
import sys
import time
import zipfile
from pathlib import Path
from typing import Callable, Iterable

sys.path.insert(0, str(Path(__file__).resolve().parent))
import docxlint  # noqa: E402
import wordcheck  # noqa: E402

Pkg = dict[str, bytes]

# ---------------------------------------------------------------------------
# Package I/O
# ---------------------------------------------------------------------------


def read_pkg(data: bytes) -> Pkg:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {i.filename: z.read(i) for i in z.infolist() if not i.is_dir()}


def write_pkg(pkg: Pkg) -> bytes:
    buf = io.BytesIO()
    names = sorted(pkg, key=lambda n: (n != "[Content_Types].xml", n))
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name in names:
            z.writestr(zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0)), pkg[name], compress_type=zipfile.ZIP_DEFLATED)
    return buf.getvalue()


def text(pkg: Pkg, name: str) -> str | None:
    data = pkg.get(name)
    return data.decode("utf-8") if data is not None else None


def put(pkg: Pkg, name: str, xml: str) -> None:
    pkg[name] = xml.encode("utf-8")


def xml_parts(pkg: Pkg) -> list[str]:
    return [n for n in pkg if n.endswith(".xml") or n.endswith(".rels")]


# ---------------------------------------------------------------------------
# Raw XML element spans
# ---------------------------------------------------------------------------

TOKEN = re.compile(r"<!--.*?-->|<\?.*?\?>|<!\[CDATA\[.*?\]\]>|<(/?)([A-Za-z_][\w.\-]*(?::[A-Za-z_][\w.\-]*)?)((?:[^>\"']|\"[^\"]*\"|'[^']*')*?)(/?)>", re.S)


def child_spans(xml: str, start: int, end: int) -> list[tuple[int, int, str]]:
    """Direct child elements of xml[start:end] as (start, end, qname)."""
    out: list[tuple[int, int, str]] = []
    depth = 0
    open_start, open_name = 0, ""
    for m in TOKEN.finditer(xml, start, end):
        if m.group(2) is None:
            continue
        closing, name, selfclose = m.group(1) == "/", m.group(2), m.group(4) == "/"
        if closing:
            depth -= 1
            if depth == 0:
                out.append((open_start, m.end(), open_name))
        elif selfclose:
            if depth == 0:
                out.append((m.start(), m.end(), name))
        else:
            if depth == 0:
                open_start, open_name = m.start(), name
            depth += 1
    return out


def inner(xml: str, span: tuple[int, int, str]) -> tuple[int, int] | None:
    """Inner range of an element span, or None for a self-closing element."""
    s, e, name = span
    head = TOKEN.match(xml, s)
    if head is None or head.group(4) == "/":
        return None
    close = xml.rfind(f"</{name}", s, e)
    return head.end(), close


def find_element(xml: str, qname: str, start: int = 0, end: int | None = None) -> tuple[int, int, str] | None:
    """First element with this qname at any depth."""
    end = len(xml) if end is None else end
    for m in TOKEN.finditer(xml, start, end):
        if m.group(2) == qname and m.group(1) != "/":
            if m.group(4) == "/":
                return m.start(), m.end(), qname
            spans = child_spans(xml, m.start(), end)
            return spans[0] if spans else None
    return None


def all_elements(xml: str, qnames: set[str]) -> list[tuple[int, int, str]]:
    """Outermost elements with any of these qnames (not nested in each other)."""
    out: list[tuple[int, int, str]] = []
    pos = 0
    while True:
        best = None
        for m in TOKEN.finditer(xml, pos):
            if m.group(2) in qnames and m.group(1) != "/":
                best = m
                break
        if best is None:
            return out
        if best.group(4) == "/":
            span = (best.start(), best.end(), best.group(2))
        else:
            span = child_spans(xml, best.start(), len(xml))[0]
        out.append(span)
        pos = span[1]


def splice(xml: str, spans: Iterable[tuple[int, int]], replacement: Callable[[str], str] | str = "") -> str:
    parts, last = [], 0
    for s, e in sorted(spans):
        if s < last:
            continue
        parts.append(xml[last:s])
        parts.append(replacement(xml[s:e]) if callable(replacement) else replacement)
        last = e
    parts.append(xml[last:])
    return "".join(parts)


def remove_elements(xml: str, qnames: set[str]) -> str:
    return splice(xml, [(s, e) for s, e, _ in all_elements(xml, qnames)])


def unwrap_elements(xml: str, qnames: set[str]) -> str:
    def keep_inner(fragment: str) -> str:
        span = child_spans(fragment, 0, len(fragment))[0]
        rng = inner(fragment, span)
        return fragment[rng[0]:rng[1]] if rng else ""
    return splice(xml, [(s, e) for s, e, _ in all_elements(xml, qnames)], keep_inner)


# ---------------------------------------------------------------------------
# Relationship / content-type helpers
# ---------------------------------------------------------------------------

REL = re.compile(r"<Relationship\b[^>]*/>")


def attr(tag: str, name: str) -> str | None:
    m = re.search(rf'\b{name}="([^"]*)"', tag)
    return m.group(1) if m else None


def rels_path(part: str) -> str:
    d, n = posixpath.split(part)
    return posixpath.join(d, "_rels", n + ".rels")


def source_of(rels: str) -> str:
    d, n = posixpath.split(rels)
    return posixpath.join(posixpath.dirname(d), n[: -len(".rels")])


def resolve(source: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def drop_relationships(pkg: Pkg, predicate: Callable[[str, str], bool]) -> None:
    """Remove relationships where predicate(type_suffix, resolved_target) holds."""
    for name in [n for n in pkg if n.endswith(".rels")]:
        xml = text(pkg, name)
        source = source_of(name)
        def keep(m: re.Match[str]) -> str:
            tag = m.group(0)
            ty = (attr(tag, "Type") or "").rsplit("/", 1)[-1]
            target = attr(tag, "Target") or ""
            external = attr(tag, "TargetMode") == "External"
            return "" if predicate(ty, "" if external else resolve(source, target)) else tag
        put(pkg, name, REL.sub(keep, xml))


def drop_parts(pkg: Pkg, names: Iterable[str]) -> None:
    names = set(names)
    for n in names:
        pkg.pop(n, None)
        pkg.pop(rels_path(n), None)
    drop_relationships(pkg, lambda _t, target: target in names)
    ct = text(pkg, "[Content_Types].xml")
    if ct:
        put(pkg, "[Content_Types].xml", re.sub(r'<Override\b[^>]*PartName="/([^"]*)"[^>]*/>', lambda m: "" if m.group(1) in names else m.group(0), ct))


def parts_of_type(pkg: Pkg, type_suffix: str) -> list[str]:
    found = []
    for name in [n for n in pkg if n.endswith(".rels")]:
        source = source_of(name)
        for tag in REL.findall(text(pkg, name)):
            if (attr(tag, "Type") or "").endswith("/" + type_suffix) and attr(tag, "TargetMode") != "External":
                found.append(resolve(source, attr(tag, "Target") or ""))
    return found


STORY_PARTS = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$")


def edit_stories(pkg: Pkg, fn: Callable[[str], str]) -> None:
    for name in [n for n in pkg if STORY_PARTS.match(n)]:
        put(pkg, name, fn(text(pkg, name)))


# ---------------------------------------------------------------------------
# Ablations
# ---------------------------------------------------------------------------

KNOWN_NS = (
    "http://schemas.openxmlformats.org/",
    "http://purl.org/",
    "http://www.w3.org/",
    "http://schemas.microsoft.com/office/",
    "urn:schemas-microsoft-com:",
)


def ab_drop_embedded_fonts(pkg: Pkg) -> None:
    ft = text(pkg, "word/fontTable.xml")
    if ft:
        put(pkg, "word/fontTable.xml", remove_elements(ft, {"w:embedRegular", "w:embedBold", "w:embedItalic", "w:embedBoldItalic"}))
    drop_parts(pkg, [n for n in pkg if n.startswith("word/fonts/")])


def ab_strip_foreign_markup(pkg: Pkg) -> None:
    for name in xml_parts(pkg):
        xml = text(pkg, name)
        foreign = {p for p, uri in re.findall(r'xmlns:([\w.\-]+)="([^"]*)"', xml) if not uri.startswith(KNOWN_NS)}
        if not foreign:
            continue
        xml = remove_elements(xml, {q for q in set(re.findall(r"<([\w.\-]+:[\w.\-]+)", xml)) if q.split(":")[0] in foreign})
        for p in foreign:
            xml = re.sub(rf'\s{re.escape(p)}:[\w.\-]+="[^"]*"', "", xml)
            xml = re.sub(rf'\sxmlns:{re.escape(p)}="[^"]*"', "", xml)
        def fix_ignorable(m: re.Match[str]) -> str:
            kept = [p for p in m.group(2).split() if p not in foreign]
            return f' {m.group(1)}:Ignorable="{" ".join(kept)}"' if kept else ""
        xml = re.sub(r'\s([\w.\-]+):Ignorable="([^"]*)"', fix_ignorable, xml)
        put(pkg, name, xml)


def ab_relativize_targets(pkg: Pkg) -> None:
    for name in [n for n in pkg if n.endswith(".rels")]:
        source = source_of(name)
        def rel(m: re.Match[str]) -> str:
            tag = m.group(0)
            target = attr(tag, "Target") or ""
            if attr(tag, "TargetMode") == "External" or not target.startswith("/"):
                return tag
            relative = posixpath.relpath(target.lstrip("/"), posixpath.dirname(source) or ".")
            return tag.replace(f'Target="{target}"', f'Target="{relative}"')
        put(pkg, name, REL.sub(rel, text(pkg, name)))


def drop_type(type_suffix: str) -> Callable[[Pkg], None]:
    return lambda pkg: drop_parts(pkg, parts_of_type(pkg, type_suffix))


def ab_drop_docprops(pkg: Pkg) -> None:
    drop_parts(pkg, [n for n in pkg if n.startswith("docProps/")])


def ab_drop_headers_footers(pkg: Pkg) -> None:
    drop_parts(pkg, parts_of_type(pkg, "header") + parts_of_type(pkg, "footer"))
    doc = text(pkg, "word/document.xml")
    put(pkg, "word/document.xml", remove_elements(doc, {"w:headerReference", "w:footerReference"}))


def ab_drop_drawings(pkg: Pkg) -> None:
    edit_stories(pkg, lambda x: remove_elements(x, {"w:drawing", "w:pict", "w:object", "mc:AlternateContent"}))
    drop_parts(pkg, parts_of_type(pkg, "image"))


def ab_drop_math(pkg: Pkg) -> None:
    edit_stories(pkg, lambda x: splice(x, [(s, e) for s, e, _ in all_elements(x, {"m:oMathPara", "m:oMath"})], "<w:r><w:t>math</w:t></w:r>"))


def ab_drop_fields(pkg: Pkg) -> None:
    def fn(x: str) -> str:
        x = unwrap_elements(x, {"w:fldSimple"})
        runs = [(s, e) for s, e, _ in all_elements(x, {"w:r"}) if re.search(r"<w:(fldChar|instrText)\b", x[s:e])]
        return splice(x, runs)
    edit_stories(pkg, fn)


def ab_unwrap_hyperlinks(pkg: Pkg) -> None:
    edit_stories(pkg, lambda x: unwrap_elements(x, {"w:hyperlink"}))


def ab_drop_bookmarks(pkg: Pkg) -> None:
    edit_stories(pkg, lambda x: remove_elements(x, {"w:bookmarkStart", "w:bookmarkEnd"}))


def ab_drop_run_formatting(pkg: Pkg) -> None:
    edit_stories(pkg, lambda x: remove_elements(x, {"w:rPr"}))


def ab_drop_paragraph_formatting(pkg: Pkg) -> None:
    edit_stories(pkg, lambda x: remove_elements(x, {"w:pPr"}))


def ab_drop_empty_text(pkg: Pkg) -> None:
    edit_stories(pkg, lambda x: re.sub(r"<w:t(?:\s[^>]*)?(?:/>|></w:t>)", "", x))


def ab_minimal_styles(pkg: Pkg) -> None:
    if "word/styles.xml" in pkg:
        put(pkg, "word/styles.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')


def ab_normalize_package(pkg: Pkg) -> None:
    """Dedupe content types and relationship ids, drop relationships to missing parts."""
    ct = text(pkg, "[Content_Types].xml")
    if ct:
        seen: set[str] = set()
        def once(m: re.Match[str]) -> str:
            key = (m.group(1) + ":" + (attr(m.group(0), "PartName") or attr(m.group(0), "Extension") or "")).lower()
            if key in seen:
                return ""
            seen.add(key)
            return m.group(0)
        put(pkg, "[Content_Types].xml", re.sub(r"<(Override|Default)\b[^>]*/>", once, ct))
    names = {n.lower() for n in pkg}
    for name in [n for n in pkg if n.endswith(".rels")]:
        source = source_of(name)
        ids: set[str] = set()
        def keep(m: re.Match[str]) -> str:
            tag = m.group(0)
            rid = attr(tag, "Id") or ""
            external = attr(tag, "TargetMode") == "External"
            missing = not external and resolve(source, attr(tag, "Target") or "").lower() not in names
            if rid in ids or missing:
                return ""
            ids.add(rid)
            return tag
        put(pkg, name, REL.sub(keep, text(pkg, name)))


ABLATIONS: list[tuple[str, Callable[[Pkg], None]]] = [
    ("normalize-package", ab_normalize_package),
    ("drop-embedded-fonts", ab_drop_embedded_fonts),
    ("strip-foreign-markup", ab_strip_foreign_markup),
    ("relativize-targets", ab_relativize_targets),
    ("drop-docprops", ab_drop_docprops),
    ("drop-custom-properties", drop_type("custom-properties")),
    ("drop-theme", drop_type("theme")),
    ("drop-settings", drop_type("settings")),
    ("drop-font-table", drop_type("fontTable")),
    ("minimal-styles", ab_minimal_styles),
    ("drop-numbering", drop_type("numbering")),
    ("drop-headers-footers", ab_drop_headers_footers),
    ("drop-drawings", ab_drop_drawings),
    ("drop-math", ab_drop_math),
    ("drop-fields", ab_drop_fields),
    ("unwrap-hyperlinks", ab_unwrap_hyperlinks),
    ("drop-bookmarks", ab_drop_bookmarks),
    ("drop-empty-text", ab_drop_empty_text),
    ("drop-run-formatting", ab_drop_run_formatting),
    ("drop-paragraph-formatting", ab_drop_paragraph_formatting),
]

# ---------------------------------------------------------------------------
# Oracle
# ---------------------------------------------------------------------------


def signature(r: dict) -> tuple:
    def err(p: dict | None):
        e = (p or {}).get("error") or {}
        return e.get("wordError") or e.get("scode") or e.get("hresult")
    pv = r.get("protectedView") or {}
    return (r["verdict"], err(r.get("strict")), err(pv), pv.get("edited"))


class Oracle:
    def __init__(self, fail_on: str, loose: bool, batch: int, timeout_ms: int) -> None:
        self.fail_on, self.loose, self.batch, self.timeout_ms = fail_on, loose, batch, timeout_ms
        self.target: tuple | None = None
        # Lint error rules present in the input. A candidate that adds a new one
        # has broken the package in some other way (e.g. a row left without
        # cells, which Word also rejects with 5121) and must not count.
        self.allowed_lint: set[str] | None = None
        self.drift_rejected = 0
        self.calls = 0
        self.files = 0

    def failing(self, r: dict) -> bool:
        strict = r.get("strict") or {}
        pv = r.get("protectedView") or {}
        strict_bad = not strict.get("opened") or any(d["kind"] == "problem" for d in r.get("dialogs", []))
        pv_bad = bool(pv) and (not pv.get("opened") or pv.get("edited") is False)
        bad = {"any": r["verdict"] != "ok", "strict": strict_bad, "protected": pv_bad}[self.fail_on]
        return bad and (self.loose or self.target is None or signature(r) == self.target)

    def run(self, pkgs: list[bytes]) -> list[dict]:
        out: list[dict] = []
        for i in range(0, len(pkgs), self.batch):
            chunk = pkgs[i:i + self.batch]
            self.calls += 1
            self.files += len(chunk)
            res = wordcheck.verify([(f"c{i + j:04d}.docx", b) for j, b in enumerate(chunk)], repair=False, timeout_ms=self.timeout_ms)
            out.extend(res["results"])
        return out

    def drifts(self, pkg: bytes) -> bool:
        if self.allowed_lint is None or self.loose:
            return False
        rules = {f.rule for f in docxlint.lint_bytes(pkg) if f.severity == "error"}
        return not rules <= self.allowed_lint

    def test(self, pkgs: list[bytes]) -> list[bool]:
        verdicts: list[bool | None] = [False if self.drifts(p) else None for p in pkgs]
        self.drift_rejected += sum(1 for v in verdicts if v is False)
        todo = [p for p, v in zip(pkgs, verdicts) if v is None]
        results = iter(self.failing(r) for r in self.run(todo)) if todo else iter(())
        return [next(results) if v is None else v for v in verdicts]


# ---------------------------------------------------------------------------
# ddmin over child spans of a container element
# ---------------------------------------------------------------------------


def ddmin(items: list[int], test_batch: Callable[[list[list[int]]], list[bool]], log: Callable[[str], None]) -> list[int]:
    if not items:
        return items
    if test_batch([[]])[0]:
        return []
    n = 2
    while len(items) >= 2:
        size = len(items)
        chunks = [items[k * size // n:(k + 1) * size // n] for k in range(n)]
        chunks = [c for c in chunks if c]
        comps = [[x for x in items if x not in set(c)] for c in chunks]
        res = test_batch(chunks + comps)
        rc, rp = res[:len(chunks)], res[len(chunks):]
        if any(rc):
            items, n = chunks[rc.index(True)], 2
        elif any(rp):
            items, n = comps[rp.index(True)], max(n - 1, 2)
        elif n >= len(items):
            break
        else:
            n = min(len(items), n * 2)
        log(f"      {len(items)} left (granularity {n})")
    return items


class Doc:
    """document.xml plus the package around it, rebuilt per candidate."""

    def __init__(self, pkg: Pkg) -> None:
        self.pkg = pkg
        self.xml = text(pkg, "word/document.xml")

    def build(self, xml: str) -> bytes:
        pkg = dict(self.pkg)
        put(pkg, "word/document.xml", xml)
        return write_pkg(pkg)


def minimize_container(doc: Doc, container: tuple[int, int, str], keep: Callable[[str], bool],
                       oracle: Oracle, log: Callable[[str], None], filler: str = "") -> str:
    """ddmin the direct children of `container` in doc.xml; children where keep(qname) are fixed."""
    xml = doc.xml
    rng = inner(xml, container)
    if rng is None:
        return xml
    kids = child_spans(xml, rng[0], rng[1])
    movable = [i for i, (_, _, q) in enumerate(kids) if not keep(q)]
    if len(movable) == 0:
        return xml

    def render(subset: list[int]) -> str:
        chosen = set(subset)
        drop = [(kids[i][0], kids[i][1]) for i in movable if i not in chosen]
        out = splice(xml, drop)
        if filler and not any(i in chosen for i in movable):
            # Keep the container valid (a cell must end with a paragraph).
            close = out.rfind(f"</{container[2]}", 0, container[1] - sum(e - s for s, e in drop))
            out = out[:close] + filler + out[close:]
        return out

    def test_batch(subsets: list[list[int]]) -> list[bool]:
        return oracle.test([doc.build(render(s)) for s in subsets])

    log(f"    {container[2]}: {len(movable)} children")
    best = ddmin(movable, test_batch, log)
    return render(best)


def body_span(xml: str) -> tuple[int, int, str]:
    span = find_element(xml, "w:body")
    if span is None:
        raise SystemExit("no w:body in word/document.xml")
    return span


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--out")
    ap.add_argument("--fail-on", choices=["any", "strict", "protected"], default="any")
    ap.add_argument("--loose", action="store_true", help="any failure counts, not only the original signature")
    ap.add_argument("--batch", type=int, default=16, help="files per Word job")
    ap.add_argument("--timeout-ms", type=int, default=30_000)
    ap.add_argument("--skip-ablations", action="store_true")
    args = ap.parse_args()

    src = Path(args.file).resolve()
    out = Path(args.out) if args.out else src.with_suffix(".bisect")
    out.mkdir(parents=True, exist_ok=True)
    started = time.time()
    lines: list[str] = []

    def log(msg: str) -> None:
        print(msg, flush=True)
        lines.append(msg)

    oracle = Oracle(args.fail_on, args.loose, args.batch, args.timeout_ms)
    original = src.read_bytes()
    pkg = read_pkg(original)
    report: dict = {"file": str(src), "failOn": args.fail_on}

    log(f"# confirm {src.name}")
    base, rezipped = oracle.run([original, write_pkg(pkg)])
    report["original"] = base
    if not oracle.failing(base):
        log(f"  input does not fail (verdict {base['verdict']}); nothing to bisect")
        report["conclusion"] = "input-ok"
        (out / "report.json").write_text(json.dumps(report, indent=2, ensure_ascii=False))
        return 0
    oracle.target = signature(base)
    log(f"  fails: {base['verdict']} signature={oracle.target}")
    for key in ("strict", "protectedView"):
        err = (base.get(key) or {}).get("error")
        if err:
            log(f"  {key}: {err.get('description', '').splitlines()[0]} [{err.get('scode') or err.get('hresult')}]")
    findings = [f for f in docxlint.lint_bytes(original) if f.severity == "error"]
    oracle.allowed_lint = {f.rule for f in findings}
    report["lint"] = [f"{f.rule} {f.part}: {f.message}" for f in findings]
    for f in findings:
        log(f"  lint: {f.rule} {f.part}: {f.message}")
    if not oracle.failing(rezipped):
        log("  a plain re-zip opens: the ZIP container is the trigger (entry flags, order, compression)")
        report["conclusion"] = "zip-container"
        (out / "report.json").write_text(json.dumps(report, indent=2, ensure_ascii=False))
        return 0

    cures: list[str] = []
    if not args.skip_ablations:
        log("# ablate")
        variants = []
        for name, fn in ABLATIONS:
            p = dict(pkg)
            try:
                fn(p)
            except Exception as err:  # an ablation that cannot apply is just skipped
                log(f"  {name}: skipped ({err})")
                continue
            if p == pkg:
                continue
            variants.append((name, write_pkg(p)))
        results = oracle.run([b for _, b in variants])
        for (name, _), r in zip(variants, results):
            cured = r["verdict"] == "ok"
            if cured:
                cures.append(name)
            log(f"  {name:28} -> {r['verdict']}{'   <-- CURES' if cured else ''}")
        report["ablations"] = {name: r["verdict"] for (name, _), r in zip(variants, results)}
        report["cures"] = cures

    log("# minimize body")
    doc = Doc(pkg)
    doc.xml = minimize_container(doc, body_span(doc.xml), lambda q: q == "w:sectPr", oracle, log)
    for qname, keep, filler in (
        ("w:tbl", lambda q: q in {"w:tblPr", "w:tblGrid"}, ""),
        ("w:tr", lambda q: q in {"w:trPr", "w:tblPrEx"}, ""),
        ("w:tc", lambda q: q == "w:tcPr", "<w:p/>"),
        ("w:p", lambda q: q == "w:pPr", ""),
    ):
        index = 0
        while True:
            spans = all_elements(doc.xml[body_span(doc.xml)[0]:], {qname})
            if index >= len(spans):
                break
            offset = body_span(doc.xml)[0]
            s, e, q = spans[index]
            before = doc.xml
            doc.xml = minimize_container(doc, (s + offset, e + offset, q), keep, oracle, log, filler)
            if doc.xml == before:
                index += 1
            else:
                index += 1
    minimal = doc.build(doc.xml)
    check = oracle.run([minimal])[0]
    if not oracle.failing(check):
        log("  warning: the minimized file no longer fails; Word is flaky for this input")
    (out / "minimal.docx").write_bytes(minimal)
    b = body_span(doc.xml)
    fragment = doc.xml[b[0]:b[1]]
    (out / "fragment.xml").write_text(fragment, encoding="utf-8")
    report.update({
        "conclusion": "minimized",
        "minimal": check,
        "fragmentBytes": len(fragment),
        "oracleCalls": oracle.calls,
        "driftRejected": oracle.drift_rejected,
        "filesTested": oracle.files,
        "seconds": round(time.time() - started, 1),
        "log": lines,
    })
    (out / "report.json").write_text(json.dumps(report, indent=2, ensure_ascii=False))
    log(f"# done in {report['seconds']}s, {oracle.files} Word opens")
    log(f"  cures: {', '.join(cures) or 'none'}")
    log(f"  minimal body ({len(fragment)} bytes) in {out / 'fragment.xml'}")
    print(fragment[:3000])
    return 0


if __name__ == "__main__":
    sys.exit(main())
