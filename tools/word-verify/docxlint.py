#!/usr/bin/env python3
"""Static checks for DOCX packages, for the defects that make Word refuse or
repair a file. Runs anywhere (no Word), so it belongs in CI next to the
generator. `fixtures.py` builds one mutated file per rule and records what
real Word does with it, which is where each rule's severity comes from.

    docxlint.py FILE.docx... [--json]

Exit 1 when any file has an error-level finding.
"""
from __future__ import annotations

import argparse
import io
import json
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass
from pathlib import Path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
MAIN_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
OFFICE_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
# Characters XML 1.0 forbids even as character references.
BAD_CHARS = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]|&#(?:x0*(?:[0-8bcef]|1[0-9a-f])|0*(?:[0-8]|1[124-9]|2[0-9]|3[01]));", re.I)


# Namespaces Word understands. Elements outside these must be mc:Ignorable.
KNOWN_NS = (
    "http://schemas.openxmlformats.org/",
    "http://purl.org/",
    "http://www.w3.org/",
    "http://schemas.microsoft.com/",
    "urn:schemas-microsoft-com:",
)
# Run content that is only valid inside <w:r>.
RUN_CONTENT = {"t", "tab", "br", "cr", "sym", "drawing", "pict", "object", "fldChar", "instrText", "delText", "delInstrText",
               "noBreakHyphen", "softHyphen", "footnoteReference", "endnoteReference", "commentReference", "lastRenderedPageBreak", "rPr"}
# Integer (twips, half-points, eighths) attributes Word parses strictly.
NUMERIC_ATTRS = {"w", "h", "top", "bottom", "left", "right", "header", "footer", "gutter", "before", "after", "line", "pos", "sz", "space", "firstLine", "hanging", "start", "end"}
# Elements whose w:val / w:w etc. are enumerations or strings, not numbers.
NON_NUMERIC_VAL = {"lang", "rFonts", "pStyle", "rStyle", "tblStyle", "jc", "vAlign", "color", "shd", "highlight", "u", "tab"}


@dataclass
class Finding:
    rule: str
    severity: str  # error | warning
    part: str
    message: str


class Lint:
    def __init__(self, data: bytes) -> None:
        self.findings: list[Finding] = []
        self.zip = zipfile.ZipFile(io.BytesIO(data))
        self.parts: dict[str, bytes] = {}
        self.trees: dict[str, ET.Element] = {}
        self.raw: dict[str, str] = {}
        self._rels: dict[str, dict[str, tuple[str, str, bool]]] = {}

    def add(self, rule: str, severity: str, part: str, message: str) -> None:
        self.findings.append(Finding(rule, severity, part, message))

    # -- container ---------------------------------------------------------
    def check_zip(self) -> None:
        seen: dict[str, str] = {}
        for info in self.zip.infolist():
            name = info.filename
            if name.endswith("/"):
                self.add("zip-directory-entry", "warning", name, "directory entry in package")
                continue
            if "\\" in name:
                self.add("zip-backslash-name", "error", name, "part name uses a backslash")
            key = name.lower()
            if key in seen:
                self.add("zip-duplicate-entry", "error", name, f"duplicate of {seen[key]} (part names are case-insensitive)")
            seen[key] = name
            if info.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
                self.add("zip-compression", "error", name, f"compression method {info.compress_type}")
            if info.flag_bits & 0x1:
                self.add("zip-encrypted", "error", name, "encrypted entry")
            if info.file_size >= 0xFFFFFFFF or info.compress_size >= 0xFFFFFFFF:
                self.add("zip-zip64", "warning", name, "ZIP64 entry")
            self.parts[name] = self.zip.read(info)

    # -- XML ---------------------------------------------------------------
    def check_xml(self) -> None:
        for name, data in self.parts.items():
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            try:
                raw = data.decode("utf-8")
            except UnicodeDecodeError as err:
                self.add("xml-encoding", "error", name, f"not UTF-8: {err}")
                continue
            self.raw[name] = raw
            m = BAD_CHARS.search(raw)
            if m:
                ctx = raw[max(0, m.start() - 30):m.end() + 30]
                self.add("xml-illegal-char", "error", name, f"character not allowed in XML 1.0 near {ctx!r}")
            try:
                self.trees[name] = ET.fromstring(data)
            except ET.ParseError as err:
                self.add("xml-malformed", "error", name, str(err))
                continue
            ignorable = re.search(r'\bmc:Ignorable="([^"]*)"', raw)
            if ignorable:
                declared = set(re.findall(r'xmlns:([\w.\-]+)=', raw))
                for prefix in ignorable.group(1).split():
                    if prefix not in declared:
                        self.add("mc-ignorable-undeclared", "error", name, f"mc:Ignorable names undeclared prefix {prefix!r}")

    # -- content types -------------------------------------------------------
    def check_content_types(self) -> dict[str, str]:
        root = self.trees.get("[Content_Types].xml")
        if root is None:
            self.add("ct-missing", "error", "[Content_Types].xml", "no content types part")
            return {}
        defaults: dict[str, str] = {}
        overrides: dict[str, str] = {}
        for el in root:
            tag = el.tag.split("}")[-1]
            if tag == "Default":
                ext = (el.get("Extension") or "").lower()
                if ext in defaults:
                    self.add("ct-duplicate-default", "error", "[Content_Types].xml", f"Default Extension {ext!r} listed twice")
                defaults[ext] = el.get("ContentType") or ""
            elif tag == "Override":
                part = (el.get("PartName") or "").lower()
                if part in overrides:
                    self.add("ct-duplicate-override", "error", "[Content_Types].xml", f"Override {el.get('PartName')!r} listed twice")
                overrides[part] = el.get("ContentType") or ""
                if part.lstrip("/") not in {n.lower() for n in self.parts}:
                    self.add("ct-override-missing-part", "warning", "[Content_Types].xml", f"Override for missing part {el.get('PartName')!r}")
        types: dict[str, str] = {}
        for name in self.parts:
            if name == "[Content_Types].xml":
                continue
            base = posixpath.basename(name)
            ext = base.rsplit(".", 1)[-1].lower() if "." in base else ""
            ct = overrides.get("/" + name.lower()) or defaults.get(ext)
            if not ct:
                self.add("ct-part-untyped", "error", name, "part has no content type")
            else:
                types[name] = ct
        return types

    # -- relationships -------------------------------------------------------
    def rels_for(self, part: str) -> dict[str, tuple[str, str, bool]]:
        if part not in self._rels:
            self._rels[part] = self._read_rels(part)
        return self._rels[part]

    def _read_rels(self, part: str) -> dict[str, tuple[str, str, bool]]:
        d, n = posixpath.split(part)
        path = posixpath.join(d, "_rels", n + ".rels")
        root = self.trees.get(path)
        out: dict[str, tuple[str, str, bool]] = {}
        if root is None:
            return out
        for el in root:
            rid, ty, target = el.get("Id") or "", el.get("Type") or "", el.get("Target") or ""
            external = el.get("TargetMode") == "External"
            if rid in out:
                self.add("rel-duplicate-id", "error", path, f"relationship Id {rid!r} used twice")
            if not external:
                resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(d, target))
                if resolved.lower() not in {p.lower() for p in self.parts}:
                    self.add("rel-missing-target", "error", path, f"{rid} ({ty.rsplit('/', 1)[-1]}) points at missing part {target!r}")
                if target.startswith("/"):
                    self.add("rel-absolute-target", "warning", path, f"{rid} uses absolute target {target!r}")
                target = resolved
            out[rid] = (ty, target, external)
        return out

    def check_relationships(self, types: dict[str, str]) -> str | None:
        root_rels = self.rels_for("")
        main = [t for ty, t, ext in root_rels.values() if ty == OFFICE_DOC and not ext]
        if len(main) != 1:
            self.add("rel-office-document", "error", "_rels/.rels", f"{len(main)} officeDocument relationships")
            return None
        if types.get(main[0]) != MAIN_CT:
            self.add("ct-main-document", "error", main[0], f"main part content type is {types.get(main[0])!r}")
        for name in self.parts:
            if not name.endswith(".rels") and name != "[Content_Types].xml":
                self.rels_for(name)
        return main[0]

    # -- WordprocessingML ----------------------------------------------------
    def check_story(self, part: str) -> None:
        root = self.trees.get(part)
        if root is None:
            return
        rels = self.rels_for(part)
        for el in root.iter():
            for key, value in el.attrib.items():
                if key.startswith("{" + R + "}") and key.split("}")[1] in ("id", "embed", "link", "pict") and value not in rels:
                    self.add("rel-dangling-rid", "error", part, f"<{el.tag.split('}')[-1]} r:{key.split('}')[1]}={value!r}> has no relationship")
        for tbl in root.iter(f"{{{W}}}tbl"):
            rows = tbl.findall(f"{{{W}}}tr")
            if not rows:
                # Word 16 opens this (fixture wml-table-without-rows), but it is never intended.
                self.add("wml-table-no-rows", "warning", part, "table without rows")
            if tbl.find(f"{{{W}}}tblGrid") is None:
                self.add("wml-table-no-grid", "warning", part, "table without tblGrid")
            for tr in rows:
                if not tr.findall(f"{{{W}}}tc") and not tr.findall(f"{{{W}}}sdt"):
                    self.add("wml-row-no-cells", "error", part, "table row without cells")
        for tc in root.iter(f"{{{W}}}tc"):
            blocks = [c for c in tc if c.tag.split("}")[-1] not in ("tcPr",)]
            if not blocks or blocks[-1].tag != f"{{{W}}}p":
                self.add("wml-cell-last-not-paragraph", "error", part, "table cell must end with a paragraph")
        for el in root.iter():
            for key, value in el.attrib.items():
                if key.startswith("{" + W + "}") and key.split("}")[1] in NUMERIC_ATTRS and el.tag.split("}")[-1] not in NON_NUMERIC_VAL:
                    where = f"<w:{el.tag.split('}')[-1]} w:{key.split('}')[1]}={value!r}>"
                    if re.fullmatch(r"-?\d+", value) or (key.endswith("}val") and not re.search(r"\d", value)):
                        continue
                    if re.fullmatch(r"-?\d+\.\d+", value):
                        # Word 16 rounds these (fixtures decimal-twips-*); other consumers may not.
                        self.add("wml-fractional-number", "warning", part, f"{where} is fractional")
                    else:
                        # NaN in w:gridCol makes Word repair the file (fixture nan-grid-col).
                        self.add("wml-bad-number", "error", part, f"{where} is not a number")
        for p_el in root.iter(f"{{{W}}}p"):
            for child in p_el:
                local = child.tag.split("}")[-1]
                if child.tag.startswith("{" + W + "}") and local in RUN_CONTENT:
                    # Word refuses the whole file (5121, fixture text-outside-run).
                    self.add("wml-run-content-outside-run", "error", part, f"<w:{local}> is a direct child of <w:p>")
        ignorable_uris = self.ignorable_namespaces(part)
        for el in root.iter():
            if not el.tag.startswith("{"):
                continue
            uri = el.tag[1:].split("}")[0]
            if not uri.startswith(KNOWN_NS) and uri not in ignorable_uris:
                # Word refuses the whole file (5121, fixture foreign-element-not-ignorable).
                self.add("foreign-element", "error", part, f"element {el.tag.split('}')[-1]!r} in namespace {uri!r} is not covered by mc:Ignorable")
                break
        ids: dict[str, int] = {}
        for docpr in root.iter(f"{{{WP}}}docPr"):
            ids[docpr.get("id", "")] = ids.get(docpr.get("id", ""), 0) + 1
        for value, count in ids.items():
            if count > 1:
                self.add("wml-duplicate-docpr-id", "warning", part, f"wp:docPr id {value!r} used {count} times")
        if part.endswith("document.xml"):
            body = root.find(f"{{{W}}}body")
            if body is not None and len(body) and body[-1].tag != f"{{{W}}}sectPr" and body.find(f"{{{W}}}sectPr") is not None:
                self.add("wml-sectpr-not-last", "warning", part, "body sectPr is not the last child")

    def ignorable_namespaces(self, part: str) -> set[str]:
        raw = self.raw.get(part, "")
        declared = dict(re.findall(r'xmlns:([\w.\-]+)="([^"]*)"', raw))
        out: set[str] = set()
        for value in re.findall(r'\b[\w.\-]+:Ignorable="([^"]*)"', raw):
            out.update(declared[p] for p in value.split() if p in declared)
        return out

    def check_fonts(self) -> None:
        part = "word/fontTable.xml"
        root = self.trees.get(part)
        if root is None:
            return
        rels = self.rels_for(part)
        for el in root.iter():
            tag = el.tag.split("}")[-1]
            if not tag.startswith("embed"):
                continue
            rid = el.get(f"{{{R}}}id")
            key = el.get(f"{{{W}}}fontKey") or ""
            if rid not in rels:
                continue
            target = rels[rid][1]
            data = self.parts.get(target)
            if data is None:
                continue
            if not re.fullmatch(r"\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}", key):
                self.add("font-key-format", "error", part, f"fontKey {key!r} is not a braced GUID")
                continue
            hexkey = key.strip("{}").replace("-", "")
            k = bytes(int(hexkey[i:i + 2], 16) for i in range(30, -1, -2))
            head = bytes(data[i] ^ k[i % 16] for i in range(4)) + data[4:4]
            if head not in (b"\x00\x01\x00\x00", b"OTTO", b"true", b"ttcf"):
                self.add("font-obfuscation", "error", target, f"de-obfuscated header {head.hex()} is not a font (wrong fontKey?)")

    def check_app_properties(self) -> None:
        """Word parses AppVersion as a number with the user's regional format:
        a semver value opens on en-US but is error 5121 on cs-CZ/sk-SK/pl-PL."""
        part = "docProps/app.xml"
        data = self.parts.get(part)
        if data is None:
            return
        match = re.search(rb"<(?:\w+:)?AppVersion>([^<]*)</(?:\w+:)?AppVersion>", data)
        if match and not re.fullmatch(rb"\d{1,2}\.\d{4}", match.group(1).strip()):
            self.add("app-version-format", "error", part,
                     f"AppVersion {match.group(1).decode(errors='replace')!r} is not XX.YYYY; Word on comma-decimal locales refuses the file (5121)")

    def run(self) -> list[Finding]:
        self.check_zip()
        self.check_xml()
        types = self.check_content_types()
        main = self.check_relationships(types)
        if main:
            story_types = ("header", "footer", "footnotes", "endnotes", "comments")
            stories = [main] + [t for ty, t, ext in self.rels_for(main).values() if not ext and ty.rsplit("/", 1)[-1] in story_types]
            for story in stories:
                self.check_story(story)
        self.check_fonts()
        self.check_app_properties()
        return self.findings


def lint_bytes(data: bytes) -> list[Finding]:
    try:
        return Lint(data).run()
    except zipfile.BadZipFile as err:
        return [Finding("zip-invalid", "error", "", str(err))]


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()
    failed = False
    report = {}
    for f in args.files:
        findings = lint_bytes(Path(f).read_bytes())
        report[f] = [asdict(x) for x in findings]
        failed |= any(x.severity == "error" for x in findings)
        if not args.json:
            print(f"{'FAIL' if any(x.severity == 'error' for x in findings) else 'ok  '} {f}")
            for x in findings:
                print(f"     {x.severity:7} {x.rule:32} {x.part}: {x.message}")
    if args.json:
        print(json.dumps(report, indent=2, ensure_ascii=False))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
