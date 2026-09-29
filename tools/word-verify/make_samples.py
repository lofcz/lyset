#!/usr/bin/env python3
"""Turn real print-document IRs into anonymous CI samples.

Every prose string is replaced by filler of the same length and word
shape, so layout and structure survive but no user content does. Enum-like
fields, formulas and numbers are kept. Images become a small grey PNG of
the same aspect. Output parses with lyset's strict IR schema.

    make_samples.py IR.json... --out samples/
"""
from __future__ import annotations

import argparse
import base64
import json
import re
import struct
import zlib
from pathlib import Path

# Fields whose values are part of the schema, not content.
KEEP = {"kind", "style", "tone", "variant", "locale", "mime", "tex", "mathml", "color", "align", "format",
        "orientation", "position", "fit", "shape", "layout", "direction", "reveal", "type", "listStyle",
        "pageNumberFormat", "numbering", "marker", "language", "lang", "accent", "href", "url", "fontFamily",
        "decoration", "size", "weight", "valign", "vAlign", "hAlign", "border", "fill", "background", "mode", "state"}
FILLER = "lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua"


def filler_like(text: str) -> str:
    words = iter(FILLER.split() * 50)
    def word(m: re.Match[str]) -> str:
        w = next(words)
        n = len(m.group(0))
        out = (w * (n // len(w) + 1))[:n]
        return out.capitalize() if m.group(0)[0].isupper() else out
    return re.sub(r"[^\W\d_]+", word, text)


def png(width: int, height: int) -> str:
    raw = b"".join(b"\x00" + b"\xc8\xc8\xc8" * width for _ in range(height))
    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
    data = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")
    return base64.b64encode(data).decode()


def image_size(b64: str) -> tuple[int, int]:
    try:
        head = base64.b64decode(b64[:64] + "=" * (-len(b64[:64]) % 4))
        if head[:8] == b"\x89PNG\r\n\x1a\n":
            w, h = struct.unpack(">II", head[16:24])
            return w, h
    except Exception:
        pass
    return 400, 300


def scrub(value, key: str | None = None):
    if isinstance(value, dict):
        out = {}
        for k, v in value.items():
            if k == "data" and isinstance(v, str) and value.get("mime", "").startswith("image/"):
                w, h = image_size(v)
                scale = max(1, max(w, h) // 64)
                out[k] = png(max(1, w // scale), max(1, h // scale))
                out["mime"] = "image/png"
            else:
                out[k] = scrub(v, k)
        return out
    if isinstance(value, list):
        return [scrub(v, key) for v in value]
    if isinstance(value, str) and key not in KEEP:
        return filler_like(value)
    return value


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("irs", nargs="+")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    for path in map(Path, args.irs):
        doc = scrub(json.loads(path.read_text()))
        name = f"{path.parent.name[:4]}-{path.stem.replace('.ir', '')}.ir.json"
        (out / name).write_text(json.dumps(doc, ensure_ascii=False, indent=1))
        print(out / name)


if __name__ == "__main__":
    main()
