#!/usr/bin/env python3
"""Host-side client for the word-verify service running in a Windows VM.

The VM forwards TCP 47400 to the host's loopback. The service opens each
file in a hidden Microsoft Word and returns ok / repair / reject with Word's
own error text.

    wordcheck.py check [--no-repair] [--no-protected-view] [--pdf] [--timeout-ms N] [--json] [--out DIR] FILE...
    wordcheck.py status
    wordcheck.py log
    wordcheck.py deploy     # stage exe, launcher and token on the share; hot-update a running service

Environment: WORD_VERIFY_URL (default http://127.0.0.1:47400),
WORD_VERIFY_HOME (the share folder, default ../../../_win_share/word-verify
relative to this file), WORD_VERIFY_TOKEN (default: <home>/token.txt).
Exit codes: 0 all ok, 1 some file not ok, 2 infrastructure problem.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import secrets
import shutil
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Iterable

HERE = Path(__file__).resolve().parent
EXE = HERE / "target/x86_64-pc-windows-msvc/release/word-verify.exe"


class InfraError(RuntimeError):
    pass


def home() -> Path:
    env = os.environ.get("WORD_VERIFY_HOME")
    return Path(env) if env else (HERE / "../../../_win_share/word-verify").resolve()


def base_url() -> str:
    return os.environ.get("WORD_VERIFY_URL", "http://127.0.0.1:47400").rstrip("/")


def token() -> str | None:
    if os.environ.get("WORD_VERIFY_TOKEN"):
        return os.environ["WORD_VERIFY_TOKEN"]
    path = home() / "token.txt"
    return path.read_text().strip() if path.is_file() else None


def request(method: str, path: str, body: bytes | None = None, *, timeout: float = 30,
            content_type: str = "application/json") -> dict:
    req = urllib.request.Request(base_url() + path, data=body, method=method)
    if body is not None:
        req.add_header("Content-Type", content_type)
    tok = token()
    if tok:
        req.add_header("Authorization", f"Bearer {tok}")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read() or b"{}")
    except urllib.error.HTTPError as err:
        detail = err.read().decode("utf-8", "replace")[:500]
        raise InfraError(f"{method} {path}: HTTP {err.code} {detail}") from None
    except (urllib.error.URLError, ConnectionError, TimeoutError, OSError) as err:
        raise InfraError(f"{method} {path}: {err} (is the word-verify service running in the VM?)") from None


def health() -> dict:
    return request("GET", "/health", timeout=10)


def verify(files: Iterable[Path | tuple[str, bytes]], *, repair: bool = True, pdf: bool = False,
           protected_view: bool = True, timeout_ms: int = 45_000) -> dict:
    """Open each file in Word. Items are paths or (name, bytes) pairs."""
    payload_files = []
    for item in files:
        name, data = (item.name, item.read_bytes()) if isinstance(item, Path) else item
        payload_files.append({"name": name, "data": base64.b64encode(data).decode()})
    body = json.dumps({"files": payload_files, "repair": repair, "pdf": pdf, "protectedView": protected_view,
                       "timeoutMs": timeout_ms}).encode()
    passes = 1 + (1 if protected_view else 0) + (1 if repair else 0)
    budget = 60 + len(payload_files) * (timeout_ms / 1000) * (passes + 0.3)
    return request("POST", "/verify", body, timeout=budget)


def save_artifacts(result: dict, out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for r in result.get("results", []):
        stem = Path(r["file"]).stem
        for key, suffix in (("pdfBase64", ".word.pdf"), ("repairedBase64", ".word-repaired.docx")):
            data = r.pop(key, None)
            if data:
                path = out_dir / f"{stem}{suffix}"
                path.write_bytes(base64.b64decode(data))
                r[key.replace("Base64", "Path")] = str(path)


def describe(r: dict) -> str:
    lines = [f"{r['verdict'].upper():8} {r.get('source', r['file'])}"]
    strict = r.get("strict") or {}
    err = strict.get("error")
    if err:
        code = err.get("scode") or err.get("hresult")
        word_no = f" Word#{err['wordError']}" if err.get("wordError") is not None else ""
        lines.append(f"         strict open: {err.get('description')} [{code}{word_no}]")
    stats = strict.get("stats") or (r.get("repair") or {}).get("stats")
    if stats:
        lines.append("         stats: " + ", ".join(f"{k}={v}" for k, v in stats.items() if v is not None))
    for d in r.get("dialogs", []):
        clicked = f"  -> clicked {d['clicked']}" if d.get("clicked") else ""
        lines.append(f"         dialog[{d['kind']}]: {d['text'][:300]}{clicked}")
    pv = r.get("protectedView")
    if pv:
        pv_err = pv.get("error")
        state = "opened" if pv.get("opened") else "FAILED"
        if pv.get("opened") and pv.get("edited") is not None:
            state += ", enable editing " + ("ok" if pv["edited"] else "FAILED")
        detail = f": {pv_err.get('description')} [{pv_err.get('scode') or pv_err.get('hresult')}]" if pv_err else ""
        lines.append(f"         protected view: {state}{detail}")
    rep = r.get("repair")
    if rep:
        why = f" ({rep['error']['description']})" if rep.get("error") else ""
        lines.append(f"         repair pass: {'recovered' if rep.get('opened') else 'failed'}{why}")
    diff = r.get("packageDiff")
    if diff and (diff.get("removed") or diff.get("droppedRelationships")):
        lines.append(f"         Word dropped parts: {diff.get('removed')} rels: {diff.get('droppedRelationships')}")
    for key in ("pdfPath", "repairedPath"):
        if r.get(key):
            lines.append(f"         {key}: {r[key]}")
    return "\n".join(lines)


def cmd_check(args: argparse.Namespace) -> int:
    files = [Path(f).resolve() for f in args.files]
    missing = [str(f) for f in files if not f.is_file()]
    if missing:
        print(f"missing: {', '.join(missing)}", file=sys.stderr)
        return 2
    try:
        result = verify(files, repair=not args.no_repair, pdf=args.pdf, protected_view=not args.no_protected_view,
                        timeout_ms=args.timeout_ms)
    except InfraError as err:
        print(f"wordcheck: {err}", file=sys.stderr)
        return 2
    for r, f in zip(result.get("results", []), files):
        r["source"] = str(f)
    save_artifacts(result, Path(args.out) if args.out else files[0].parent)
    if args.json:
        print(json.dumps(result, indent=2, ensure_ascii=False))
    else:
        word = result.get("word") or {}
        print(f"# {result.get('tool')} / Word {word.get('version')} build {word.get('build')}")
        for r in result.get("results", []):
            print(describe(r))
    return 0 if all(r["verdict"] == "ok" for r in result.get("results", [])) else 1


def cmd_status(_: argparse.Namespace) -> int:
    try:
        state = health()
    except InfraError as err:
        print(f"wordcheck: {err}", file=sys.stderr)
        return 2
    local = sha256(EXE) if EXE.is_file() else None
    state["localExeSha256"] = local
    state["upToDate"] = local == state.get("exeSha256")
    print(json.dumps(state, indent=2))
    return 0


def cmd_log(_: argparse.Namespace) -> int:
    try:
        print("\n".join(request("GET", "/log").get("lines", [])))
    except InfraError as err:
        print(f"wordcheck: {err}", file=sys.stderr)
        return 2
    return 0


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def cmd_deploy(_: argparse.Namespace) -> int:
    if not EXE.is_file():
        print("build first: cargo xwin build --release --target x86_64-pc-windows-msvc", file=sys.stderr)
        return 2
    dest = home()
    dest.mkdir(parents=True, exist_ok=True)
    tok = dest / "token.txt"
    if not tok.is_file():
        tok.write_text(secrets.token_hex(24) + "\n")
        print(f"created {tok}")
    for src in [EXE, *sorted((HERE / "guest").glob("*.bat"))]:
        tmp = dest / (src.name + ".tmp")
        shutil.copyfile(src, tmp)
        os.replace(tmp, dest / src.name)
        print(f"staged {dest / src.name}")
    want = sha256(EXE)
    try:
        running = health()
    except InfraError:
        print("service not reachable: run install-autostart.bat from the share inside the VM")
        return 0
    if running.get("exeSha256") == want:
        print("service already runs this build")
        return 0
    print("pushing the new build to the running service")
    try:
        request("POST", "/update", EXE.read_bytes(), content_type="application/octet-stream", timeout=60)
    except InfraError as err:
        print(f"wordcheck: {err}", file=sys.stderr)
        return 2
    deadline = time.time() + 60
    while time.time() < deadline:
        time.sleep(1.5)
        try:
            if health().get("exeSha256") == want:
                print("service restarted on the new build")
                return 0
        except InfraError:
            continue
    print("wordcheck: service did not come back on the new build within 60s", file=sys.stderr)
    return 2


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)
    check = sub.add_parser("check", help="open files in Word and report")
    check.add_argument("files", nargs="+")
    check.add_argument("--no-repair", action="store_true", help="skip the OpenAndRepair pass")
    check.add_argument("--no-protected-view", action="store_true", help="skip the Protected View pass")
    check.add_argument("--pdf", action="store_true", help="also save Word's own PDF rendering")
    check.add_argument("--timeout-ms", type=int, default=45_000)
    check.add_argument("--json", action="store_true")
    check.add_argument("--out", help="where to save PDFs and repaired copies (default: next to the first file)")
    check.set_defaults(func=cmd_check)
    sub.add_parser("status").set_defaults(func=cmd_status)
    sub.add_parser("log").set_defaults(func=cmd_log)
    sub.add_parser("deploy").set_defaults(func=cmd_deploy)
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
