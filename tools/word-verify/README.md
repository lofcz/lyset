# word-verify

Real Microsoft Word as a test oracle for DOCX output, plus the tools around
it: a static linter backed by Word evidence, a fixture generator that
produces that evidence, and a bisector that shrinks a failing file to the
smallest fragment Word still rejects.

| Piece | Runs on | What it does |
|---|---|---|
| `word-verify.exe` | Windows with Word | Opens files in a hidden Word, returns verdict and Word's own error text |
| `wordcheck.py` | host | Client: `check`, `status`, `log`, `deploy` |
| `docxlint.py` | anywhere, no Word | Static checks for package defects Word refuses or repairs. Runs in CI |
| `fixtures.py` | host | Mutates a good file into one defect per case, records Word's verdict next to the linter's |
| `docxbisect.py` | host | Ablates package features, then delta-debugs the body down to a minimal failing fragment |
| `make_samples.py` | host | Anonymises real print IRs into CI samples (`samples/`) |

## What Word does with a file

Each file goes through up to three passes on a private desktop, so nothing
appears on screen and modal prompts cannot block the run:

1. **strict**: `Documents.Open`, read-write, like a double-click on a local file.
2. **protected view**: `ProtectedViewWindows.Open`, then **Enable Editing**. Every
   browser download carries a mark of the web and opens this way.
3. **repair** (only after a failure): `OpenAndRepair`, then saves Word's
   repaired copy and reports which parts Word dropped.

The verdict is `ok`, `repair` (Word could recover it), `reject` or `timeout`.
Failures carry Word's exception text and number, for example:

| Word error | Message | Typical cause |
|---|---|---|
| 5121 | Word experienced an error trying to open the file | Not well-formed XML, a control character, `w:t` outside `w:r`, a row without cells, a foreign element that is not `mc:Ignorable`. Also what Protected View shows when its sandbox cannot start |
| 5792 | The file appears to be corrupted | Package defects: duplicate content types or relationship ids, missing parts, dangling `r:id`, NaN in `w:gridCol` |

`fixtures.py` output against Word 16.0.20430 is the source of this table.

## Setup (once per VM)

The service runs inside a Windows VM with Word, reachable from the host on
TCP 47400. With Winboat, add the port to `~/.winboat/docker-compose.yml`
(`USER_PORTS: …,47400` and `- 127.0.0.1:47400:47400/tcp`) and recreate the
container. Ports 44298 and 44300 belong to local Priprava, do not reuse them.

```bash
cargo xwin build --release --target x86_64-pc-windows-msvc   # in this folder
python3 wordcheck.py deploy                                   # stages exe, launcher, token on the share
```

In the VM, run `\\host.lan\Data\_win_share\word-verify\install-autostart.bat`
once. It copies the service to `%LOCALAPPDATA%\lyset-word-verify`, opens the
firewall port and registers a logon task. Later builds are pushed with
`wordcheck.py deploy`, which hot-swaps the running service.

The host talks to the VM over HTTP, not through the shared folder: the
Windows SMB client caches listings of a share whose files change behind
Samba's back, which delayed job pickup by minutes.

## Everyday use

```bash
python3 wordcheck.py check out/*.docx                  # verdicts, Word errors, stats
python3 wordcheck.py check --pdf --out /tmp/w file.docx  # also Word's own PDF
python3 docxlint.py out/*.docx                        # no Word needed
python3 docxbisect.py failing.docx                     # minimal failing fragment
python3 fixtures.py good.docx --out /tmp/fixtures      # re-derive the evidence
```

Exit codes: 0 all good, 1 some file failed, 2 the service or setup is broken.

## Caveats

- Word only reports what the machine it runs on does. A file that opens here
  can still fail on a machine whose Protected View sandbox is broken, or under
  a security product or policy the VM does not have.
- Protected View needs the private desktop to admit low-integrity and
  AppContainer processes. `isolate.rs` sets that descriptor. Without it every
  Protected View open fails with error 5121, which looks exactly like a bad file.
