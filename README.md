# word-engine-mcp

[한국어 안내 → README.ko.md](README.ko.md)

A local MCP server that drives the **real Word desktop application** (COM automation) to do the things python-docx fundamentally cannot: **compute field/TOC values, count real pages, render true-fidelity PDF, convert/repair legacy formats, produce redline comparisons, and run native mail merge**.

> Design philosophy: this server **complements** library workflows instead of replacing them.
> Editing text, tables and styles is faster with python-docx — but a document built with
> python-docx has `?` where TOC page numbers should be, and no library has a concept of
> "pages" (that requires a layout engine). This MCP handles only the engine-exclusive part,
> keeping its tool surface tiny (8 tools).

## Requirements

- Windows 10+ with a **logged-in interactive desktop session** (Word has no true headless mode)
- Microsoft Office (Word) installed and licensed — verified on **Office 2016+ (Word 16.0)**
- Python 3.10+ — verified on **3.12**
- [Claude Code](https://claude.com/claude-code) or any MCP client

## Install

```powershell
git clone https://github.com/Feynman520/d01-p04-word-engine-mcp.git
cd d01-p04-word-engine-mcp
py -3.12 -m venv .venv          # or: python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### Register with Claude Code

Run this in the cloned folder (uses absolute paths, so it works from anywhere afterwards):

```powershell
claude mcp add word-automation --scope user -- "$PWD\.venv\Scripts\python.exe" "$PWD\server.py"
```

`--scope user` makes it available in every project. Use `--scope project` to limit it to one project.

### Verify

```powershell
$py = ".\.venv\Scripts\python.exe"; $env:PYTHONUTF8 = "1"
& $py tests\smoke_com.py      # field update / page stats / PDF / convert / compare / mail merge
& $py tests\server_tools.py   # MCP tool registration (does not launch Word)
```

## Tools (6 core + 2 diagnostics)

| # | Tool | Input → Output | Why engine-only |
|---|---|---|---|
| ① | `word_update_fields` | `src_path, out_path` → `{out_path,fields_updated}` | Computes displayed values of TOC, cross-references, PAGE/SEQ, index (python-docx stores field codes only — shown as `?`) |
| ② | `word_read_layout` | `path, update_fields?, include_readability?` → `{pages,words,lines,characters,...}` | **Real page count** and layout statistics (libraries have no page concept) |
| ③ | `word_export_pdf` | `src_path, out_path, update_fields?, pdf_a?, from_page?, to_page?, create_bookmarks?` → `{out_path}` | Word render engine PDF (PDF/A, heading bookmarks, page ranges) |
| ④ | `word_convert` | `src_path, out_path, repair?` → `{out_path,format,repaired}` | `.doc`/`.rtf`/`.odt`/`.html`/`.txt` ↔ `.docx` conversion and **corrupt-file repair** (python-docx is .docx-only) |
| ⑤ | `word_compare` | `original_path, revised_path, out_path, author?, granularity?` → `{out_path,revisions}` | Native `CompareDocuments` redline (no library equivalent) |
| ⑥ | `word_mail_merge` | `template_path, data_path, out_path` → `{out_path,records}` | Native mail merge (per-record documents) |
| — | `word_health` | → `{alive, word_version}` | Session check (launches Word on first call) |
| — | `word_restart` | → `{alive}` | Recovery from COM errors |

Typical flow: build a `.docx` with python-docx → `word_update_fields` to bake TOC/numbers →
`word_read_layout` to check the real page count → `word_export_pdf` for the final PDF.
Originals are never modified; results are always written to `out_path`.

### Mail merge data source

The most reliable source is a **`.docx` containing a table with one header row + data rows**
(no database driver involved, fully unattended). `.csv`/`.xlsx` are also accepted, but some
environments show Word's SQL confirmation prompt (depends on an HKCU setting).

## Architecture notes

- **Single STA worker thread** (`engine/session.py`): every Word call is serialized onto one
  dedicated thread (win32com COM objects are thread-bound; FastMCP may hop threads).
- **Lazy session**: Word starts on the first tool call, is reused, and closes with the server.
- **`DispatchEx` + early binding (`gencache.EnsureDispatch`)**: a dedicated instance, and
  argument-heavy methods like `ExportAsFixedFormat`/`CompareDocuments` are marshalled
  correctly via the type library. The one-time makepy output is redirected away from stdout
  (the JSON-RPC channel).
- **`Visible=False`**, macros blocked on open via `AutomationSecurity=ForceDisable`.
- **Originals preserved**: inputs open read-only, results are written to new paths, then
  `Close(SaveChanges=False)`.
- **Zombie prevention**: Word gives no stable window handle, so the dedicated instance PID is
  identified by diffing the `WINWORD.EXE` process list before/after `DispatchEx`, and
  force-killed at shutdown if it survives `Quit()`.
- **RPC-rejection retry**: `RPC_E_CALL_REJECTED` right after startup is retried with backoff.

## Limitations

- Not suitable for unattended/service sessions (needs an interactive desktop).
- Text/table/style editing is faster with python-docx — that is by design.

## License

[MIT](LICENSE)
