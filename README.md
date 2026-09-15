# word-mcp-live

**A semantic engine for Word documents, exposed over [MCP](https://modelcontextprotocol.io/).**

An AI agent that edits a `.docx` by rewriting its text loses everything the text is not:
character styles, theme fonts, fields, bookmarks, comment anchors, tracked changes,
numbering, section layout. This server exists so that it does not have to. Its core is a
pure OOXML engine that works on the XML of the package itself — it runs on Linux, in a
container, on a machine where Microsoft Word has never been installed — and every edit it
makes is scoped to the range it was asked to change.

On top of that core it exposes three families of tools: the cross-platform document tools,
a stateless `doc_*` surface designed for agents, and live tools that drive a running copy
of Word on Windows and macOS.

[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)

---

## What makes it different

- **Non-degradation is the contract, not a side effect.** Text and formatting outside the
  targeted range come out byte-identical; parts of the package that were not targeted are
  untouched. These are tested invariants, not intentions — see
  [Fidelity guarantees](#fidelity-guarantees).
- **It refuses rather than approximates.** A range that would cut a field in half or
  swallow an image raises a typed error instead of producing a document Word has to repair
  on open.
- **Stateless addressing.** A `doc_*` locator is a plain dict an agent can store and send
  back later. There is no session, no handle, no id table to keep alive — and a locator
  that no longer describes the document fails loudly instead of quietly hitting the wrong
  paragraph.
- **Dry run everywhere.** Every mutating `doc_*` tool takes `dry_run`, and reports the
  before/after text of each paragraph it *would* change without writing anything.
- **No Word required for the core.** The engine reads and writes the `.docx` package
  directly. Word is needed only by the live tools.

## Architecture

The engine lives in `word_document_server/engine/`. It is synchronous Python over `lxml`,
sitting on the OPC implementation shipped with python-docx. It has no MCP dependency and
no Word dependency, and it reports failure through typed exceptions rather than return
codes.

The layers, bottom to top — each one only knows about the ones below it:

| Layer | Module | What it owns |
|---|---|---|
| Package | `package` | Opening a `.docx`, reaching its parts and stories, saving it atomically. Content types and `.rels` are delegated to `docx.opc` rather than hand-rolled. |
| Text model | `textmodel` | What a paragraph *shows*, and where each character sits. The single definition of a character offset. |
| Ranges | `ranges` | The only layer that writes text: split, delete, insert, replace over `[start, end)`. |
| Formatting | `format` | Writes `w:rPr` onto the runs a range already resolved to — property by property, never by rebuilding runs. |

Above them sit the modules that answer a specific question, each on top of that same base:

- `find` — where a string occurs, searched on the visible text of each paragraph.
- `locators` — resolving a stateless locator to a place in the document.
- `inspect` — what is in this document, and how do I name it.
- `revisions` — recording, listing, accepting and rejecting `w:ins` / `w:del`.
- `styles`, `theme`, `effective` — what a style says, what the theme resolves to, and what
  a reader actually sees once both are applied.
- `numbering` — list definitions in `numbering.xml`, and the `w:numPr` that points at them.
- `table_styles` — applying and clearing an existing table style, with its `w:tblLook` mask.
- `merge` — appending one package into another by importing elements, not text.

Three small modules are shared throughout: `xmlns` (the namespace table and `qn()`), `ids`
(allocating the identifiers OOXML requires to be unique), and `errors` (the typed exception
family).

### Text policy

One decision runs through the whole engine: the offset `0` of a paragraph means the same
thing in `find`, in `ranges`, in a locator's `expect_text` and in a `changes` report,
because exactly one module decides which characters exist. `textmodel` classifies every
element it meets into one of four buckets:

- **skipped** — property containers (`w:pPr`, `w:rPr`, `w:sdtPr`, …) carry formatting, not
  content.
- **visible** — text a reader sees, including runs wrapped in `w:hyperlink`, `w:ins`,
  `w:sdt` or `w:fldSimple`, which `paragraph.text` silently drops.
- **hidden** — text under a tracked deletion, and field instructions: present in the XML,
  not part of what the document reads.
- **opaque / markers** — images, field characters, note and comment references, bookmark
  boundaries: zero-width positions that carry no character but that a range must not cut
  through.

### Errors

Every deliberate failure derives from `EngineError`, so one `except` clause catches the
family. Operating-system failures are *not* wrapped: a missing file is still a
`FileNotFoundError`. A locator failure carries one of four codes — `invalid`, `not_found`,
`ambiguous`, `stale_anchor` — and a refused range carries a reason such as
`opaque-content`, `empty-range` or `unsplittable-range`. Callers branch on the code; humans
read the message, which names the closest candidates when there are any.

## The three tool families

`word_document_server/tools/platforms.py` is the single source of truth for which platforms
each tool runs on. It feeds [TOOLS.md](TOOLS.md), the counts below, and the
`doc_capabilities` tool.

### 1. Cross-platform document tools

The historical python-docx tools — create, format, tables, comments, footnotes, tracked
changes, layout, protection. They work on a **closed** file, on any OS.

These have been migrated onto the engine. The invariant held across that migration is on
their **names, parameters and return formats**: a client calling `format_text` or
`search_and_replace` calls it exactly as before. It is not a promise that nothing
changed — a handful of behaviours were deliberately fixed or tightened, and each one is
listed under `[Unreleased]` in [CHANGELOG.md](CHANGELOG.md). Read it before upgrading.

### 2. The `doc_*` semantic surface

Seventeen tools designed for an agent rather than for a script. They share one addressing
model, one report shape, and a dry-run mode.

**Locators.** A locator is a dict naming exactly one of five forms:

| Form | Example |
|---|---|
| `paragraph` | `{"paragraph": 12}` |
| `find` | `{"find": "ABC Corporation", "occurrence": 2}` |
| `bookmark` | `{"bookmark": "signature_block"}` |
| `heading` | `{"heading": "Schedule A"}` |
| `table` | `{"table": 0, "row": 1, "col": 2, "paragraph": 0}` |

Two optional keys work on every form. `story` selects a header, a footer, the footnotes
part and so on (`{"find": "Confidential", "story": "header1"}`); it defaults to the main
document. `expect_text` is what the caller believes is there — an exact text or a prefix.
If it does not match, the call fails as `stale_anchor` rather than editing the wrong
paragraph, and the error lists the paragraphs that *do* carry that text so the caller can
re-aim without re-reading the document.

**Report shape.** Every call that returns answers with at least four keys:

```json
{ "status": "ok", "dry_run": false, "changes": [], "warnings": [] }
```

`changes` holds one entry per paragraph touched — `{story, paragraph, before, after}` —
and tools add their own keys on top (`matches`, `saved`, `runs`, …). A failure answers
with a different, equally uniform shape:

```json
{ "status": "error", "code": "stale_anchor", "message": "…" }
```

**The tools.**

- Reading: `doc_inspect`, `doc_find`, `doc_get_effective_format`, `doc_list_styles`,
  `doc_get_style`, `doc_find_style_usage`, `doc_capabilities`, `doc_audit`, `doc_compare`
- Writing: `doc_edit_text`, `doc_format_range`, `doc_apply_list`, `doc_apply_table_style`,
  `doc_create_style`, `doc_update_style`, `doc_delete_style`
- Batching: `doc_apply_edits` applies a sequence of text and formatting edits as one
  all-or-nothing unit — either the whole batch lands or the file is left as it was.

`doc_audit` reads a document and reports what is wrong with it in one read-only pass:
paragraphs that read like a heading but carry no heading style, direct formatting
overriding a style, dangling `numId`/`abstractNumId` references, bookmark and comment
ranges missing one end, unused custom styles, near-duplicate style names, who left which
tracked changes, and more — see its docstring for the full list of checks. Every finding
carries the locator of the place it is about. `doc_compare` reads two
documents and reports the structural difference between them — added, removed and
changed paragraphs and tables, section and counter deltas — so a caller can confirm that
a batch of edits touched exactly what it meant to and nothing else.

`doc_capabilities` reports what *this* process on *this* machine can reach: the platform,
whether LibreOffice is on `PATH` (the `convert_to_pdf` fallback outside Windows), and the
tool family counts. It does not launch Word, so it cannot tell you whether Word is
installed — only a live tool call can.

### 3. Live tools (Word automation)

`word_live_*` and `word_screen_capture` drive a Word application that is already running,
so edits appear in the open document as you watch. On Windows every destructive live tool
is wrapped in an undo record, which makes each call a single Ctrl+Z. They are the one
family that is not cross-platform: COM on Windows, JavaScript for Automation on macOS. A
tool counts as macOS-capable only when its macOS branch reaches a real implementation, not
a stub — the exact split is in [TOOLS.md](TOOLS.md).

<!-- tool-counts:start -->
**137 tools** across two modes — see the [complete tool reference](TOOLS.md) for details.

| Category | Count |
|----------|-------|
| Cross-platform (python-docx) | 75 |
| V2 semantic engine (`doc_*`, python-docx) | 17 |
| Windows Live (COM automation) | 45 |
| macOS Live (JXA automation) | 31 (of the 45 live tools) |
<!-- tool-counts:end -->

## A `doc_*` call, end to end

Given a document whose paragraph 1 reads
`This agreement is entered into by ABC Corporation and the client, effective 1 January 2026.`
with `ABC Corporation` in bold.

**Find it first.** `doc_find` with `{"filename": "agreement.docx", "pattern": "ABC Corporation"}`
(it also takes `regex`, `case`, `whole_word`, `stories` and `max_results`):

```json
{
  "status": "ok",
  "dry_run": false,
  "changes": [],
  "warnings": [],
  "matches": [
    {
      "story": "document",
      "paragraph": 1,
      "start": 34,
      "end": 49,
      "text": "ABC Corporation",
      "context": "This agreement is entered into by ABC Corporation and the client, effective 1 January 202"
    },
    {
      "story": "document",
      "paragraph": 2,
      "start": 0,
      "end": 15,
      "text": "ABC Corporation",
      "context": "ABC Corporation shall deliver the services described in"
    }
  ],
  "truncated": false
}
```

**Then rehearse the edit.** `doc_edit_text` with
`{"filename": "agreement.docx", "locator": {"find": "ABC Corporation", "occurrence": 1}, "action": "replace", "text": "XYZ Ltd", "dry_run": true}`:

```json
{
  "status": "ok",
  "dry_run": true,
  "changes": [
    {
      "story": "document",
      "paragraph": 1,
      "before": "This agreement is entered into by ABC Corporation and the client, effective 1 January 2026.",
      "after": "This agreement is entered into by XYZ Ltd and the client, effective 1 January 2026."
    }
  ],
  "warnings": [],
  "saved": false
}
```

Nothing was written (`saved: false`). Dropping `dry_run` performs the same edit for real —
and the replacement keeps the bold, because `ranges` edits the runs in place rather than
rebuilding them.

**And when the document has moved under you**, `expect_text` stops the call instead of
letting it land somewhere wrong:

```json
{
  "status": "error",
  "code": "stale_anchor",
  "message": "expected 'This agreement was signed by' at paragraph 1 of story 'document', found 'This agreement is entered into by ABC Co...'; no paragraph of that story carries that text"
}
```

**And once a batch of edits has landed**, `doc_audit` says what is left to fix and
`doc_compare` says what actually moved. `doc_audit` takes just `filename` and returns
findings — `heading_like_paragraph`, `unused_custom_style`, `dangling_num_id`, and more,
each carrying the locator to act on it. `doc_compare` takes two filenames and reports
added, removed and changed paragraphs and tables, section and counter deltas, against a
copy taken before the batch started:

```json
{
  "identical": false,
  "paragraphs": {
    "added": [],
    "removed": [],
    "changed": [
      {
        "story": "document",
        "paragraph": 1,
        "fields": ["text"],
        "text_before": "This agreement is entered into by ABC Corporation and the client...",
        "text_after": "This agreement is entered into by XYZ Ltd and the client..."
      }
    ]
  },
  "tables": {"added": [], "removed": [], "changed": []},
  "counters_changed": []
}
```

`tests/e2e/test_agent_flows.py` chains a full loop of this kind — `doc_inspect`,
`doc_find`, a tracked-change edit, a comment, a character style, a numbered list, a table
style, `doc_audit`, `doc_compare` against an untouched copy of the fixture, and
`validate_package` — on the `combined` fixture, then reopens the result in LibreOffice
when it is installed, plus one scenario for each of the four ways `doc_edit_text` refuses
rather than guesses: a range that would cut a field in half, a stale anchor, a search with
more than one match and no `occurrence`, and a range that would land inside an existing
tracked deletion.

## Fidelity guarantees

Each of these is enforced by the engine and covered by tests under `tests/engine/` and
`tests/tools/`:

- **Nothing outside the range changes.** Text and formatting outside `[start, end)` come
  out as they went in, including on randomised ranges over the full fixture set.
- **Parts that were not targeted are identical.** Editing the body does not perturb the
  headers, the footnotes, the theme or the media.
- **Bookmarks and comment ranges survive.** Markers inside a deleted range are moved to
  where the range started, never dropped, so the annotations that framed the text still
  frame something.
- **Containers are never removed.** Only `w:r` elements are deleted, and only when the
  range covers them whole. An emptied `w:hyperlink` keeps its `r:id`, an emptied `w:ins`
  keeps the revision it stands for, an emptied `w:sdt` keeps the content control.
- **Fields are atomic.** A range that strictly contains a field character, an image, an
  embedded object or a note reference is refused (`opaque-content`) rather than approximated.
- **`w:del` is never rewritten.** Tracked deletions are structure to preserve; only
  `revisions` accepts or rejects them, and it refuses as a whole rather than applying in
  part.
- **Annotation ids are unique.** Every id is allocated against the whole package, not the
  part being edited, so a new bookmark never shadows an existing one.
- **Schema order is respected.** `w:rPr` and `w:pPr` are `xsd:sequence`: every property
  this engine writes is inserted at its ECMA-376 rank, whatever order the caller listed
  them in. Appending is never an option.
- **Saving is atomic.** The destination is either the previous file or the complete new
  one — never a truncated package.

### Known gaps

Deliberate, documented limits rather than bugs to discover:

- **Nested fields.** In a nested complex field, the inner cached result sits inside the
  outer instruction and is visible text. Both spans are reported so a range can refuse to
  cut either, but the two are not disentangled further.
- **The footnote and endnote tools use a different index space.** The five tools in
  `tools/footnote_tools.py` read `paragraph_index` against python-docx's direct body
  paragraphs, while every other public `paragraph_index` counts the paragraphs of a block
  content control (`w:sdt`) too. Composing `find_text_in_document` with
  `add_footnote_to_document` can therefore aim at the wrong paragraph in a document that
  holds a block `w:sdt` — insert a table of contents and this becomes reachable.
- **Theme tint and shade are a rendering hint.** `w:themeTint` / `w:themeShade` scale
  luminance in HSL; the arithmetic is floating point here and fixed point in Word, so a
  resolved channel may land one or two units from the value Word caches. It is good enough
  to tell an agent that a run is a dark blue, and it is **never written back** into the
  document — the reference and its cached value stay the document's own business.
- **Table styles are not resolved in `doc_get_effective_format`.** A table style's
  contribution depends on conditional formatting (`w:tblStylePr`), which the effective
  format report does not compute; the report says so rather than guessing.
- **Behaviour changes in the historical tools** are listed under `[Unreleased]` in
  [CHANGELOG.md](CHANGELOG.md).

## Test harness

Non-degradation is a claim that has to be measurable, so the test suite is built around
instruments rather than around example files:

- **Generated fixtures** (`tests/fixtures/builders.py`). Every fixture is built on demand
  from the python-docx template plus hand-written XML fragments, with frozen ZIP
  timestamps and literal ids, so `build_x() == build_x()` byte for byte. **No binary
  fixture is ever committed.**
- **Canonical snapshots** (`tests/support/snapshot.py`). Turns a package into a comparable
  value using `zipfile` and `lxml` only — never python-docx, so it stays usable to
  characterise python-docx itself. This is what lets a test state "nothing changed except
  paragraph 4" and have it verified.
- **Package validation** (`tests/support/package_check.py`). Checks relationships in both
  directions: a `.rels` entry whose target part is gone, and a part still carrying an `r:`
  reference whose relationship no longer exists.
- **Characterization tests** (`tests/characterization/`) pin the behaviour of the existing
  tools, including a no-op round trip, so a migration that changes something has to say so.
- **LibreOffice is optional and never a source of truth.** It produces `.docx` fixtures
  from committed `.fodt` sources and can confirm a document reopens without error. Tests
  that need it are marked `libreoffice` and **skip** when `soffice` is missing. No test
  requires Microsoft Word.

```bash
PATH="$HOME/.local/bin:$PATH" uv run pytest -q          # the suite
PATH="$HOME/.local/bin:$PATH" uv run pytest -q -m "not libreoffice"
bash scripts/check.sh                                   # sync + lint + tests
```

`scripts/gen_tools_md.py` generates [TOOLS.md](TOOLS.md) and the tool-count table above
from the live server; `--check` fails if either has drifted. Note that `ruff` is configured
to cover `word_document_server/engine/`, `tests/` and `scripts/` only — a green
`ruff check .` says nothing about the older `tools/`, `core/` and `utils/` packages.

## Installation

This server runs from a clone. Python 3.11 or later.

```bash
git clone https://github.com/ykarapazar/word-mcp-live.git
cd word-mcp-live
uv sync            # or: pip install -e .
```

`python-docx`, `lxml`, `fastmcp` and `msoffcrypto-tool` are installed for you. `pywin32`
and `Pillow` come along on Windows only, for the live tools. `uv` is typically installed
under `~/.local/bin`; prefix commands with `PATH="$HOME/.local/bin:$PATH"` if that is not
on your `PATH`.

> The PyPI package named `word-mcp-live` is the upstream project this repository forked
> from. It does not contain the OOXML engine or the `doc_*` tools described here. Install
> from this clone.

### MCP client configuration

`pyproject.toml` declares the console entry point `word_mcp_server`. Point your client at
it through `uv`, with the clone as the working directory:

```json
{
  "mcpServers": {
    "word": {
      "command": "uv",
      "args": ["run", "--directory", "/absolute/path/to/word-mcp-live", "word_mcp_server"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

This is the form used by Claude Desktop (`claude_desktop_config.json`), Claude Code
(`.mcp.json`), Cursor (`~/.cursor/mcp.json`) and Windsurf; VS Code nests the same object
under `"mcp": {"servers": {…}}`.

### Environment variables

| Variable | Default | Effect |
|---|---|---|
| `MCP_AUTHOR` | `"Author"` | Author recorded on tracked changes and comments (`defaults.py`). |
| `MCP_AUTHOR_INITIALS` | `""` | Initials recorded on comments (`defaults.py`). |
| `MCP_TRANSPORT` | `stdio` | `stdio`, `sse` or `streamable-http`; an unknown value falls back to `stdio` with a warning. |
| `MCP_HOST` | `0.0.0.0` | Bind address, for the HTTP transports. |
| `MCP_PORT` | `8000` | Bind port, for the HTTP transports. |

For a hosted deployment, see [RENDER_DEPLOYMENT.md](RENDER_DEPLOYMENT.md).

## Documentation

| | |
|---|---|
| [TOOLS.md](TOOLS.md) | Every tool, with its platforms and whether it writes. Generated. |
| [CHANGELOG.md](CHANGELOG.md) | Release history, and the behaviour changes under `[Unreleased]`. |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Development setup, project layout, adding a tool. |
| [RENDER_DEPLOYMENT.md](RENDER_DEPLOYMENT.md) | Running the server over HTTP. |
| [PRIVACY.md](PRIVACY.md) | Privacy policy. |
| [LICENSE](LICENSE) | MIT. |

## Privacy

The server runs entirely on your machine and on the documents you point it at. Nothing is
collected, transmitted or stored elsewhere. See [PRIVACY.md](PRIVACY.md).

## Acknowledgments

Built on [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)
(MIT) and on [ykarapazar/word-mcp-live](https://github.com/ykarapazar/word-mcp-live), whose
live-editing tools this repository keeps.

Libraries: [python-docx](https://python-docx.readthedocs.io/) ·
[lxml](https://lxml.de/) · [FastMCP](https://gofastmcp.com/) ·
[pywin32](https://github.com/mhammond/pywin32)

## License

MIT — see [LICENSE](LICENSE).

<!-- mcp-name: io.github.ykarapazar/word-mcp-live -->
