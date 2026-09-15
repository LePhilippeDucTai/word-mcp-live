# Contributing to word-mcp-live

Thanks for your interest in contributing! This guide covers the basics of setting up a development environment and adding new tools.

## Development Setup

```bash
git clone https://github.com/ykarapazar/word-mcp-live.git
cd word-mcp-live
uv sync
```

`uv` must be on `PATH` (installed under `~/.local/bin` by default); prefix commands with
`PATH="$HOME/.local/bin:$PATH"` if it isn't.

For Windows Live tools, you also need:
- Windows 10/11 with Microsoft Word installed
- `pip install pywin32`

## Project Structure

```
word_document_server/
  main.py               # FastMCP server — legacy tools registered here; register_v2_tools() wires the doc_* ones
  defaults.py           # Default author name, initials
  engine/                # Pure OOXML engine: no MCP, no Word (see "Engine-backed tools" below)
    package.py           # DocxPackage: open, .stories(), atomic save
    textmodel.py          # What a paragraph's visible text is, and where each character sits
    ranges.py             # The only layer that writes text: split, delete, insert, replace
    format.py             # Direct run/paragraph formatting, in schema order
    find.py, locators.py, inspect.py, revisions.py, styles.py, theme.py, effective.py,
    numbering.py, table_styles.py, merge.py, audit.py, compare.py, errors.py, ids.py, xmlns.py
  tools/
    v2/                   # The doc_* tools, one module per family; registry.py documents the contract
    document_tools.py   # Document management (create, copy, info)
    content_tools.py    # Content manipulation (paragraphs, tables, images)
    format_tools.py     # Formatting (text, tables, styles)
    layout_tools.py     # Page layout, headers, watermarks
    footnote_tools.py   # Footnotes and endnotes
    protection_tools.py # Password protection, signatures
    comment_tools.py    # Read comments
    comment_write_tools.py  # Write comments
    hyperlink_tools.py  # Hyperlink management
    tracked_changes_tools.py  # Tracked changes via OOXML
    live_tools.py       # Windows COM editing tools
    live_read_tools.py  # Windows COM reading tools
    live_layout_tools.py    # Windows COM layout tools
    screen_capture_tools.py # Word window screenshot
  core/
    word_com.py         # COM helpers (get_word_app, undo_record)
```

## Code Style

- Python 3.11+ — use modern syntax (type hints, `match` statements where appropriate)
- Every tool function must have a docstring — it becomes the tool's description in MCP
- Use `description=impl_func.__doc__` in `@mcp.tool()` to keep a single source of truth
- Follow existing patterns in `main.py` for tool registration

## Adding a New Tool

### 1. Write the implementation

Add your function to the appropriate file in `word_document_server/tools/`. For a new cross-platform tool:

```python
# word_document_server/tools/content_tools.py

def my_new_tool(file_path: str, param: str) -> str:
    """Short description of what this tool does.

    Args:
        file_path: Path to the Word document
        param: What this parameter controls

    Returns:
        Success message or result
    """
    doc = Document(file_path)
    # ... implementation ...
    doc.save(file_path)
    return f"Done: {param}"
```

### 2. Register in main.py

```python
from word_document_server.tools import content_tools

@mcp.tool(
    annotations=ToolAnnotations(readOnlyHint=False),
    description=content_tools.my_new_tool.__doc__,
)
def my_new_tool(file_path: str, param: str) -> str:
    return content_tools.my_new_tool(file_path, param)
```

### 3. For Windows Live tools

Live tools use COM automation via `pywin32`. Use the helpers in `core/word_com.py`:

```python
from word_document_server.core.word_com import get_word_app, find_open_document, undo_record

def word_live_my_tool(filename: str = None) -> str:
    """Description of the live tool."""
    app = get_word_app()
    doc = find_open_document(app, filename)

    with undo_record(app, "MCP: My Tool"):
        # ... COM operations ...
        pass

    return "Done"
```

All destructive live tools must be wrapped with `undo_record` so each operation appears as a single Ctrl+Z entry in Word.

### 4. For engine-backed (`doc_*`) tools

The `doc_*` surface is not a thin wrapper around the legacy tools above: it is built
directly on `word_document_server/engine/`, and that engine follows conventions the
legacy code does not always honor.

- **Never mutate through `paragraph.runs`.** Clearing a paragraph's runs and rebuilding
  them from scratch (`for run in paragraph.runs: run.clear()`, then `paragraph.add_run(...)`)
  drops the formatting, fields, images and revisions the discarded runs carried — the
  exact defect this migration fixed in `format_text` and in `add_header_footer` (see
  `CHANGELOG.md`). Edit through `engine.ranges` (`resolve`, `delete_range`, `insert_text`,
  `replace_range`, `split_run`) instead: it only ever removes a `w:r` that a range covers
  whole, and never touches what it does not.
- **Never reconstruct a document with a fresh `Document()`.** Copying `paragraph.text`
  and a style name into a blank template — the pattern `add_table_of_contents` and
  `merge_documents` used before this migration — silently drops comments, footnotes,
  headers and footers, hyperlinks, bookmarks, fields and tracked changes, because none of
  those live on `paragraph.text`. Work on the package that is already open
  (`engine.package.DocxPackage`) and import or append elements directly.
- **Save through `DocxPackage.save`, not `doc.save(path)` on a half-built object.** It
  writes a temporary file in the destination's own directory, `fsync`s it, and only then
  `os.replace`s the target, so a failure mid-write leaves the original file intact
  instead of a truncated package. Do not write a `.docx` any other way.
- **Raise a typed error; do not return a success message that is not one.** Every
  deliberate failure the engine raises derives from `engine.errors.EngineError`
  (`PackageError`, `LocatorError(code)`, `UnsupportedRange(reason)`, `InvalidText`,
  `UnsupportedRevision`, `IdExhausted`), so a caller can branch on `.code` / `.reason`
  instead of parsing a string; `tools.v2.registry.error_code` is what turns one into the
  `code` a `doc_*` report answers with. A function that swallows a failure and reports
  success anyway is the `create_custom_style` defect this migration fixed (see
  `CHANGELOG.md`) — do not reintroduce it.
- **Register a new `doc_*` tool through `tools/v2`, not through `main.py`.** Add the
  function to the relevant module of `word_document_server/tools/v2/` (or a new one) and
  list it in that module's `TOOLS: list[ToolSpec]`. `register_v2_tools` (called once from
  `main.py`) imports every module of the package and registers what it finds there —
  `main.py` never names a `doc_*` tool, and a module without a `TOOLS` list is silently
  skipped. FastMCP derives the input schema from the function's signature and the
  description from its docstring; the registry wraps every call in the shared report
  envelope (`status`, `dry_run`, `changes`, `warnings`) and the error mapping above, so
  the function itself takes plain arguments, returns a plain `dict`, and raises on
  failure — it never builds the envelope or catches anything. See the module docstring of
  `word_document_server/tools/v2/registry.py` for the full contract.

## Running Tests

```bash
PATH="$HOME/.local/bin:$PATH" uv run pytest tests/ -q
```

Or run the full check (sync + lint + tests) with:

```bash
bash scripts/check.sh
```

Tests are marked `libreoffice` (requires a LibreOffice installation) and `characterization`
(pinning existing behavior) where relevant.

## Pull Request Guidelines

1. Keep PRs focused — one feature or fix per PR
2. After adding or removing a tool, run `uv run python scripts/gen_tools_md.py` — it regenerates `TOOLS.md` and the block between `<!-- tool-counts:start -->` and `<!-- tool-counts:end -->` in `README.md` from the registry — and never edit that block by hand: `uv run python scripts/gen_tools_md.py --check` and `tests/test_docs_sync.py` fail on a stale copy.
3. Add a changelog entry under `## [Unreleased]` in `CHANGELOG.md`
4. Test cross-platform tools on at least one platform; live tools require Windows + Word
