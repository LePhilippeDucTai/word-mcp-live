#!/usr/bin/env bash
# The local equivalent of the `checks` job in .github/workflows/ci.yml. Keep the
# two in step: a contributor who passes this and then fails CI learns to stop
# trusting one of them.
set -euo pipefail

export PATH="$HOME/.local/bin:$PATH"

uv sync
uv run ruff check .
uv run python scripts/gen_tools_md.py --check
uv lock --check
uv build
uv run pytest tests/ -q
