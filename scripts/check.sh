#!/usr/bin/env bash
set -euo pipefail

export PATH="$HOME/.local/bin:$PATH"

uv sync
uv run ruff check .
uv run pytest tests/ -q
