"""agent-docs — Document lifecycle suite for Hermes Agent.

Three modules:
  - create  — Generate .docx/.xlsx/.pptx/.pdf/.odt/.ods/.odp from structured JSON
  - convert — Format conversion via pandoc + LaTeX
  - read    — Extract text and metadata from any document via LiteParse
"""

from __future__ import annotations

import os
from pathlib import Path

__version__ = "1.0.0"

# ── Configuration ─────────────────────────────────────────────

def get_docs_venv_python() -> str:
    """Path to the Python binary in the docs venv.

    Override with DOCS_VENV_PYTHON env var. Defaults to
    ~/.venvs/docs/bin/python.
    """
    return os.environ.get(
        "DOCS_VENV_PYTHON",
        str(Path.home() / ".venvs" / "docs" / "bin" / "python"),
    )


def get_output_dir() -> Path:
    """Default output directory for generated documents.

    Override with DOCS_OUTPUT_DIR env var. Defaults to ~/Documents.
    """
    d = os.environ.get("DOCS_OUTPUT_DIR")
    if d:
        p = Path(d).expanduser()
    else:
        p = Path.home() / "Documents"
    p.mkdir(parents=True, exist_ok=True)
    return p


__all__ = ["get_docs_venv_python", "get_output_dir"]