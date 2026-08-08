---
name: document-setup
description: Install and configure the hermes-docs document lifecycle suite — create, convert, and read documents.
---

# Document Setup

This skill covers installation and configuration of the hermes-docs MCP server,
which provides three document lifecycle tools:

- **doc_read** — Extract text from any document (office formats → Markdown, images → OCR)
- **doc_convert** — Convert between formats via pandoc + LaTeX
- **doc_create_document** — Generate .docx/.xlsx/.pptx/.pdf/.odt/.ods/.odp from structured JSON or HTML

## Prerequisites

### Python 3.11+

The MCP server and document libraries require Python 3.11 or later.

### System packages (for convert module)

```bash
# Debian/Ubuntu
sudo apt install pandoc texlive-xetex

# macOS
brew install pandoc basictex

# Arch
sudo pacman -S pandoc texlive
```

### System packages (for PDF creation via WeasyPrint)

```bash
# Debian/Ubuntu
sudo apt install libcairo2 libpango-1.0-0 libgdk-pixbuf-2.0-0

# macOS
brew install cairo pango gdk-pixbuf
```

## Installation

### Option A: pip install (recommended)

```bash
pip install hermes-docs[all]
```

Or install only the modules you need:

```bash
pip install hermes-docs[create]    # python-docx, openpyxl, python-pptx, weasyprint, odfpy
pip install hermes-docs[read]      # firecrawl-anydoc, liteparse, PyMuPDF
pip install hermes-docs[convert]   # pypandoc (also needs system pandoc)
```

### Option B: Dedicated venv (isolated from main environment)

```bash
python3 -m venv ~/.venvs/docs
~/.venvs/docs/bin/pip install hermes-docs[all]
```

Then set `DOCS_VENV_PYTHON` to point to the venv's Python:
```bash
export DOCS_VENV_PYTHON=~/.venvs/docs/bin/python
```

## Configuration

| Variable | Purpose | Default |
|----------|---------|---------|
| `DOCS_VENV_PYTHON` | Python binary for document libraries | System Python |
| `DOCS_OUTPUT_DIR` | Default output directory for generated documents | `~/Documents` |

## MCP Server Setup

The `mcp.json` in this package declares the MCP server. Clients that support
Agent Plugins v1 will auto-discover it. For manual MCP configuration:

```json
{
  "mcpServers": {
    "hermes-docs": {
      "command": "python3",
      "args": ["mcp_server.py"],
      "env": {}
    }
  }
}
```

## Verification

After setup, test each tool:

1. **Read**: `doc_read` on a .txt file → should return text content
2. **Convert**: `doc_convert` on a .md file to .pdf → should produce a PDF
3. **Create**: `doc_create_document` with format=docx and content="Hello" → should produce a .docx

## Troubleshooting

- **pandoc not found**: Install system pandoc package (see Prerequisites)
- **xelatex not found**: Install texlive-xetex (LaTeX engine for PDF conversion)
- **WeasyPrint import error**: Install cairo/pango/gdk-pixbuf system libraries
- **anydoc import error**: `pip install firecrawl-anydoc` in the docs venv
- **liteparse import error**: `pip install liteparse` (requires Tesseract for OCR)