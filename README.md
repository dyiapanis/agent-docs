# agent-docs

A portable document lifecycle suite: create, convert, and read documents
from structured JSON or raw HTML/CSS.

## Cross-Platform Compatibility

This package conforms to the [Agent Plugins v1.0.0](https://agent-plugins.org/specification)
specification. It works with any compatible AI agent client that supports
MCP servers, including:

- [Hermes Agent](https://hermes-agent.nousresearch.com)
- [Cursor](https://cursor.com)
- [Claude Desktop](https://claude.ai/download)
- Any client that supports the Model Context Protocol

### Two installation paths

**Portable plugin (v1 package)** — works with any v1-compatible client:

```bash
# Hermes Agent
hermes plugins install dyiapanis/agent-docs

# Other clients: follow your client's plugin installation instructions
```

**pip install** — use the Python library directly or with Hermes native plugins:

```bash
pip install hermes-docs[all]
```

Or install only the modules you need:

```bash
pip install hermes-docs[create]    # python-docx, openpyxl, python-pptx, weasyprint, odfpy
pip install hermes-docs[read]      # firecrawl-anydoc, liteparse, PyMuPDF
pip install hermes-docs[convert]   # pypandoc (also needs system pandoc)
```

## Modules

### Create

Generate documents from structured JSON content or raw HTML/CSS:

- **DOCX** — Word documents (python-docx)
- **XLSX** — Spreadsheets (openpyxl)
- **PPTX** — Presentations (python-pptx)
- **PDF** — Print-ready documents and visual designs (WeasyPrint)
  - Document-style: structured JSON → HTML → PDF
  - Visual-style: raw HTML/CSS (Grid, Flexbox, gradients, themes) → PDF
- **ODT/ODS/ODP** — OpenDocument formats (odfpy)

Accepts plain strings (auto-normalized), HTML, or structured dicts with
sections, sheets, slides, tables, images, and page breaks.

### Convert

Format conversion via pandoc + LaTeX:

- Markdown → PDF/DOCX/HTML/ODT/EPUB
- Any pandoc-supported format pair
- Configurable page size, margins, headers/footers

### Read

Extract text and metadata from any document:

- **Office formats** (DOCX, DOC, XLSX, XLSB, PPTX, PPT, ODT, ODS, ODP, RTF,
  EPUB, CSV) and **text-based PDFs** → GitHub-Flavored Markdown via
  [anydoc](https://github.com/firecrawl/anydoc) (~4ms, in-process)
- **Scanned PDFs** and **images** (JPG, PNG, TIFF) → OCR via
  [LiteParse](https://github.com/run-llama/liteparse) (Tesseract)
- **Plain text** (.txt, .md, .json, .py, .yaml, etc.) → direct file read
- Returns structured Markdown (office formats) or plain text (OCR, plain text)

## Prerequisites

### Create module

The create module uses a separate Python venv for heavy document libraries.
Set `DOCS_VENV_PYTHON` to point to it:

```bash
# Create a dedicated venv
python -m venv ~/.venvs/docs
~/.venvs/docs/bin/pip install python-docx openpyxl python-pptx weasyprint odfpy Pillow lxml PyMuPDF nano-pdf xlsxwriter
```

Or install the libraries into your main Python environment:

```bash
pip install hermes-docs[create]
```

### Read module

```bash
~/.venvs/docs/bin/pip install firecrawl-anydoc liteparse PyMuPDF
```

Or:

```bash
pip install hermes-docs[read]
```

### Convert module

```bash
pip install hermes-docs[convert]
```

Requires system-installed pandoc and a LaTeX engine:

```bash
# Debian/Ubuntu
sudo apt install pandoc texlive-xetex

# macOS
brew install pandoc basictex

# Arch
sudo pacman -S pandoc texlive
```

### WeasyPrint system dependencies (PDF creation)

```bash
# Debian/Ubuntu
sudo apt install libcairo2 libpango-1.0-0 libgdk-pixbuf-2.0-0

# macOS
brew install cairo pango gdk-pixbuf
```

## Installation from source

```bash
git clone https://github.com/dyiapanis/agent-docs.git
cd agent-docs
pip install -e ".[all]"
```

## Configuration

| Variable | Purpose | Default |
|----------|---------|---------|
| `DOCS_VENV_PYTHON` | Python binary for document libraries | System Python |
| `DOCS_OUTPUT_DIR` | Default output directory for generated documents | `~/Documents` |

## Usage

### Create a PDF from HTML/CSS

```python
from hermes_docs.create import create_document

result = create_document(
    format="pdf",
    content={"html": "<html><body><h1>Hello</h1></body></html>"},
    output_path="/tmp/output.pdf",
    options={"page_size": "A4"}
)
```

### Read a DOCX file

```python
from hermes_docs.read import read

text = read("/path/to/document.docx")
# Returns GitHub-Flavored Markdown
```

### Convert Markdown to PDF

```python
from hermes_docs.convert import doc_convert

result = doc_convert(
    input_path="/path/to/input.md",
    to_format="pdf",
    output_path="/path/to/output.pdf"
)
```

### Via MCP

The bundled MCP server exposes five tools: `doc_read`, `doc_read_meta`,
`doc_convert`, `doc_create_document`, `doc_edit_pdf`. Any MCP-compatible
client can call these directly.

## License

MIT