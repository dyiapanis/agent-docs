# hermes-docs

A document lifecycle suite for [Hermes Agent](https://hermes-agent.nousresearch.com):
create, convert, and read documents from structured JSON or raw HTML/CSS.

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
# Create the docs venv
python -m venv ~/.hermes/venvs/docs
~/.hermes/venvs/docs/bin/pip install python-docx openpyxl python-pptx weasyprint odfpy Pillow lxml PyMuPDF nano-pdf xlsxwriter
```

Or install the libraries into the main Hermes venv:

```bash
pip install hermes-docs[create]
```

### Read module

```bash
~/.hermes/venvs/docs/bin/pip install firecrawl-anydoc liteparse PyMuPDF
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
sudo apt install pandoc texlive-xetex
```

## Installation

```bash
pip install hermes-docs[all]
```

Or from source:

```bash
git clone https://github.com/dyiapanis/agent-docs.git
cd hermes-docs
pip install -e ".[all]"
```

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
    src="/path/to/input.md",
    to_format="pdf",
    output_path="/path/to/output.pdf"
)
```

## License

MIT