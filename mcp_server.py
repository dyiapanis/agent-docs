#!/usr/bin/env python3
"""agent-docs MCP server — stdio Model Context Protocol server.

Exposes three tools:
  - doc_read:       Read/extract text from any document
  - doc_convert:    Convert between formats via pandoc
  - doc_create:     Create documents from structured JSON or raw HTML/CSS

Run standalone for testing:
    python3 mcp_server.py

Or wire into an MCP client via stdio.
"""

from __future__ import annotations

import json
import sys
from typing import Any

# ── MCP protocol ─────────────────────────────────────────────────────────────

def _read_message() -> dict | None:
    """Read one JSON-RPC message from stdin (Content-Length framed)."""
    headers = {}
    while True:
        line = sys.stdin.readline()
        if not line:
            return None
        line = line.strip()
        if not line:
            break
        if ":" in line:
            key, val = line.split(":", 1)
            headers[key.strip().lower()] = val.strip()
    length = int(headers.get("content-length", 0))
    if length == 0:
        return None
    body = sys.stdin.read(length)
    try:
        return json.loads(body)
    except json.JSONDecodeError:
        return None


def _write_message(msg: dict) -> None:
    """Write a JSON-RPC message to stdout with Content-Length framing."""
    body = json.dumps(msg)
    sys.stdout.write(f"Content-Length: {len(body)}\r\n\r\n{body}")
    sys.stdout.flush()


def _error(req_id: Any, code: int, message: str) -> dict:
    return {"jsonrpc": "2.0", "id": req_id, "error": {"code": code, "message": message}}


def _result(req_id: Any, result: Any) -> dict:
    return {"jsonrpc": "2.0", "id": req_id, "result": result}


# ── Tool definitions ─────────────────────────────────────────────────────────

TOOLS = [
    {
        "name": "doc_read",
        "description": (
            "Read and extract text from documents. Uses anydoc for office formats "
            "(DOCX, XLSX, PPTX, ODT, RTF, EPUB, PDF → Markdown) and LiteParse for OCR "
            "(scanned PDFs, images). Plain text files read directly."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {
                    "type": "string",
                    "description": "Absolute path to the document.",
                },
                "max_chars": {
                    "type": "integer",
                    "description": "Maximum characters to return.",
                    "default": 1000000,
                },
                "extract_images": {
                    "type": "boolean",
                    "description": "Extract text from images via OCR (LiteParse + Tesseract).",
                    "default": False,
                },
            },
            "required": ["path"],
        },
    },
    {
        "name": "doc_read_meta",
        "description": "Read document metadata (PDF info, image EXIF, etc.) via LiteParse.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {
                    "type": "string",
                    "description": "Absolute path to the document.",
                },
            },
            "required": ["path"],
        },
    },
    {
        "name": "doc_convert",
        "description": (
            "Convert documents between formats (PDF, DOCX, HTML, ODT, EPUB, etc.) "
            "using pandoc + LaTeX. Requires system pandoc and texlive-xetex for PDF output."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {
                    "type": "string",
                    "description": "Absolute path to the source document.",
                },
                "to_format": {
                    "type": "string",
                    "description": "Target format: pdf, docx, html, odt, epub, etc.",
                },
                "from_format": {
                    "type": "string",
                    "description": "Optional source format hint (e.g. markdown, html). Pandoc infers from extension if omitted.",
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional absolute output path. Defaults to the documents directory (DOCS_OUTPUT_DIR or ~/Documents).",
                },
                "extra_args": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Optional extra pandoc CLI flags.",
                },
                "category": {
                    "type": "string",
                    "description": "Optional subfolder under documents/ (e.g. 'reports').",
                },
            },
            "required": ["input_path", "to_format"],
        },
    },
    {
        "name": "doc_create_document",
        "description": (
            "Create .docx, .xlsx, .pptx, .pdf, .odt, .ods, or .odp documents. "
            "For PDFs: pass structured JSON for document-style output, or pass "
            '{\"html\": \"<html>...</html>\"} with full CSS (Grid, Flexbox, gradients, '
            "themes) for visual/print-ready PDFs via weasyprint. "
            "Supports page_size, header, footer, page_numbers options."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "format": {
                    "type": "string",
                    "enum": ["docx", "xlsx", "pptx", "pdf", "odt", "ods", "odp"],
                    "description": "Output document format.",
                },
                "content": {
                    "anyOf": [
                        {"type": "string", "description": "Plain text or HTML content (auto-detected and normalized)."},
                        {"type": "object", "description": "Structured content dict. Format-specific schema."},
                    ],
                    "description": "Document content. Accepts plain text, HTML, or structured dict.",
                },
                "output_path": {
                    "type": "string",
                    "description": "Absolute path for the output file. Defaults to the documents directory.",
                },
                "template": {
                    "type": "string",
                    "description": "Optional template path (.dotx, .odt, .xltx, .potx, etc.)",
                },
                "options": {
                    "type": "object",
                    "description": "Optional format-specific options (title, page_size, auto_size, etc.)",
                },
            },
            "required": ["format", "content"],
        },
    },
    {
        "name": "doc_edit_pdf",
        "description": "Edit a PDF page using search-and-replace via PyMuPDF. Provide a natural language instruction like 'change <search> to <replace>'.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {
                    "type": "string",
                    "description": "Absolute path to the PDF file.",
                },
                "page": {
                    "type": "integer",
                    "description": "Page number to edit (1-indexed).",
                    "default": 1,
                },
                "instruction": {
                    "type": "string",
                    "description": "Natural language instruction for the edit (e.g. 'change \"foo\" to \"bar\"').",
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional output path. Defaults to <original>_edited.pdf",
                },
            },
            "required": ["path", "instruction"],
        },
    },
]

# ── Tool dispatch ────────────────────────────────────────────────────────────

def _handle_tool(name: str, args: dict) -> str:
    """Dispatch a tool call to the appropriate handler and return JSON result."""
    try:
        if name == "doc_read":
            from agent_docs.read import read
            text = read(
                args["path"],
                max_chars=args.get("max_chars", 1_000_000),
                extract_images=args.get("extract_images", False),
            )
            return json.dumps({
                "success": True,
                "path": args["path"],
                "text": text,
                "length": len(text),
                "message": f"Read {len(text)} chars from {args['path']}",
            })

        elif name == "doc_read_meta":
            from agent_docs.read import read_meta
            meta = read_meta(args["path"])
            return json.dumps({"success": True, "path": args["path"], "metadata": meta})

        elif name == "doc_convert":
            from agent_docs.convert import doc_convert
            result = doc_convert(
                input_path=args["input_path"],
                to_format=args["to_format"],
                from_format=args.get("from_format"),
                output_path=args.get("output_path"),
                extra_args=args.get("extra_args"),
                category=args.get("category"),
            )
            return json.dumps(result)

        elif name == "doc_create_document":
            from agent_docs.create import create_document
            output_path = args.get("output_path")
            if not output_path:
                from agent_docs import get_output_dir
                import time
                ts = int(time.time())
                fmt = args["format"]
                output_path = str(get_output_dir() / f"doc_{ts}.{fmt}")
            result = create_document(
                format=args["format"],
                content=args["content"],
                output_path=output_path,
                template=args.get("template"),
                options=args.get("options"),
            )
            return json.dumps(result)

        elif name == "doc_edit_pdf":
            from agent_docs.create import _edit_pdf_text
            result = _edit_pdf_text(
                path=args["path"],
                page=args.get("page", 1),
                instruction=args["instruction"],
                output_path=args.get("output_path"),
            )
            return json.dumps(result)

        else:
            return json.dumps({"success": False, "error": f"Unknown tool: {name}"})

    except Exception as e:
        return json.dumps({"success": False, "error": f"{type(e).__name__}: {e}"})


# ── Request handlers ─────────────────────────────────────────────────────────

def _handle_request(msg: dict) -> dict | None:
    method = msg.get("method", "")
    req_id = msg.get("id")
    params = msg.get("params", {})

    if method == "initialize":
        return _result(req_id, {
            "protocolVersion": "2024-11-05",
            "capabilities": {"tools": {}},
            "serverInfo": {"name": "agent-docs", "version": "2.0.0"},
        })

    elif method == "notifications/initialized":
        return None  # notification — no response

    elif method == "tools/list":
        return _result(req_id, {"tools": TOOLS})

    elif method == "tools/call":
        tool_name = params.get("name")
        tool_args = params.get("arguments", {})
        result_str = _handle_tool(tool_name, tool_args)
        return _result(req_id, {
            "content": [{"type": "text", "text": result_str}],
        })

    elif method == "ping":
        return _result(req_id, {})

    else:
        if req_id is not None:
            return _error(req_id, -32601, f"Method not found: {method}")
        return None


# ── Main loop ────────────────────────────────────────────────────────────────

def main():
    while True:
        msg = _read_message()
        if msg is None:
            break
        response = _handle_request(msg)
        if response is not None:
            _write_message(response)


if __name__ == "__main__":
    main()