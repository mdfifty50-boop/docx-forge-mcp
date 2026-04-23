# docx-forge-mcp

**MCP server for Word document (.docx) creation and manipulation — the production-grade document automation tool for AI agents.**

Generate contracts, reports, proposals, and compliance documents directly from agent workflows. No Word installation required.

[![npm](https://img.shields.io/npm/v/docx-forge-mcp)](https://www.npmjs.com/package/docx-forge-mcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-blue)](https://modelcontextprotocol.io)

---

## Why docx-forge-mcp?

Word documents are the default format for contracts, compliance reports, legal agreements, and business proposals — especially in enterprise and government workflows. This MCP server gives AI agents the ability to produce and manipulate `.docx` files programmatically, without any Office installation, UI automation, or file format guesswork.

**Only 1 competitor exists** for Word document manipulation via MCP as of 2026. This server fills that gap with a production-quality, fully-tested implementation.

---

## Tools

| Tool | Description |
|------|-------------|
| `create_document` | Create a new `.docx` from a title and markdown content |
| `read_document` | Extract text, headings, paragraphs, and metadata from a `.docx` |
| `add_section` | Append a new section (heading + body) to an existing `.docx` |
| `replace_text` | Find and replace text — ideal for template variable substitution |
| `add_table` | Insert a formatted table with bold headers into a `.docx` |
| `merge_documents` | Combine multiple `.docx` files into one |
| `export_to_pdf` | Convert `.docx` to PDF via LibreOffice or pandoc |
| `get_document_stats` | Get word count, page estimate, section count, table count |

---

## Install

### Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "docx-forge": {
      "command": "npx",
      "args": ["docx-forge-mcp"]
    }
  }
}
```

### Cursor

Add to your `.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "docx-forge": {
      "command": "npx",
      "args": ["docx-forge-mcp"]
    }
  }
}
```

### Manual (Node.js)

```bash
npm install -g docx-forge-mcp
docx-forge-mcp
```

---

## Usage Examples

### Create a contract from a template

```
create_document(
  title="Service Agreement",
  content="# Parties\n\n**Client:** {{CLIENT_NAME}}\n**Provider:** Acme Corp\n\n## Scope of Work\n\n{{SCOPE}}\n\n## Payment\n\n{{PAYMENT_TERMS}}",
  outputPath="/tmp/contract.docx"
)

replace_text(filePath="/tmp/contract.docx", find="{{CLIENT_NAME}}", replace="TechCorp Ltd")
replace_text(filePath="/tmp/contract.docx", find="{{SCOPE}}", replace="Software development and consulting services.")
replace_text(filePath="/tmp/contract.docx", find="{{PAYMENT_TERMS}}", replace="Net 30 days. $15,000/month.")

export_to_pdf(filePath="/tmp/contract.docx", outputPath="/tmp/contract.pdf")
```

### Build a structured report

```
create_document(
  title="Q1 2026 Performance Report",
  content="## Executive Summary\n\nRevenue grew 34% YoY.",
  outputPath="/tmp/report.docx"
)

add_section(
  filePath="/tmp/report.docx",
  heading="Revenue Breakdown",
  content="Product A: $450K\nProduct B: $320K\nServices: $180K",
  headingLevel=2
)

add_table(
  filePath="/tmp/report.docx",
  headers=["Product", "Q1 Revenue", "Growth"],
  rows=[["Product A", "$450K", "+41%"], ["Product B", "$320K", "+28%"], ["Services", "$180K", "+19%"]]
)

get_document_stats(filePath="/tmp/report.docx")
```

### Read and inspect an existing document

```
read_document(filePath="/path/to/existing.docx")
# Returns: { text, paragraphs[], headings[], metadata: { wordCount, fileSizeBytes, ... } }
```

### Merge chapter files into a book

```
merge_documents(
  filePaths=["/docs/ch1.docx", "/docs/ch2.docx", "/docs/ch3.docx"],
  outputPath="/docs/complete-manual.docx"
)
```

---

## Markdown Support

`create_document` and `add_section` both accept markdown content:

| Markdown | Result |
|----------|--------|
| `# Heading 1` | H1 heading |
| `## Heading 2` | H2 heading |
| `**bold text**` | Bold text |
| `*italic text*` | Italic text |
| `- item` | Bullet list item |
| `1. item` | Numbered list item |
| `---` | Horizontal divider |
| Blank line | Paragraph break |

---

## PDF Export

`export_to_pdf` requires a system-level converter. It tries in order:

1. **LibreOffice** (best fidelity) — `sudo apt-get install libreoffice`
2. **pandoc** — `sudo apt-get install pandoc`

If neither is available, the tool returns `success: false` with installation instructions. The source `.docx` file is always preserved.

---

## Resources

| URI | Description |
|-----|-------------|
| `docx-forge://usage-guide` | Step-by-step guide with workflow examples |

---

## Dependencies

- [`docx`](https://www.npmjs.com/package/docx) — `.docx` creation (no Office required)
- [`mammoth`](https://www.npmjs.com/package/mammoth) — `.docx` reading and text extraction
- [`@modelcontextprotocol/sdk`](https://www.npmjs.com/package/@modelcontextprotocol/sdk) — MCP protocol
- [`zod`](https://www.npmjs.com/package/zod) — Input validation

---

## Requirements

- Node.js >= 18.0.0
- No Microsoft Word or Office installation required
- PDF export requires LibreOffice or pandoc (optional)

---

## Development

```bash
git clone https://github.com/mdfifty50-boop/docx-forge-mcp
cd docx-forge-mcp
npm install
npm test        # Run test suite
npm start       # Start MCP server (stdio)
```

---

## License

MIT — see [LICENSE](LICENSE)

---

## Related MCP Servers

- [agentic-observability-mcp](https://github.com/mdfifty50-boop/agent-observability-mcp) — Trace, cost-track, and monitor agent actions
- [agent-guard-mcp](https://github.com/mdfifty50-boop/agent-guard-mcp) — Loop detection and circuit breakers
- [compliance-shield-mcp](https://github.com/mdfifty50-boop/compliance-shield-mcp) — EU AI Act audit trails
