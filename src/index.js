#!/usr/bin/env node

/**
 * docx-forge-mcp — MCP server for Word document (.docx) creation and manipulation.
 *
 * Tools:
 *   create_document   — Create a new .docx from markdown content
 *   read_document     — Extract text, headings, and metadata from a .docx
 *   add_section       — Append a heading + content section to an existing .docx
 *   replace_text      — Find and replace text in a .docx
 *   add_table         — Insert a table into an existing .docx
 *   merge_documents   — Combine multiple .docx files into one
 *   export_to_pdf     — Convert .docx to PDF via LibreOffice or pandoc
 *   get_document_stats — Word count, page estimate, section count, table count
 */

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import {
  createDocument,
  readDocument,
  addSection,
  replaceText,
  addTable,
  mergeDocuments,
  exportToPdf,
  getDocumentStats,
} from './document.js';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';


const __dirname = dirname(fileURLToPath(import.meta.url));
const pkg = JSON.parse(readFileSync(join(__dirname, '..', 'package.json'), 'utf8'));
const startTime = Date.now();
let toolCallCount = 0;

function wrap(fn) {
  return async (...args) => {
    toolCallCount++;
    try { return await fn(...args); }
    catch (e) { return { content: [{ type: 'text', text: JSON.stringify({ error: e.message }) }] }; }
  };
}
const server = new McpServer({
  name: 'docx-forge-mcp',
  version: pkg.version,
  description:
    'Production-grade Word document (.docx) creation and manipulation for AI agents — contracts, reports, proposals, compliance documents',
});

server.tool('health_check', 'Returns server health, uptime, version, and call stats', {},
  wrap(async () => ({
    content: [{ type: 'text', text: JSON.stringify({
      status: 'healthy', server: 'docx-forge-mcp', version: pkg.version,
      uptime_seconds: Math.floor((Date.now() - startTime) / 1000),
      tool_calls_served: toolCallCount,
    }, null, 2) }],
  }))
);

// ═══════════════════════════════════════════════════════
// TOOL: create_document
// ═══════════════════════════════════════════════════════

server.tool(
  'create_document',
  'Create a new .docx Word document from markdown or plain text content. Supports headings (#, ##, ###), bold (**text**), italic (*text*), bullet lists (- item), numbered lists (1. item), and blank lines between paragraphs.',
  {
    title: z.string().min(1).describe('Document title — appears as the document heading'),
    content: z.string().describe(
      'Document body content as markdown. Supports headings (#, ##, ###), bold (**text**), italic (*text*), bullet lists (- item), and numbered lists (1. item).'
    ),
    author: z.string().optional().describe('Author name stored in document metadata'),
    outputPath: z
      .string()
      .min(1)
      .describe('Absolute or relative path where the .docx file will be saved (e.g., /tmp/report.docx)'),
  },
  async (params) => {
    try {
      const result = await createDocument(params);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: true,
                filePath: result.filePath,
                paragraphCount: result.paragraphCount,
                message: `Document created successfully at ${result.filePath}`,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (err) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ success: false, error: err.message }, null, 2) }],
        isError: true,
      };
    }
  }
);

// ═══════════════════════════════════════════════════════
// TOOL: read_document
// ═══════════════════════════════════════════════════════

server.tool(
  'read_document',
  'Extract all text, paragraph list, headings, and file metadata from an existing .docx file. Returns structured data suitable for further processing or display.',
  {
    filePath: z
      .string()
      .min(1)
      .describe('Path to the .docx file to read'),
  },
  async (params) => {
    try {
      const result = await readDocument(params);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: true,
                text: result.text,
                paragraphs: result.paragraphs,
                headings: result.headings,
                metadata: result.metadata,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (err) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ success: false, error: err.message }, null, 2) }],
        isError: true,
      };
    }
  }
);

// ═══════════════════════════════════════════════════════
// TOOL: add_section
// ═══════════════════════════════════════════════════════

server.tool(
  'add_section',
  'Append a new section (heading + body content) to the end of an existing .docx document. The document is updated in place.',
  {
    filePath: z.string().min(1).describe('Path to the existing .docx file'),
    heading: z.string().min(1).describe('Section heading text'),
    content: z.string().describe('Section body content (markdown supported)'),
    headingLevel: z
      .number()
      .int()
      .min(1)
      .max(6)
      .default(2)
      .describe('Heading level: 1 = H1 (largest) through 6 = H6 (smallest). Default: 2'),
  },
  async (params) => {
    try {
      const result = await addSection(params);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: true,
                filePath: result.filePath,
                sectionsAdded: result.sectionsAdded,
                message: `Section "${params.heading}" added to ${result.filePath}`,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (err) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ success: false, error: err.message }, null, 2) }],
        isError: true,
      };
    }
  }
);

// ═══════════════════════════════════════════════════════
// TOOL: replace_text
// ═══════════════════════════════════════════════════════

server.tool(
  'replace_text',
  'Find and replace text in an existing .docx document. Useful for template substitution — e.g., replacing {{CLIENT_NAME}} with an actual name, or updating dates and contract values.',
  {
    filePath: z.string().min(1).describe('Path to the .docx file to modify'),
    find: z.string().min(1).describe('Text string to search for'),
    replace: z.string().describe('Replacement text (can be empty string to delete the found text)'),
    replaceAll: z
      .boolean()
      .default(true)
      .describe('If true (default), replace every occurrence. If false, replace only the first occurrence.'),
  },
  async (params) => {
    try {
      const result = await replaceText(params);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: true,
                filePath: result.filePath,
                replacementsCount: result.replacementsCount,
                message:
                  result.replacementsCount > 0
                    ? `Replaced ${result.replacementsCount} occurrence(s) of "${params.find}"`
                    : `Text "${params.find}" not found in document`,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (err) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ success: false, error: err.message }, null, 2) }],
        isError: true,
      };
    }
  }
);

// ═══════════════════════════════════════════════════════
// TOOL: add_table
// ═══════════════════════════════════════════════════════

server.tool(
  'add_table',
  'Insert a formatted table into an existing .docx document. The table is appended at the end of the document with bold header row and standard cell borders.',
  {
    filePath: z.string().min(1).describe('Path to the existing .docx file'),
    headers: z
      .array(z.string())
      .min(1)
      .describe('Column header labels (e.g., ["Name", "Role", "Department"])'),
    rows: z
      .array(z.array(z.string()))
      .describe(
        '2D array of table data. Each inner array is one row and must match the number of headers. Example: [["Alice", "Engineer", "Tech"], ["Bob", "Designer", "UX"]]'
      ),
  },
  async (params) => {
    try {
      // Validate row column counts match headers
      const colCount = params.headers.length;
      for (let i = 0; i < params.rows.length; i++) {
        if (params.rows[i].length !== colCount) {
          throw new Error(
            `Row ${i + 1} has ${params.rows[i].length} columns but headers define ${colCount} columns`
          );
        }
      }

      const result = await addTable(params);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: true,
                filePath: result.filePath,
                tableRows: result.tableRows,
                tableColumns: result.tableColumns,
                message: `Table inserted: ${result.tableColumns} columns × ${result.tableRows} data rows`,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (err) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ success: false, error: err.message }, null, 2) }],
        isError: true,
      };
    }
  }
);

// ═══════════════════════════════════════════════════════
// TOOL: merge_documents
// ═══════════════════════════════════════════════════════

server.tool(
  'merge_documents',
  'Combine multiple .docx files into a single document. Each source document is separated by a horizontal divider. Useful for assembling reports, consolidating contracts, or combining chapter files.',
  {
    filePaths: z
      .array(z.string())
      .min(2)
      .describe('Array of paths to .docx files to merge (minimum 2). Files are merged in the order given.'),
    outputPath: z
      .string()
      .min(1)
      .describe('Path where the merged .docx file will be saved'),
  },
  async (params) => {
    try {
      const result = await mergeDocuments(params);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: true,
                filePath: result.filePath,
                sourceCount: result.sourceCount,
                totalParagraphs: result.totalParagraphs,
                message: `Merged ${result.sourceCount} documents into ${result.filePath}`,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (err) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ success: false, error: err.message }, null, 2) }],
        isError: true,
      };
    }
  }
);

// ═══════════════════════════════════════════════════════
// TOOL: export_to_pdf
// ═══════════════════════════════════════════════════════

server.tool(
  'export_to_pdf',
  'Convert a .docx file to PDF. Attempts conversion via LibreOffice (best fidelity) then pandoc. Returns success status and the method used. Requires LibreOffice or pandoc to be installed on the system.',
  {
    filePath: z.string().min(1).describe('Path to the source .docx file'),
    outputPath: z.string().min(1).describe('Path where the output .pdf file will be saved'),
  },
  async (params) => {
    try {
      const result = await exportToPdf(params);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: result.success,
                filePath: result.filePath,
                method: result.method,
                message: result.message,
              },
              null,
              2
            ),
          },
        ],
        isError: !result.success,
      };
    } catch (err) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ success: false, error: err.message }, null, 2) }],
        isError: true,
      };
    }
  }
);

// ═══════════════════════════════════════════════════════
// TOOL: get_document_stats
// ═══════════════════════════════════════════════════════

server.tool(
  'get_document_stats',
  'Get statistics for a .docx document: word count, character count, estimated page count (at 250 words/page), number of sections (headings), and number of tables.',
  {
    filePath: z.string().min(1).describe('Path to the .docx file to analyze'),
  },
  async (params) => {
    try {
      const result = await getDocumentStats(params);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: true,
                ...result,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (err) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ success: false, error: err.message }, null, 2) }],
        isError: true,
      };
    }
  }
);

// ═══════════════════════════════════════════════════════
// RESOURCE: usage guide
// ═══════════════════════════════════════════════════════

server.resource(
  'usage-guide',
  'docx-forge://usage-guide',
  async () => ({
    contents: [
      {
        uri: 'docx-forge://usage-guide',
        mimeType: 'text/markdown',
        text: `# docx-forge-mcp Usage Guide

## Quick Start

### Create a document
Use \`create_document\` with a title and markdown content:
- Headings: \`# H1\`, \`## H2\`, \`### H3\`
- Bold: \`**text**\`
- Italic: \`*text*\`
- Bullet list: \`- item\`
- Numbered list: \`1. item\`

### Template substitution
1. Create document with placeholders: \`{{CLIENT_NAME}}\`, \`{{DATE}}\`
2. Use \`replace_text\` to substitute each placeholder with real values

### Build a report programmatically
1. \`create_document\` — scaffold with title and intro
2. \`add_section\` — append each section (Summary, Findings, Recommendations)
3. \`add_table\` — insert data tables
4. \`get_document_stats\` — verify length and structure
5. \`export_to_pdf\` — deliver final PDF (requires LibreOffice or pandoc)

### Merge documents
Combine multiple .docx files with \`merge_documents\`. Pass file paths in order.

## Workflow: Contract Generation
\`\`\`
create_document(title="Service Agreement", content="# Parties\\n\\n{{PARTY_A}}...")
replace_text(find="{{PARTY_A}}", replace="ACME Corp")
replace_text(find="{{PARTY_B}}", replace="Client LLC")
add_table(headers=["Service", "Price", "Timeline"], rows=[...])
export_to_pdf(filePath="contract.docx", outputPath="contract.pdf")
\`\`\`

## PDF Export
Requires system-level tools. Install with:
- LibreOffice: \`sudo apt-get install libreoffice\`
- pandoc: \`sudo apt-get install pandoc\`
LibreOffice produces higher-fidelity output.
`,
      },
    ],
  })
);

// ═══════════════════════════════════════════════════════
// Start server
// ═══════════════════════════════════════════════════════

const transport = new StdioServerTransport();
await server.connect(transport);
