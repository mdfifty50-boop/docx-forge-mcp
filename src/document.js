/**
 * document.js — Core .docx creation and manipulation logic
 * Uses `docx` package for writing, `mammoth` for reading.
 */

import fs from 'node:fs';
import path from 'node:path';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
} from 'docx';
import mammoth from 'mammoth';

// ──────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────

/**
 * Map a heading level integer (1-6) to the docx HeadingLevel enum.
 * @param {number} level
 * @returns {string}
 */
function toHeadingLevel(level) {
  const map = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
    5: HeadingLevel.HEADING_5,
    6: HeadingLevel.HEADING_6,
  };
  return map[level] ?? HeadingLevel.HEADING_1;
}

/**
 * Ensure the output directory exists. Creates it recursively if not.
 * @param {string} filePath
 */
function ensureDir(filePath) {
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

/**
 * Assert that a file exists; throws a descriptive error if not.
 * @param {string} filePath
 */
function assertExists(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }
}

/**
 * Parse markdown text into an array of docx Paragraph objects.
 * Supports: headings (# to ######), bold (**text**), italic (*text*),
 * unordered list items (- / * / +), ordered list items (1.), horizontal rules (---),
 * and plain paragraphs. Blank lines become empty paragraphs.
 *
 * @param {string} markdown
 * @returns {Paragraph[]}
 */
function markdownToParagraphs(markdown) {
  const lines = markdown.split('\n');
  const paragraphs = [];

  for (const raw of lines) {
    const line = raw.trimEnd();

    // Heading
    const headingMatch = line.match(/^(#{1,6})\s+(.*)/);
    if (headingMatch) {
      const level = headingMatch[1].length;
      const text = headingMatch[2].trim();
      paragraphs.push(
        new Paragraph({
          heading: toHeadingLevel(level),
          children: parseInline(text),
        })
      );
      continue;
    }

    // Horizontal rule
    if (/^[-*_]{3,}$/.test(line.trim())) {
      paragraphs.push(new Paragraph({ text: '' }));
      continue;
    }

    // Unordered list item
    const ulMatch = line.match(/^(\s*)[-*+]\s+(.*)/);
    if (ulMatch) {
      paragraphs.push(
        new Paragraph({
          bullet: { level: Math.floor(ulMatch[1].length / 2) },
          children: parseInline(ulMatch[2]),
        })
      );
      continue;
    }

    // Ordered list item
    const olMatch = line.match(/^(\s*)\d+\.\s+(.*)/);
    if (olMatch) {
      paragraphs.push(
        new Paragraph({
          numbering: { reference: 'default-numbering', level: Math.floor(olMatch[1].length / 2) },
          children: parseInline(olMatch[2]),
        })
      );
      continue;
    }

    // Blank line → spacer paragraph
    if (line.trim() === '') {
      paragraphs.push(new Paragraph({ text: '' }));
      continue;
    }

    // Plain paragraph
    paragraphs.push(
      new Paragraph({ children: parseInline(line) })
    );
  }

  return paragraphs;
}

/**
 * Parse inline markdown (bold, italic, bold+italic, plain) into TextRun objects.
 * @param {string} text
 * @returns {TextRun[]}
 */
function parseInline(text) {
  const runs = [];
  // Match: ***bold+italic***, **bold**, *italic*, plain
  const pattern = /(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|([^*]+))/g;
  let match;
  while ((match = pattern.exec(text)) !== null) {
    if (match[2]) {
      runs.push(new TextRun({ text: match[2], bold: true, italics: true }));
    } else if (match[3]) {
      runs.push(new TextRun({ text: match[3], bold: true }));
    } else if (match[4]) {
      runs.push(new TextRun({ text: match[4], italics: true }));
    } else if (match[5]) {
      runs.push(new TextRun({ text: match[5] }));
    }
  }
  return runs.length > 0 ? runs : [new TextRun({ text })];
}

// ──────────────────────────────────────────────
// Tool implementations
// ──────────────────────────────────────────────

/**
 * Create a new .docx document from markdown or structured content.
 *
 * @param {object} params
 * @param {string} params.title
 * @param {string} params.content - Markdown string
 * @param {string} [params.author]
 * @param {string} params.outputPath
 * @returns {{ filePath: string, paragraphCount: number }}
 */
export async function createDocument({ title, content, author, outputPath }) {
  ensureDir(outputPath);

  const titleParagraph = new Paragraph({
    heading: HeadingLevel.TITLE,
    children: [new TextRun({ text: title, bold: true })],
  });

  const bodyParagraphs = markdownToParagraphs(content);

  const doc = new Document({
    creator: author ?? 'docx-forge-mcp',
    title,
    description: `Created by docx-forge-mcp`,
    sections: [
      {
        properties: {},
        children: [titleParagraph, ...bodyParagraphs],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);

  return {
    filePath: path.resolve(outputPath),
    paragraphCount: bodyParagraphs.length + 1,
  };
}

/**
 * Extract text, structure, and metadata from a .docx file.
 *
 * @param {object} params
 * @param {string} params.filePath
 * @returns {{ text: string, paragraphs: string[], headings: string[], metadata: object }}
 */
export async function readDocument({ filePath }) {
  assertExists(filePath);

  const buffer = fs.readFileSync(filePath);

  // Extract raw text
  const { value: rawText } = await mammoth.extractRawText({ buffer });

  // Extract HTML to identify headings
  const { value: html } = await mammoth.convertToHtml({ buffer });

  // Parse headings from HTML
  const headings = [];
  const headingRegex = /<h[1-6][^>]*>(.*?)<\/h[1-6]>/gi;
  let hMatch;
  while ((hMatch = headingRegex.exec(html)) !== null) {
    // Strip inner tags
    headings.push(hMatch[1].replace(/<[^>]+>/g, '').trim());
  }

  // Split text into non-empty paragraphs
  const paragraphs = rawText
    .split('\n')
    .map((p) => p.trim())
    .filter((p) => p.length > 0);

  // Basic metadata from file stats
  const stat = fs.statSync(filePath);
  const metadata = {
    filePath: path.resolve(filePath),
    fileSizeBytes: stat.size,
    lastModified: stat.mtime.toISOString(),
    paragraphCount: paragraphs.length,
    headingCount: headings.length,
    wordCount: rawText.split(/\s+/).filter((w) => w.length > 0).length,
  };

  return {
    text: rawText,
    paragraphs,
    headings,
    metadata,
  };
}

/**
 * Add a new section (heading + paragraphs) to an existing .docx file.
 * Appends to the end of the last section.
 *
 * @param {object} params
 * @param {string} params.filePath
 * @param {string} params.heading
 * @param {string} params.content - Markdown string
 * @param {number} [params.headingLevel]
 * @returns {{ filePath: string, sectionsAdded: number }}
 */
export async function addSection({ filePath, heading, content, headingLevel = 1 }) {
  assertExists(filePath);

  // Read existing document text via mammoth
  const existing = await readDocument({ filePath });

  // Rebuild full content: existing paragraphs + new section
  const existingText = existing.paragraphs.join('\n\n');
  const hashes = '#'.repeat(Math.min(Math.max(headingLevel, 1), 6));
  const newContent = `${hashes} ${heading}\n\n${content}`;
  const combined = existingText
    ? `${existingText}\n\n${newContent}`
    : newContent;

  // Use the existing file's title (first heading or filename)
  const docTitle = existing.headings[0] ?? path.basename(filePath, '.docx');

  await createDocument({
    title: docTitle,
    content: combined,
    outputPath: filePath,
  });

  return {
    filePath: path.resolve(filePath),
    sectionsAdded: 1,
  };
}

/**
 * Find and replace text in a .docx document.
 * Rebuilds the document with replaced content via raw text extraction.
 *
 * @param {object} params
 * @param {string} params.filePath
 * @param {string} params.find
 * @param {string} params.replace
 * @param {boolean} [params.replaceAll]
 * @returns {{ filePath: string, replacementsCount: number }}
 */
export async function replaceText({ filePath, find, replace, replaceAll = true }) {
  assertExists(filePath);

  const existing = await readDocument({ filePath });
  const docTitle = existing.headings[0] ?? path.basename(filePath, '.docx');

  let replacementsCount = 0;
  const originalText = existing.text;

  let modifiedText;
  if (replaceAll) {
    const regex = new RegExp(escapeRegex(find), 'g');
    modifiedText = originalText.replace(regex, () => {
      replacementsCount++;
      return replace;
    });
  } else {
    const idx = originalText.indexOf(find);
    if (idx !== -1) {
      modifiedText = originalText.slice(0, idx) + replace + originalText.slice(idx + find.length);
      replacementsCount = 1;
    } else {
      modifiedText = originalText;
    }
  }

  await createDocument({
    title: docTitle,
    content: modifiedText,
    outputPath: filePath,
  });

  return {
    filePath: path.resolve(filePath),
    replacementsCount,
  };
}

/**
 * Escape a string for use in a RegExp.
 * @param {string} str
 * @returns {string}
 */
function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Insert a table into an existing .docx document.
 * Appends the table at the end of the document.
 *
 * @param {object} params
 * @param {string} params.filePath
 * @param {string[]} params.headers
 * @param {string[][]} params.rows
 * @returns {{ filePath: string, tableRows: number, tableColumns: number }}
 */
export async function addTable({ filePath, headers, rows }) {
  assertExists(filePath);

  const existing = await readDocument({ filePath });
  const docTitle = existing.headings[0] ?? path.basename(filePath, '.docx');
  const existingText = existing.paragraphs.join('\n\n');

  const columnCount = headers.length;

  const headerRow = new TableRow({
    children: headers.map(
      (header) =>
        new TableCell({
          children: [
            new Paragraph({
              children: [new TextRun({ text: header, bold: true })],
              alignment: AlignmentType.CENTER,
            }),
          ],
          width: { size: Math.floor(9000 / columnCount), type: WidthType.DXA },
        })
    ),
    tableHeader: true,
  });

  const dataRows = rows.map(
    (row) =>
      new TableRow({
        children: row.map(
          (cell) =>
            new TableCell({
              children: [new Paragraph({ text: String(cell ?? '') })],
              width: { size: Math.floor(9000 / columnCount), type: WidthType.DXA },
            })
        ),
      })
  );

  const table = new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 9000, type: WidthType.DXA },
  });

  // Rebuild document with table appended
  const titleParagraph = new Paragraph({
    heading: HeadingLevel.TITLE,
    children: [new TextRun({ text: docTitle, bold: true })],
  });

  const bodyParagraphs = markdownToParagraphs(existingText);

  const doc = new Document({
    creator: 'docx-forge-mcp',
    title: docTitle,
    sections: [
      {
        properties: {},
        children: [titleParagraph, ...bodyParagraphs, new Paragraph({ text: '' }), table],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(filePath, buffer);

  return {
    filePath: path.resolve(filePath),
    tableRows: rows.length,
    tableColumns: columnCount,
  };
}

/**
 * Merge multiple .docx files into one combined document.
 *
 * @param {object} params
 * @param {string[]} params.filePaths
 * @param {string} params.outputPath
 * @returns {{ filePath: string, sourceCount: number, totalParagraphs: number }}
 */
export async function mergeDocuments({ filePaths, outputPath }) {
  if (!filePaths || filePaths.length === 0) {
    throw new Error('filePaths must be a non-empty array');
  }

  for (const fp of filePaths) {
    assertExists(fp);
  }

  ensureDir(outputPath);

  // Extract text from each document and concatenate
  const parts = [];
  for (const fp of filePaths) {
    const result = await readDocument({ filePath: fp });
    const docTitle = result.headings[0] ?? path.basename(fp, '.docx');
    parts.push(`# ${docTitle}\n\n${result.text.trim()}`);
  }

  const combinedContent = parts.join('\n\n---\n\n');
  const mergedTitle = 'Merged Document';

  const result = await createDocument({
    title: mergedTitle,
    content: combinedContent,
    outputPath,
  });

  return {
    filePath: result.filePath,
    sourceCount: filePaths.length,
    totalParagraphs: result.paragraphCount,
  };
}

/**
 * Export a .docx file to PDF.
 * Tries: libreoffice --headless, then pandoc. Falls back gracefully with instructions.
 *
 * @param {object} params
 * @param {string} params.filePath
 * @param {string} params.outputPath
 * @returns {{ filePath: string, method: string, success: boolean, message: string }}
 */
export async function exportToPdf({ filePath, outputPath }) {
  assertExists(filePath);
  ensureDir(outputPath);

  const { execFile } = await import('node:child_process');
  const { promisify } = await import('node:util');
  const execFileAsync = promisify(execFile);

  // Try LibreOffice first
  try {
    const outDir = path.dirname(outputPath);
    await execFileAsync('libreoffice', [
      '--headless',
      '--convert-to',
      'pdf',
      '--outdir',
      outDir,
      filePath,
    ]);

    // LibreOffice names the output after the source file
    const libreOutput = path.join(outDir, path.basename(filePath, '.docx') + '.pdf');
    if (fs.existsSync(libreOutput) && libreOutput !== outputPath) {
      fs.renameSync(libreOutput, outputPath);
    }

    if (fs.existsSync(outputPath)) {
      return {
        filePath: path.resolve(outputPath),
        method: 'libreoffice',
        success: true,
        message: 'Converted via LibreOffice headless',
      };
    }
  } catch {
    // LibreOffice not available — try pandoc
  }

  // Try pandoc
  try {
    await execFileAsync('pandoc', [filePath, '-o', outputPath]);
    if (fs.existsSync(outputPath)) {
      return {
        filePath: path.resolve(outputPath),
        method: 'pandoc',
        success: true,
        message: 'Converted via pandoc',
      };
    }
  } catch {
    // pandoc not available either
  }

  // Neither tool available — inform caller
  return {
    filePath: null,
    method: 'none',
    success: false,
    message:
      'PDF export requires LibreOffice or pandoc. Install with: ' +
      '`sudo apt-get install libreoffice` or `sudo apt-get install pandoc`. ' +
      'The source .docx file is intact at: ' + path.resolve(filePath),
  };
}

/**
 * Get word count, page estimate, section count, and table count for a .docx file.
 *
 * @param {object} params
 * @param {string} params.filePath
 * @returns {{ wordCount: number, pageEstimate: number, sectionCount: number, tableCount: number, characterCount: number }}
 */
export async function getDocumentStats({ filePath }) {
  assertExists(filePath);

  const buffer = fs.readFileSync(filePath);
  const { value: rawText } = await mammoth.extractRawText({ buffer });
  const { value: html } = await mammoth.convertToHtml({ buffer });

  const wordCount = rawText.split(/\s+/).filter((w) => w.length > 0).length;
  const characterCount = rawText.length;

  // Standard estimate: ~250 words per page
  const pageEstimate = Math.max(1, Math.round(wordCount / 250));

  // Count headings as sections
  const headingMatches = html.match(/<h[1-6][^>]*>/gi) ?? [];
  const sectionCount = headingMatches.length;

  // Count tables
  const tableMatches = html.match(/<table[^>]*>/gi) ?? [];
  const tableCount = tableMatches.length;

  return {
    wordCount,
    characterCount,
    pageEstimate,
    sectionCount,
    tableCount,
    filePath: path.resolve(filePath),
  };
}
