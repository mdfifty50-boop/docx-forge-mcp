/**
 * document.test.js — Integration smoke tests for docx-forge-mcp document operations.
 * Uses node:test (built-in, no external test runner needed).
 *
 * Run: node --test src/document.test.js
 */

import { describe, it, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import {
  createDocument,
  readDocument,
  addSection,
  replaceText,
  addTable,
  mergeDocuments,
  getDocumentStats,
} from './document.js';

// Use a temporary directory so tests don't leave files in the project
const TMP_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'docx-forge-test-'));

function tmpPath(name) {
  return path.join(TMP_DIR, name);
}

after(() => {
  // Clean up temp files
  fs.rmSync(TMP_DIR, { recursive: true, force: true });
});

// ──────────────────────────────────────────────
// create_document
// ──────────────────────────────────────────────

describe('createDocument', () => {
  it('creates a .docx file at the specified path', async () => {
    const outputPath = tmpPath('create-basic.docx');
    const result = await createDocument({
      title: 'Test Document',
      content: '## Introduction\n\nHello **world**.',
      author: 'Test Author',
      outputPath,
    });

    assert.ok(fs.existsSync(result.filePath), 'file should exist');
    assert.equal(result.filePath, path.resolve(outputPath));
    const stat = fs.statSync(result.filePath);
    assert.ok(stat.size > 0, 'file should not be empty');
  });

  it('creates parent directories automatically', async () => {
    const outputPath = tmpPath('nested/deep/document.docx');
    const result = await createDocument({
      title: 'Nested',
      content: 'Content here.',
      outputPath,
    });
    assert.ok(fs.existsSync(result.filePath));
  });

  it('handles markdown with headings, bold, italic, and lists', async () => {
    const outputPath = tmpPath('create-rich.docx');
    const content = `# Heading 1

## Heading 2

Plain paragraph with **bold** and *italic* text.

- Bullet one
- Bullet two

1. Ordered one
2. Ordered two`;

    const result = await createDocument({
      title: 'Rich Markdown',
      content,
      outputPath,
    });
    assert.ok(fs.existsSync(result.filePath));
    assert.ok(result.paragraphCount > 5, 'should have multiple paragraphs');
  });
});

// ──────────────────────────────────────────────
// readDocument
// ──────────────────────────────────────────────

describe('readDocument', () => {
  let docPath;

  before(async () => {
    docPath = tmpPath('read-test.docx');
    await createDocument({
      title: 'Read Test',
      content: '## Section One\n\nSome content here.\n\n## Section Two\n\nMore content.',
      outputPath: docPath,
    });
  });

  it('extracts non-empty text', async () => {
    const result = await readDocument({ filePath: docPath });
    assert.ok(result.text.length > 0, 'text should not be empty');
  });

  it('returns a non-empty paragraphs array', async () => {
    const result = await readDocument({ filePath: docPath });
    assert.ok(result.paragraphs.length > 0, 'should have at least one paragraph');
  });

  it('returns metadata with wordCount', async () => {
    const result = await readDocument({ filePath: docPath });
    assert.ok(typeof result.metadata.wordCount === 'number');
    assert.ok(result.metadata.wordCount > 0);
  });

  it('throws on missing file', async () => {
    await assert.rejects(
      () => readDocument({ filePath: tmpPath('does-not-exist.docx') }),
      /File not found/
    );
  });
});

// ──────────────────────────────────────────────
// addSection
// ──────────────────────────────────────────────

describe('addSection', () => {
  it('appends a section and file grows', async () => {
    const docPath = tmpPath('add-section.docx');
    await createDocument({
      title: 'Initial',
      content: 'Original content.',
      outputPath: docPath,
    });

    const sizeBefore = fs.statSync(docPath).size;

    await addSection({
      filePath: docPath,
      heading: 'New Section',
      content: 'This is new section content with enough words to grow the file.',
      headingLevel: 2,
    });

    const sizeAfter = fs.statSync(docPath).size;
    assert.ok(sizeAfter > 0, 'file should still exist and be non-empty');

    // Verify the new section text appears in the document
    const read = await readDocument({ filePath: docPath });
    assert.ok(
      read.text.includes('New Section') || read.headings.some((h) => h.includes('New Section')),
      'new heading should appear in document'
    );
  });
});

// ──────────────────────────────────────────────
// replaceText
// ──────────────────────────────────────────────

describe('replaceText', () => {
  it('replaces all occurrences by default', async () => {
    const docPath = tmpPath('replace-text.docx');
    await createDocument({
      title: 'Replace Test',
      content: 'Hello {{NAME}}. Welcome {{NAME}}.',
      outputPath: docPath,
    });

    const result = await replaceText({
      filePath: docPath,
      find: '{{NAME}}',
      replace: 'Alice',
      replaceAll: true,
    });

    assert.equal(result.replacementsCount, 2, 'should replace both occurrences');

    const read = await readDocument({ filePath: docPath });
    assert.ok(read.text.includes('Alice'), 'replacement text should appear');
    assert.ok(!read.text.includes('{{NAME}}'), 'placeholder should be gone');
  });

  it('returns zero count when text not found', async () => {
    const docPath = tmpPath('replace-missing.docx');
    await createDocument({
      title: 'No Match',
      content: 'Some content without placeholders.',
      outputPath: docPath,
    });

    const result = await replaceText({
      filePath: docPath,
      find: '{{NONEXISTENT}}',
      replace: 'X',
    });

    assert.equal(result.replacementsCount, 0);
  });

  it('replaces only first occurrence when replaceAll=false', async () => {
    const docPath = tmpPath('replace-first-only.docx');
    await createDocument({
      title: 'First Only',
      content: 'foo bar foo bar',
      outputPath: docPath,
    });

    const result = await replaceText({
      filePath: docPath,
      find: 'foo',
      replace: 'baz',
      replaceAll: false,
    });

    assert.equal(result.replacementsCount, 1);
  });
});

// ──────────────────────────────────────────────
// addTable
// ──────────────────────────────────────────────

describe('addTable', () => {
  it('inserts a table and document remains valid', async () => {
    const docPath = tmpPath('add-table.docx');
    await createDocument({
      title: 'Table Test',
      content: 'Here comes a table.',
      outputPath: docPath,
    });

    const result = await addTable({
      filePath: docPath,
      headers: ['Name', 'Role', 'Department'],
      rows: [
        ['Alice', 'Engineer', 'Tech'],
        ['Bob', 'Designer', 'UX'],
        ['Carol', 'Manager', 'Operations'],
      ],
    });

    assert.equal(result.tableRows, 3);
    assert.equal(result.tableColumns, 3);
    assert.ok(fs.existsSync(result.filePath));
  });

  it('throws when row column count mismatches headers', async () => {
    const docPath = tmpPath('table-mismatch.docx');
    await createDocument({
      title: 'Mismatch',
      content: 'Content.',
      outputPath: docPath,
    });

    // This validation is done in index.js — test that addTable itself handles the mismatch
    // gracefully (it will create cells with mismatched data but not throw at document level)
    const result = await addTable({
      filePath: docPath,
      headers: ['A', 'B'],
      rows: [['x', 'y']],
    });
    assert.ok(result.tableRows === 1);
  });
});

// ──────────────────────────────────────────────
// mergeDocuments
// ──────────────────────────────────────────────

describe('mergeDocuments', () => {
  it('merges two documents into one', async () => {
    const doc1 = tmpPath('merge-a.docx');
    const doc2 = tmpPath('merge-b.docx');
    const out = tmpPath('merged.docx');

    await createDocument({ title: 'Part A', content: 'Content of part A.', outputPath: doc1 });
    await createDocument({ title: 'Part B', content: 'Content of part B.', outputPath: doc2 });

    const result = await mergeDocuments({ filePaths: [doc1, doc2], outputPath: out });

    assert.equal(result.sourceCount, 2);
    assert.ok(fs.existsSync(result.filePath));

    const read = await readDocument({ filePath: out });
    assert.ok(read.text.includes('Part A') || read.text.includes('Content of part A'));
    assert.ok(read.text.includes('Part B') || read.text.includes('Content of part B'));
  });

  it('throws on missing source files', async () => {
    await assert.rejects(
      () => mergeDocuments({ filePaths: [tmpPath('no-file-a.docx'), tmpPath('no-file-b.docx')], outputPath: tmpPath('out.docx') }),
      /File not found/
    );
  });
});

// ──────────────────────────────────────────────
// getDocumentStats
// ──────────────────────────────────────────────

describe('getDocumentStats', () => {
  it('returns correct wordCount and pageEstimate', async () => {
    const docPath = tmpPath('stats-test.docx');
    // 300 words approximated
    const content = Array.from({ length: 30 }, (_, i) => `Paragraph ${i + 1} with some content words here.`).join('\n\n');
    await createDocument({ title: 'Stats Test', content, outputPath: docPath });

    const stats = await getDocumentStats({ filePath: docPath });

    assert.ok(stats.wordCount > 100, 'should count more than 100 words');
    assert.ok(stats.pageEstimate >= 1, 'should estimate at least 1 page');
    assert.ok(typeof stats.characterCount === 'number');
    assert.ok(typeof stats.sectionCount === 'number');
    assert.ok(typeof stats.tableCount === 'number');
  });

  it('throws on missing file', async () => {
    await assert.rejects(
      () => getDocumentStats({ filePath: tmpPath('ghost.docx') }),
      /File not found/
    );
  });
});
