import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { createDocument, readDocument, addSection, replaceText, addTable } from '../document.js';
import { mkdirSync, existsSync } from 'node:fs';

const TEST_DIR = '/tmp/mcp-test-docx-' + Date.now();
mkdirSync(TEST_DIR, { recursive: true });
const TEST_FILE = TEST_DIR + '/test.docx';

describe('docx-forge-mcp', () => {
  it('creates a document', async () => {
    const result = await createDocument({ title: 'Test Report', content: 'This is a test document.', author: 'Test', outputPath: TEST_FILE });
    assert.ok(result);
    assert.ok(existsSync(TEST_FILE) || result.path || result.created);
  });

  it('reads a document', async () => {
    const result = await readDocument({ filePath: TEST_FILE });
    assert.ok(result);
  });

  it('adds a section', async () => {
    const result = await addSection({ filePath: TEST_FILE, heading: 'New Section', content: 'Section content here', headingLevel: 2 });
    assert.ok(result);
  });

  it('replaces text', async () => {
    const result = await replaceText({ filePath: TEST_FILE, find: 'test', replace: 'production' });
    assert.ok(result);
  });

  it('adds a table', async () => {
    const result = await addTable({ filePath: TEST_FILE, headers: ['Name', 'Value'], rows: [['CPU', '90%'], ['Memory', '4GB']] });
    assert.ok(result);
  });
});
