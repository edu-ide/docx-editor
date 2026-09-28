/**
 * A Word shape survives editing and saving its paragraph.
 *
 * Word writes lines and shapes inside mc:AlternateContent (a wps:wsp with a VML
 * fallback). The editor does not model them, and its default selective save
 * rewrites an edited paragraph from the model, so typing into the paragraph
 * that held a shape used to delete the shape without a warning. The shape is
 * now kept as parsed and written back unchanged.
 */

import { test, expect } from '@playwright/test';
import { EditorPage } from '../helpers/editor-page';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import JSZip from 'jszip';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/** Its paragraph 0FABF7F7 holds a rectangle Word wrapped in mc:AlternateContent. */
const FIXTURE = path.join(__dirname, '..', 'fixtures', 'issue-483-firstline-marker.docx');
const HOLDER_PARA_ID = '0FABF7F7';

/** The body paragraph with the given w14:paraId, as written. */
function paragraphXml(xml: string, paraId: string): string {
  const at = xml.indexOf(`w14:paraId="${paraId}"`);
  expect(at).toBeGreaterThan(0);
  const start = xml.lastIndexOf('<w:p ', at);
  const end = xml.indexOf('</w:p>', at);
  return xml.slice(start, end);
}

async function documentXml(bytes: Buffer): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('text');
}

test('typing into the paragraph that holds a Word shape keeps the shape on save', async ({
  page,
}) => {
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(FIXTURE);

  const original = await documentXml(fs.readFileSync(FIXTURE));
  const shape = paragraphXml(original, HOLDER_PARA_ID).match(
    /<mc:AlternateContent>[\s\S]*<\/mc:AlternateContent>/
  )?.[0];
  expect(shape).toBeTruthy();

  // Place the caret at the end of the holder paragraph's text and type.
  await page.locator('.layout-page-content').getByText('XXX-XXXXXXXXXX').first().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Z');

  const downloadPromise = page.waitForEvent('download');
  await editor.saveDocument();
  const download = await downloadPromise;
  const saved = await documentXml(fs.readFileSync((await download.path())!));

  const holder = paragraphXml(saved, HOLDER_PARA_ID);
  expect(holder).toContain(' Z');
  expect(holder).toContain(shape!);
});
