import { test, expect, devices } from '@playwright/test';
import type { BrowserContext, Page } from '@playwright/test';
import { TextReader, TextWriter, Uint8ArrayReader, Uint8ArrayWriter, ZipReader, ZipWriter } from '@zip.js/zip.js';
import { readFile, writeFile } from 'node:fs/promises';

const fixturePath = new URL('../demo/Editorial Plan.docx', import.meta.url);
const fixtureName = 'Editorial Plan.docx';
const EDITOR = '.rwe-pages';
const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAEAAAAAgCAIAAAAt/+nTAAAATklEQVR4nO3PUQkAIBTAwNfKTOazoCH8OITBAtxm7fN1wwUNaEEDWtCAFjSgBQ1oQQNa0IAWNKAFDWhBA1rQgBY0oAUNaEEDWtCAFjx2ATUjmJdZ6tDoAAAAAElFTkSuQmCC', 'base64');

const sheet = (page: Page) => page.locator(EDITOR);
const firstParagraph = (page: Page) => sheet(page).locator('[data-paragraph]').first();
const download = (page: Page) => page.getByRole('button', { name: 'Download', exact: true });

async function enableClipboard(context: BrowserContext) {
  test.skip(test.info().project.name !== 'chromium', 'System clipboard permission automation is available in Chromium.');
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
}

async function open(page: Page, bytes: Uint8Array, filename = fixtureName, discardChanges = true) {
  return page.evaluate(async ({ bytes, filename, discardChanges }) => {
    try {
      await window.editor!.open(new File([new Uint8Array(bytes)], filename), { discardChanges });
      return 'OK';
    } catch (cause) {
      return (cause as { code: string }).code;
    }
  }, { bytes: Array.from(bytes), filename, discardChanges });
}

async function exportBytes(page: Page) {
  const bytes = await page.evaluate(async () => Array.from(new Uint8Array(await (await window.editor!.exportDocx()).arrayBuffer())));
  return Buffer.from(bytes);
}

async function readZip(bytes: Uint8Array) {
  const zip = new ZipReader(new Uint8ArrayReader(bytes), { useWebWorkers: false });
  try {
    const entries = (await zip.getEntries()).filter(entry => !entry.directory);
    const parts = new Map<string, Uint8Array>();
    for (const entry of entries) parts.set(entry.filename, await entry.getData(new Uint8ArrayWriter()));
    return parts;
  } finally {
    await zip.close();
  }
}

async function patchFixture(update: (xml: string) => string, archivePath?: string) {
  const zip = new ZipReader(new Uint8ArrayReader(await readFile(fixturePath)), { useWebWorkers: false });
  const writer = new ZipWriter(new Uint8ArrayWriter(), { useWebWorkers: false });
  for (const entry of await zip.getEntries()) {
    if (entry.directory) continue;
    if (entry.filename === 'word/document.xml') {
      await writer.add(entry.filename, new TextReader(update(await entry.getData(new TextWriter()))));
    } else {
      await writer.add(entry.filename, new Uint8ArrayReader(await entry.getData(new Uint8ArrayWriter())));
    }
  }
  if (archivePath) await writer.add(archivePath, new TextReader('invalid entry'));
  await zip.close();
  return writer.close();
}

async function select(page: Page, from: string, fromOffset: number, to = from, toOffset?: number) {
  await expect(sheet(page).locator('[data-inline-start]').filter({ hasText: from }).first()).toBeAttached();
  if (to !== from) await expect(sheet(page).locator('[data-inline-start]').filter({ hasText: to }).first()).toBeAttached();
  await page.evaluate(({ from, fromOffset, to, toOffset }) => {
    const editor = document.querySelector<HTMLElement>('.rwe-pages')!;
    editor.focus();
    function point(text: string, offset?: number) {
      const element = Array.from(editor.querySelectorAll('[data-inline-start]')).find(node => node.textContent === text);
      if (!element?.firstChild) throw Error('Missing document run: ' + text);
      return [element.firstChild, offset ?? text.length] as const;
    }
    document.getSelection()!.setBaseAndExtent(...point(from, fromOffset), ...point(to, toOffset));
  }, { from, fromOffset, to, toOffset });
}

async function pasteHtml(page: Page, html: string, text = '') {
  await sheet(page).evaluate((element, content) => {
    const data = new DataTransfer();
    data.setData('text/html', content.html);
    data.setData('text/plain', content.text);
    element.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  }, { html, text });
}

test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    // Firefox intentionally ignores clipboardData supplied to a synthetic event.
    // Keep test payloads explicit without changing native browser clipboard events.
    const NativeClipboardEvent = window.ClipboardEvent;
    window.ClipboardEvent = class extends NativeClipboardEvent {
      constructor(type: string, init: ClipboardEventInit = {}) {
        super(type, init);
        if (init.clipboardData) Object.defineProperty(this, 'clipboardData', { value: init.clipboardData });
      }
    };
  });
  await page.goto('/tests/host.html');
  await expect(sheet(page)).toBeVisible();
});

test('opens a DOCX, keeps its untouched bytes, and round-trips formatting, table cells, links and a picture', async ({ page }, testInfo) => {
  const original = await readFile(fixturePath);
  expect(await open(page, original)).toBe('OK');
  await expect(sheet(page).locator('[data-page]')).toHaveCount(2);
  expect(await exportBytes(page)).toEqual(original);
  const originalParts = await readZip(original);

  await select(page, 'Overview', 0);
  await page.keyboard.press('Control+i');
  await expect(sheet(page).getByText('Overview', { exact: true })).toHaveCSS('font-style', 'italic');
  await select(page, 'Editorial', 0);
  await page.keyboard.type('Design');
  await expect(sheet(page).locator('td').getByText('Design', { exact: true })).toBeVisible();
  await select(page, 'Before sign-off', 0);
  await page.keyboard.press('Control+k');
  await page.getByLabel('Address').fill('https://example.com/review');
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await expect(sheet(page).locator('[data-href="https://example.com/review"]')).toHaveText('Before sign-off');
  await page.getByRole('tab', { name: 'Insert' }).click();
  const pick = page.waitForEvent('filechooser');
  await select(page, 'Review and handoff', 18, 'Review and handoff', 18);
  await page.getByRole('button', { name: 'Insert picture' }).click();
  await (await pick).setFiles({ name: 'example.png', mimeType: 'image/png', buffer: png });
  await expect(sheet(page).getByRole('img', { name: 'example' })).toBeVisible();

  const edited = await exportBytes(page);
  expect(await page.evaluate(() => window.editorState.dirty)).toBe(true);
  const parts = await readZip(edited);
  expect(Buffer.from(parts.get('word/_rels/document.xml.rels')!).toString()).toContain('https://example.com/review');
  for (const path of ['word/styles.xml', 'word/numbering.xml', '_rels/.rels']) {
    expect(Buffer.from(parts.get(path)!)).toEqual(Buffer.from(originalParts.get(path)!));
  }
  expect(Buffer.from(parts.get('word/media/image.png')!)).toEqual(png);
  expect(await open(page, edited)).toBe('OK');
  await expect(sheet(page).getByText('Overview', { exact: true })).toHaveCSS('font-style', 'italic');
  await expect(sheet(page).locator('td').getByText('Design', { exact: true })).toBeVisible();
  await expect(sheet(page).locator('[data-href="https://example.com/review"]')).toHaveText('Before sign-off');
  await expect(sheet(page).getByRole('img', { name: 'example' })).toBeVisible();
  await expect(sheet(page).locator('.rwe-list-label')).toHaveCount(3);

  const downloaded = page.waitForEvent('download');
  await download(page).click();
  await (await downloaded).saveAs(testInfo.outputPath('edited.docx'));
});

test('preserves protected XML and blocks edits that would cross it', async ({ page }) => {
  const protectedXml = '<w:sdt><w:sdtPr><w:tag w:val="keep-me"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Keep this text</w:t></w:r></w:p></w:sdtContent></w:sdt>';
  const original = await patchFixture(xml => xml.replace('<w:tbl>', protectedXml + '<w:tbl>'));
  expect(await open(page, original)).toBe('OK');
  await expect(sheet(page).getByText('Protected document content')).toBeVisible();
  await select(page, 'Production schedule', 0, 'Milestone');
  await page.keyboard.press('Backspace');
  expect(await exportBytes(page)).toEqual(Buffer.from(original));
  await select(page, 'Overview', 0);
  await page.keyboard.insertText('Summary');
  const output = await readZip(await exportBytes(page));
  expect(Buffer.from(output.get('word/document.xml')!).toString()).toContain(protectedXml);
});

test('opens structural tracked changes read-only and exports the original', async ({ page }) => {
  const original = await patchFixture(xml => xml.replace('<w:r><w:t xml:space="preserve">Overview</w:t></w:r>',
    '<w:moveTo w:id="12" w:author="Test"><w:r><w:t>Overview</w:t></w:r></w:moveTo>'));
  expect(await open(page, original)).toBe('OK');
  await expect(sheet(page)).toHaveAttribute('aria-readonly', 'true');
  await expect(page.getByRole('status').filter({ hasText: 'tracked changes' }).first()).toBeVisible();
  expect(await exportBytes(page)).toEqual(Buffer.from(original));
});

test('keeps unsaved edits after rejected replacement and malformed imports', async ({ page }) => {
  await firstParagraph(page).click();
  await page.keyboard.type('Unsaved work');
  const original = await readFile(fixturePath);
  expect(await open(page, original, fixtureName, false)).toBe('UNSAVED_CHANGES');
  expect(await open(page, Buffer.from('bad archive'))).toBe('INVALID_FILE');
  const malformedXml = await patchFixture(xml => xml.replace('</w:document>', ''));
  expect(await open(page, malformedXml)).toBe('INVALID_FILE');
  const traversal = await patchFixture(xml => xml, '../outside.xml');
  expect(await open(page, traversal)).toBe('INVALID_FILE');
  await expect(sheet(page)).toContainText('Unsaved work');
  await expect(page.locator('.rwe-local-state')).toHaveText('Unsaved changes');
  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'New document' }).click();
  await page.getByRole('button', { name: 'Cancel' }).click();
  await expect(sheet(page)).toContainText('Unsaved work');
  await page.getByRole('button', { name: 'New document' }).click();
  await page.getByRole('button', { name: 'Discard changes' }).click();
  await expect(sheet(page)).not.toContainText('Unsaved work');
  await expect(page.locator('.rwe-local-state')).toHaveText('Local document');
});

test('applies cross-page selection formatting and preserves it through undo and redo', async ({ page }) => {
  const original = await readFile(fixturePath);
  expect(await open(page, original)).toBe('OK');
  await select(page, 'Production schedule', 0, 'Review and handoff');
  await page.keyboard.press('Control+u');
  const before = sheet(page).locator('[data-page="1"]').getByText('Production schedule');
  const after = sheet(page).locator('[data-page="2"]').getByText('Review and handoff');
  await expect(before).toHaveCSS('text-decoration-line', 'underline');
  await expect(after).toHaveCSS('text-decoration-line', 'underline');
  await page.keyboard.press('Control+z');
  expect(await exportBytes(page)).toEqual(original);
  await page.keyboard.press('Control+Shift+z');
  await expect(before).toHaveCSS('text-decoration-line', 'underline');
  await expect(after).toHaveCSS('text-decoration-line', 'underline');
});

test('supports composing text, grouped undo, and inserting page breaks with the keyboard', async ({ page }) => {
  await firstParagraph(page).click();
  await page.keyboard.type('First page');
  await page.keyboard.press('Control+z');
  await expect(sheet(page)).not.toContainText('First page');
  await page.keyboard.press('Control+y');
  await expect(sheet(page)).toContainText('First page');
  await page.keyboard.press('Control+Enter');
  await page.keyboard.type('Second page ');
  if (test.info().project.name === 'chromium') {
    const cdp = await page.context().newCDPSession(page);
    await cdp.send('Input.imeSetComposition', { text: 'ni', selectionStart: 2, selectionEnd: 2 });
    await cdp.send('Input.imeSetComposition', { text: '你好', selectionStart: 2, selectionEnd: 2 });
    await cdp.send('Input.insertText', { text: '你好' });
  } else await page.keyboard.insertText('你好');
  await page.keyboard.type('!');
  await expect(sheet(page).locator('[data-page="2"]')).toContainText('Second page 你好!');
  await expect(sheet(page).locator('[data-page="1"]')).toContainText('First page');
});

test('uses a fit-to-width reading view on phones and collapses groups in narrow desktop panels', async ({ browser, page }) => {
  await page.setViewportSize({ width: 590, height: 800 });
  await expect(sheet(page)).toHaveAttribute('contenteditable', 'true');
  await expect(page.getByRole('button', { name: 'Align center' })).toBeHidden();
  await page.getByRole('button', { name: 'More formatting' }).click();
  await expect(page.getByRole('button', { name: 'Align center' })).toBeVisible();
  const context = await browser.newContext(devices['Pixel 7']);
  try {
    const phone = await context.newPage();
    await phone.goto('/');
    await expect(phone.locator('.rwe-document-title')).toHaveText('Editorial Plan');
    await expect(sheet(phone)).toHaveAttribute('aria-readonly', 'true');
    const paper = await sheet(phone).locator('[data-page="1"]').boundingBox();
    expect(paper!.width).toBeLessThan(412);
  } finally {
    await context.close();
  }
});

test('creates styles, numbering and a table in a new document and accepts plain-text paste', async ({ page }) => {
  await firstParagraph(page).click();
  await page.keyboard.type('Meeting notes');
  await page.getByLabel('Paragraph style').selectOption('Heading1');
  await expect(sheet(page).getByRole('heading', { name: 'Meeting notes' })).toBeVisible();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.getByLabel('Paste formatting').selectOption('text');
  await sheet(page).evaluate(element => {
    const data = new DataTransfer();
    data.setData('text/plain', 'First item\nSecond item');
    data.setData('text/html', '<b>Untrusted markup</b>');
    element.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  });
  await select(page, 'First item', 0, 'Second item');
  await page.getByRole('button', { name: 'Numbered list' }).click();
  await expect(sheet(page).locator('.rwe-list-label')).toHaveText(['1.', '2.']);
  await select(page, 'Second item', 0, 'Second item', 0);
  await page.getByRole('tab', { name: 'Insert' }).click();
  await page.getByRole('button', { name: 'Insert table' }).click();
  await page.getByLabel('Rows').fill('2');
  await page.getByLabel('Columns').fill('2');
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.keyboard.type('Cell value');
  await expect(sheet(page).getByRole('cell')).toHaveCount(4);
  await expect(sheet(page).getByRole('cell').first()).toHaveText('Cell value');
  const saved = page.waitForEvent('download');
  await download(page).click();
  const file = await saved;
  await expect(page.locator('.rwe-local-state')).toHaveText('Local document');
  expect(await open(page, await readFile(await file.path()), file.suggestedFilename())).toBe('OK');
  await expect(sheet(page).getByRole('heading', { name: 'Meeting notes' })).toBeVisible();
  await expect(sheet(page).locator('.rwe-list-label')).toHaveText(['1.', '2.']);
  await expect(sheet(page).getByRole('cell').first()).toHaveText('Cell value');
});

test('repaginates a long paragraph without losing text or selection', async ({ page }) => {
  await firstParagraph(page).click();
  const original = ('Text wraps and continues on the next page. '.repeat(190)).trim();
  await page.keyboard.insertText(original);
  await expect.poll(() => sheet(page).locator('[data-page]').count()).toBeGreaterThan(2);
  const paragraphs = sheet(page).locator('[data-paragraph]');
  const allText = () => paragraphs.evaluateAll(items => items.map(item => item.textContent).join(''));
  expect(await allText()).toBe(original);
  await page.keyboard.press('Backspace');
  expect(await allText()).toBe(original.slice(0, -1));
  await page.keyboard.press('Control+z');
  expect(await allText()).toBe(original);
});

test('composes into a blank paragraph and reopens the committed text', async ({ page }) => {
  await firstParagraph(page).click({ position: { x: 8, y: 5 } });
  const failures: string[] = [];
  page.on('pageerror', error => failures.push(error.message));
  if (test.info().project.name === 'chromium') {
    const cdp = await page.context().newCDPSession(page);
    await cdp.send('Input.imeSetComposition', { text: 'ni', selectionStart: 2, selectionEnd: 2 });
    await cdp.send('Input.insertText', { text: '你' });
    await cdp.send('Input.imeSetComposition', { text: 'hao', selectionStart: 3, selectionEnd: 3 });
    await cdp.send('Input.insertText', { text: '好' });
  } else {
    await page.keyboard.insertText('你'); await page.keyboard.insertText('好');
  }
  await expect(firstParagraph(page)).toHaveText('你好');
  expect(await open(page, await exportBytes(page))).toBe('OK');
  await expect(firstParagraph(page)).toHaveText('你好');
  expect(failures).toEqual([]);
});

test('edits repeating headers and footers and exports native page fields', async ({ page }) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  await page.getByRole('tab', { name: 'Insert' }).click();
  await page.getByRole('button', { name: 'Edit header' }).click();
  await page.keyboard.insertText('Editorial desk');
  await expect(sheet(page).locator('[data-page-region="header"]')).toHaveText(['Editorial desk', 'Editorial desk']);
  await page.getByRole('button', { name: 'Close header and footer' }).click();
  await page.getByRole('tab', { name: 'Insert' }).click();
  await page.getByRole('button', { name: 'Edit footer' }).click();
  await page.keyboard.insertText('Page ');
  await page.getByRole('button', { name: 'Page number', exact: true }).click();
  await page.keyboard.insertText(' of ');
  await page.getByRole('button', { name: 'Total pages' }).click();
  await expect(sheet(page).locator('[data-page-region="footer"]')).toHaveText(['Page 1 of 2', 'Page 2 of 2']);
  const edited = await exportBytes(page);
  const parts = await readZip(edited);
  expect(Buffer.from(parts.get('word/footer.xml')!).toString()).toContain('w:instr="PAGE"');
  expect(Buffer.from(parts.get('word/footer.xml')!).toString()).toContain('w:instr="NUMPAGES"');
  expect(await open(page, edited)).toBe('OK');
  await expect(sheet(page).locator('[data-page-region="header"]')).toHaveText(['Editorial desk', 'Editorial desk']);
  await expect(sheet(page).locator('[data-page-region="footer"]')).toHaveText(['Page 1 of 2', 'Page 2 of 2']);
});

test('modifies existing heading styles globally and creates reusable custom styles', async ({ page }) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  await select(page, 'Overview', 0, 'Overview', 0);
  await page.getByRole('button', { name: 'Modify style' }).click();
  await page.getByLabel('Style name').fill('Editorial heading');
  await page.getByLabel('Font size (pt)').fill('24');
  await page.getByLabel('Alignment', { exact: true }).selectOption('center');
  await page.getByRole('button', { name: 'Save style' }).click();
  await expect(sheet(page).getByRole('heading', { name: 'Overview' })).toHaveCSS('font-size', '32px');
  await expect(sheet(page).getByRole('heading', { name: 'Story priorities' })).toHaveCSS('text-align', 'center');
  await select(page, 'Before sign-off', 0, 'Before sign-off', 0);
  await page.getByRole('button', { name: 'New style' }).click();
  await page.getByLabel('Style name').fill('Approval heading');
  await page.getByLabel('Outline level').selectOption('7');
  await page.getByRole('button', { name: 'Save style' }).click();
  await expect(sheet(page).getByRole('heading', { name: 'Before sign-off', level: 7 })).toBeVisible();
  const bytes = await exportBytes(page);
  expect(await open(page, bytes)).toBe('OK');
  await expect(sheet(page).getByRole('heading', { name: 'Before sign-off', level: 7 })).toBeVisible();
  await expect(sheet(page).getByRole('heading', { name: 'Overview' })).toHaveCSS('font-size', '32px');
});

test('pastes Word HTML formatting, list numbering, and tables and round-trips the result', async ({ page }, testInfo) => {
  await firstParagraph(page).click();
  await sheet(page).evaluate(element => {
    const data = new DataTransfer();
    data.setData('text/plain', 'Word heading\nFirst\nSecond\nLeft\tRight');
    data.setData('text/html', `
      <html xmlns:o="urn:schemas-microsoft-com:office:office"><head><style>
      p.MsoHeading1 {font-family:Georgia;color:#243b60;font-size:22pt}
      @list l0:level1 {mso-level-start-at:3;mso-level-number-format:decimal}
      </style></head><body><!--StartFragment-->
      <p class="MsoHeading1" style="text-align:center"><strong>Word heading</strong></p>
      <p class="MsoListParagraph" style="mso-list:l0 level1 lfo1"><span style="mso-list:Ignore">3.<span>&nbsp;</span></span><i>First</i></p>
      <p class="MsoListParagraph" style="mso-list:l0 level1 lfo1"><span style="mso-list:Ignore">4.&nbsp;</span><u>Second</u></p>
      <table><tr><td style="background-color:#eff4fc"><p><a href="https://example.com">Left</a></p></td>
      <td><p><span style="background-color:yellow">Right</span></p></td></tr></table>
      <script>window.pasteScriptExecuted = true</script>
      <!--EndFragment--></body></html>`);
    element.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  });
  const heading = sheet(page).getByRole('heading', { name: 'Word heading' });
  await expect(heading).toHaveCSS('text-align', 'center');
  await expect(heading.locator('[data-inline-start]')).toHaveCSS('font-family', /Georgia/);
  await expect(sheet(page).locator('.rwe-list-label')).toHaveText(['3.', '4.']);
  await expect(sheet(page).locator('[data-href="https://example.com/"]')).toHaveText('Left');
  await expect(sheet(page).getByRole('cell')).toHaveCount(2);
  await expect(sheet(page).getByText('Right', { exact: true })).toHaveCSS('background-color', 'rgb(255, 255, 0)');
  const event = page.waitForEvent('download');
  await download(page).click();
  const saved = await event;
  const path = testInfo.outputPath('word-paste.docx');
  await saved.saveAs(path);
  expect(await open(page, await readFile(path))).toBe('OK');
  await expect(heading).toHaveCSS('text-align', 'center');
  await expect(sheet(page).locator('.rwe-list-label')).toHaveText(['3.', '4.']);
  await expect(sheet(page).getByText('Right', { exact: true })).toHaveCSS('background-color', 'rgb(255, 255, 0)');
});

test('adds comments and tracked text changes and reopens them for review', async ({ page }, testInfo) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  await select(page, 'Overview', 0);
  await page.getByRole('tab', { name: 'Review' }).click();
  await page.getByRole('button', { name: 'New comment' }).click();
  await page.getByLabel('Author').fill('Riya');
  await page.getByLabel('Comment', { exact: true }).fill('Check the scope');
  await page.getByRole('button', { name: 'Add comment' }).click();
  await expect(sheet(page).locator('.rwe-comment-range')).toHaveText('Overview');
  await select(page, 'Story priorities', 0);
  await page.getByRole('button', { name: 'Track changes' }).click();
  await page.keyboard.insertText('Priorities');
  await expect(sheet(page).locator('.rwe-review-ins')).toHaveText('Priorities');
  await expect(sheet(page).locator('.rwe-review-del')).toHaveText('Story priorities');
  const event = page.waitForEvent('download');
  await download(page).click();
  const path = testInfo.outputPath('review.docx');
  await (await event).saveAs(path);
  expect(await open(page, await readFile(path))).toBe('OK');
  await expect(sheet(page)).toHaveAttribute('aria-readonly', 'false');
  await page.getByRole('tab', { name: 'Review' }).click();
  await page.getByRole('button', { name: 'Review pane' }).click();
  await expect(page.getByText('Check the scope')).toBeVisible();
  await page.getByRole('article').filter({ hasText: 'Insertion' }).getByRole('button', { name: 'Accept' }).click();
  await page.getByRole('article').filter({ hasText: 'Deletion' }).getByRole('button', { name: 'Accept' }).click();
  await expect(sheet(page).getByRole('heading', { name: 'Priorities', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Delete comment' }).click();
  await expect(sheet(page).locator('.rwe-comment-range')).toHaveCount(0);
});

test('pastes a clipboard picture and edits its dimensions and alternative text', async ({ page }, testInfo) => {
  await firstParagraph(page).click();
  await sheet(page).evaluate((element, bytes) => {
    const data = new DataTransfer();
    data.items.add(new File([new Uint8Array(bytes)], 'Clipboard.png', { type: 'image/png' }));
    element.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  }, Array.from(png));
  const image = sheet(page).getByRole('img', { name: 'Clipboard.png' });
  await expect(image).toBeVisible();
  await image.click();
  await page.getByRole('button', { name: 'Size and alternative text' }).click();
  await page.getByLabel('Width (px)').fill('120');
  await page.getByLabel('Alternative text', { exact: true }).fill('Blue sample');
  await page.getByRole('button', { name: 'Apply' }).click();
  await expect(sheet(page).getByRole('img', { name: 'Blue sample' })).toHaveAttribute('width', '120');
  const bytes = await exportBytes(page);
  await expect(await open(page, bytes)).toBe('OK');
  await expect(sheet(page).getByRole('img', { name: 'Blue sample' })).toHaveAttribute('height', '60');
  const event = page.waitForEvent('download');
  await download(page).click();
  await (await event).saveAs(testInfo.outputPath('picture.docx'));
});

test('changes table columns and rows, applies cell shading, and updates page dimensions', async ({ page }, testInfo) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  await select(page, 'Milestone', 0, 'Milestone', 0);
  await page.getByLabel('Table action').selectOption('column-after');
  await expect(sheet(page).getByRole('cell')).toHaveCount(16);
  await page.getByLabel('Table action').selectOption('row-after');
  await expect(sheet(page).getByRole('cell')).toHaveCount(20);
  await page.getByLabel('Cell shading').fill('#ffccaa');
  await expect(sheet(page).locator('td').nth(5)).toHaveCSS('background-color', 'rgb(255, 204, 170)');
  await page.getByLabel('Table action').selectOption('delete-column');
  await page.getByRole('button', { name: 'Delete', exact: true }).click();
  await expect(sheet(page).getByRole('cell')).toHaveCount(15);
  await page.getByRole('tab', { name: 'Layout' }).click();
  await page.getByRole('button', { name: 'Page setup' }).click();
  await page.getByLabel('Width (in)').fill('10');
  await page.getByLabel('Header from top (in)').fill('0.4');
  await page.getByRole('button', { name: 'Apply' }).click();
  const event = page.waitForEvent('download');
  await download(page).click();
  const path = testInfo.outputPath('layout.docx');
  await (await event).saveAs(path);
  expect(await open(page, await readFile(path))).toBe('OK');
  await expect(sheet(page).getByRole('cell')).toHaveCount(15);
  expect(await sheet(page).locator('[data-page]').first().evaluate(node => node.getAttribute('style'))).toContain('width: 960px');
});

test('supports first-page and even-page header variants without losing the default', async ({ page }, testInfo) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  const openHeader = async () => {
    await page.getByRole('tab', { name: 'Insert' }).click();
    await page.getByRole('button', { name: 'Edit header' }).click();
    await expect(page.getByRole('group', { name: 'Header and footer tools' })).toBeVisible();
  };
  await openHeader();
  await page.keyboard.insertText('Main header');
  await page.getByLabel('Different first page').click();
  await select(page, 'Overview', 0, 'Overview', 0);
  await openHeader();
  await page.keyboard.insertText('Cover header');
  await page.getByLabel('Different odd and even pages').click();
  await select(page, 'Review and handoff', 0, 'Review and handoff', 0);
  await openHeader();
  await page.keyboard.insertText('Even header');
  const event = page.waitForEvent('download');
  await download(page).click();
  const path = testInfo.outputPath('headers.docx');
  await (await event).saveAs(path);
  expect(await open(page, await readFile(path))).toBe('OK');
  await expect(sheet(page).locator('[data-page-region="header"]')).toHaveText(['Cover header', 'Even header']);
  const parts = await readZip(await readFile(path));
  expect([...parts.values()].some(part => Buffer.from(part).toString().includes('Main header'))).toBe(true);
});

test('pastes rich clipboard data through Ctrl+V and supports merge formatting', async ({ page, context }) => {
  await enableClipboard(context);
  await firstParagraph(page).click();
  await page.getByLabel('Paste formatting').selectOption('merge');
  await firstParagraph(page).click();
  await page.evaluate(async () => {
    const html = '<p style="text-align:right"><span style="font-size:24pt;color:#b00020"><b>Copy from Word</b></span></p>';
    await navigator.clipboard.write([new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob(['Copy from Word'], { type: 'text/plain' }),
    })]);
  });
  await page.keyboard.press('Control+v');
  const inserted = sheet(page).getByText('Copy from Word', { exact: true });
  await expect(inserted).toHaveCSS('font-weight', '700');
  await expect(inserted).toHaveCSS('color', 'rgb(36, 39, 45)');
  await expect(firstParagraph(page)).not.toHaveCSS('text-align', 'right');
  await page.keyboard.press('Control+z');
  await expect(inserted).toHaveCount(0);
});

test('copies formatted text and structure through the system clipboard', async ({ page, context }) => {
  await enableClipboard(context);
  await firstParagraph(page).click();
  await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData('text/html', `<h2>Roadmap</h2>
      <p style="text-align:right">Before <b><i><u><a href="https://example.com/review">
        <span style="font-family:Georgia;font-size:18pt;color:#b00020;background-color:#ffeb3b">review notes</span></a></u></i></b> after</p>
      <ol start="4"><li>Approve draft</li><li>Publish draft</li></ol>
      <table><tr><td style="background-color:#fde0b0">Owner</td><td>Editor</td></tr></table>`);
    document.querySelector('.rwe-pages')!.dispatchEvent(new ClipboardEvent('paste', {
      clipboardData: data, bubbles: true, cancelable: true,
    }));
  });
  await expect(sheet(page).getByRole('heading', { name: 'Roadmap' })).toBeVisible();
  await select(page, 'Roadmap', 1, 'Editor');
  await page.keyboard.press('Control+c');

  const clipboard = await page.evaluate(async () => {
    const item = (await navigator.clipboard.read())[0];
    const html = item.types.includes('text/html') ? await (await item.getType('text/html')).text() : '';
    return { html, plain: await navigator.clipboard.readText() };
  });
  expect(clipboard.html).toContain('text-decoration: underline');
  expect(clipboard.plain).toContain('oadmap');
  expect(clipboard.plain).not.toContain('Roadmap');

  await page.evaluate(() => {
    const receiver = document.createElement('div');
    receiver.contentEditable = 'true';
    receiver.id = 'clipboard-receiver';
    document.body.append(receiver);
    receiver.focus();
  });
  await page.keyboard.press('Control+v');
  const receiver = page.locator('#clipboard-receiver');
  await expect(receiver.getByRole('heading', { level: 2 })).toHaveText('oadmap');
  const link = receiver.getByRole('link', { name: 'review notes' });
  await expect(link).toHaveAttribute('href', 'https://example.com/review');
  const styled = receiver.getByText('review notes', { exact: true });
  await expect(styled).toHaveCSS('font-weight', '700');
  await expect(styled).toHaveCSS('font-style', 'italic');
  await expect(styled).toHaveCSS('font-family', /Georgia/);
  await expect(styled).toHaveCSS('font-size', '24px');
  await expect(styled).toHaveCSS('color', 'rgb(176, 0, 32)');
  await expect(styled).toHaveCSS('background-color', 'rgb(255, 235, 59)');
  await expect(receiver.locator('p, div').filter({ has: page.getByRole('link', { name: 'review notes' }) })).toHaveCSS('text-align', 'right');
  await expect(receiver.locator('ol')).toHaveAttribute('start', '4');
  await expect(receiver.getByRole('listitem')).toHaveText(['Approve draft', 'Publish draft']);
  await expect(receiver.getByRole('cell', { name: 'Owner' })).toHaveCSS('background-color', 'rgb(253, 224, 176)');
  await expect(receiver.getByRole('cell', { name: 'Editor' })).toBeVisible();
});

test('cuts a selected table as HTML and restores it on undo', async ({ page, context }) => {
  await enableClipboard(context);
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  await sheet(page).getByRole('cell', { name: 'Milestone', exact: true }).click();
  await page.getByRole('button', { name: 'Select table', exact: true }).click();
  await page.keyboard.press('Control+x');
  await expect(sheet(page).getByRole('table')).toHaveCount(0);
  const content = await page.evaluate(async () => {
    const item = (await navigator.clipboard.read())[0];
    if (!item.types.includes('text/html')) return [];
    const doc = new DOMParser().parseFromString(await (await item.getType('text/html')).text(), 'text/html');
    return Array.from(doc.querySelectorAll('table tr:first-child td'), cell => cell.textContent);
  });
  expect(content[0]).toBe('Milestone');
  await page.keyboard.press('Control+z');
  await expect(sheet(page).getByRole('cell', { name: 'Milestone', exact: true })).toBeVisible();
});

test('applies text colors from visual swatches and the custom picker', async ({ page }) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  await select(page, 'Overview', 0);
  const trigger = page.getByRole('button', { name: 'Text color', exact: true });
  const palette = page.getByRole('group', { name: 'Text color palette' });
  await trigger.click();
  await expect(palette).toBeVisible();
  await expect(palette.getByRole('button', { name: 'Red', exact: true })).toHaveCSS('background-color', 'rgb(192, 0, 0)');
  await page.keyboard.press('Escape');
  await expect(palette).toBeHidden();
  await trigger.click();
  await palette.getByRole('button', { name: 'Red', exact: true }).click();
  const overview = sheet(page).getByText('Overview', { exact: true });
  await expect(overview).toHaveCSS('color', 'rgb(192, 0, 0)');

  await trigger.focus();
  await page.keyboard.press('Enter');
  await expect(palette.getByRole('button', { name: 'Red', exact: true })).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect(palette.getByRole('button', { name: 'Orange', exact: true })).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(overview).toHaveCSS('color', 'rgb(237, 125, 49)');

  await trigger.click();
  await palette.getByLabel('Custom color').fill('#347895');
  await palette.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(overview).toHaveCSS('color', 'rgb(52, 120, 149)');
  expect(await open(page, await exportBytes(page))).toBe('OK');
  await expect(overview).toHaveCSS('color', 'rgb(52, 120, 149)');
});

test('runs Word formatting shortcuts, table navigation, and ribbon KeyTips', async ({ page }) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  await select(page, 'Overview', 0);
  await page.keyboard.press('Control+Alt+2');
  const heading = sheet(page).getByRole('heading', { name: 'Overview', level: 2 });
  await expect(heading).toBeVisible();
  await page.keyboard.press('Control+E');
  await expect(heading).toHaveCSS('text-align', 'center');
  await page.keyboard.press('Control+5');
  await expect(heading).toHaveCSS('line-height', '26px');
  await page.keyboard.press('Control+BracketRight');
  await expect(heading.getByText('Overview')).toHaveCSS('font-size', /18\.66/);
  await page.keyboard.press('Shift+F3');
  await expect(sheet(page).getByRole('heading', { name: 'OVERVIEW', level: 2 })).toBeVisible();
  await page.keyboard.press('Control+Z');
  await expect(heading).toBeVisible();
  await select(page, 'Milestone', 0, 'Milestone', 0);
  await page.keyboard.press('Tab');
  expect(await page.evaluate(() => document.getSelection()?.toString())).toBe('Deliverable');
  await page.keyboard.press('Shift+Tab');
  expect(await page.evaluate(() => document.getSelection()?.toString())).toBe('Milestone');
  await page.keyboard.press('Alt+Period');
  await page.keyboard.press('w');
  await expect(page.getByRole('tab', { name: 'View', exact: true })).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('k');
  const guide = page.getByRole('dialog', { name: 'Keyboard shortcuts' });
  await expect(guide).toBeVisible();
  await guide.getByLabel('Search shortcuts').fill('Heading 2');
  await expect(guide.getByText('Apply Heading 2')).toBeVisible();
  await expect(guide.getByText('Ctrl+Alt+2', { exact: true })).toBeVisible();
  await expect(guide.getByText('Apply Heading 1')).toHaveCount(0);
});

test('uses keyboard commands for find, replace, go to page, and comments', async ({ page }) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  await select(page, 'Overview', 0);
  await page.keyboard.press('Control+Alt+M');
  await page.getByLabel('Comment', { exact: true }).fill('Review this');
  await page.keyboard.press('Control+Enter');
  await expect(page.getByText('Review this', { exact: true })).toBeVisible();
  await page.keyboard.press('Control+H');
  await page.getByLabel('Search text').fill('Week');
  await page.getByLabel('Replace with').fill('Sprint');
  await page.getByRole('button', { name: 'Replace all' }).click();
  await expect(sheet(page).getByText('Sprint 1', { exact: true })).toBeVisible();
  await expect(sheet(page).getByText('Sprint 3', { exact: true })).toBeVisible();
  await page.keyboard.press('Control+G');
  await page.getByLabel('Page number', { exact: true }).fill('2');
  await page.getByRole('button', { name: 'Go', exact: true }).click();
  await expect(page.locator('.rwe-status')).toContainText('Page 2 of 2');
  await page.keyboard.press('Control+F');
  await page.getByLabel('Search text').fill('sign-off');
  await page.getByRole('button', { name: 'Find', exact: true }).click();
  expect(await page.evaluate(() => document.getSelection()?.toString())).toBe('sign-off');
});

test('saves a template, reuses it after reload, and fills repeated and split fields', async ({ page }) => {
  const template = await patchFixture(xml => xml.replace('>Overview<', '>{{client<')
    .replace('>{{client</w:t></w:r></w:p>', '>{{client</w:t></w:r><w:r><w:t>_name}}</w:t></w:r></w:p>')
    .replace('>Editorial<', '>{{client_name}}<'));
  expect(await open(page, template)).toBe('OK');
  const library = async () => {
    await page.getByRole('tab', { name: 'File' }).click();
    await page.getByRole('button', { name: 'Document templates' }).click();
    await expect(page.getByRole('dialog', { name: 'Document templates' })).toBeVisible();
  };
  await library();
  await page.getByLabel('Template name').fill('Editorial template');
  await page.getByRole('button', { name: 'Save current document' }).click();
  await expect(page.getByRole('button', { name: 'Editorial template', exact: true })).toBeVisible();
  await page.reload();
  await library();
  await page.getByRole('button', { name: 'Editorial template', exact: true }).click();
  await page.getByLabel('Document name').fill('New editorial');
  await page.getByLabel('client_name').fill('Acme & Co');
  await page.getByRole('button', { name: 'Create document' }).click();
  await expect(page.locator('.rwe-document-title')).toHaveText('New editorial');
  await expect(sheet(page).getByRole('heading', { name: 'Acme & Co' })).toBeVisible();
  await expect(sheet(page).getByRole('cell', { name: 'Acme & Co' })).toBeVisible();
  await expect(page.locator('.rwe-local-state')).toHaveText('Unsaved changes');
  const bytes = await exportBytes(page);
  expect(await open(page, bytes, 'New editorial.docx')).toBe('OK');
  await expect(sheet(page).getByRole('heading', { name: 'Acme & Co' })).toBeVisible();
});

test('inserts a template field and confirms replacement of unsaved work', async ({ page }) => {
  await firstParagraph(page).click();
  await page.getByRole('tab', { name: 'Insert' }).click();
  await page.getByRole('button', { name: 'Insert template field' }).click();
  await page.getByLabel('Field name').fill('project-name');
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await expect(firstParagraph(page)).toHaveText('{{project-name}}');
  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'Document templates' }).click();
  await page.getByRole('button', { name: 'Save current document' }).click();
  await page.getByLabel('project-name', { exact: true }).fill('Project Sunrise');
  await page.getByRole('button', { name: 'Create document' }).click();
  await page.getByRole('button', { name: 'Discard changes' }).click();
  await expect(firstParagraph(page)).toHaveText('Project Sunrise');
});

test('imports a native Word template and preserves formatted headers when generating DOCX', async ({ page }, testInfo) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  await page.getByRole('tab', { name: 'Insert' }).click();
  await page.getByRole('button', { name: 'Edit header' }).click();
  await page.keyboard.insertText('Prepared for {{client}}');
  await page.getByRole('button', { name: 'Close header and footer' }).click();
  const original = await readZip(await exportBytes(page));
  const writer = new ZipWriter(new Uint8ArrayWriter(), { useWebWorkers: false });
  for (const [path, data] of original) {
    if (path === '[Content_Types].xml') {
      await writer.add(path, new TextReader(Buffer.from(data).toString().replace('document.main+xml', 'template.main+xml')));
    } else await writer.add(path, new Uint8ArrayReader(data));
  }
  const bytes = Buffer.from(await writer.close());

  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'Document templates' }).click();
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'Import Word file' }).click();
  await (await chooser).setFiles({ name: 'Client brief.dotx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template', buffer: bytes });
  await page.getByLabel('client', { exact: true }).fill('North & West');
  await page.getByRole('button', { name: 'Create document' }).click();
  await page.getByRole('button', { name: 'Discard changes' }).click();
  await expect(sheet(page).locator('[data-page-region="header"]')).toHaveText(['Prepared for North & West', 'Prepared for North & West']);
  const exported = await exportBytes(page);
  const parts = await readZip(exported);
  expect(Buffer.from(parts.get('[Content_Types].xml')!).toString()).toContain('document.main+xml');
  expect(parts.get('word/styles.xml')).toEqual(original.get('word/styles.xml'));
  expect(await open(page, exported, 'Client brief.docx')).toBe('OK');
  await expect(sheet(page).locator('[data-page-region="header"]').first()).toHaveText('Prepared for North & West');
  const downloadEvent = page.waitForEvent('download');
  await download(page).click();
  await (await downloadEvent).saveAs(testInfo.outputPath('generated.docx'));
});

test('supports the plain-text paste shortcut with source formatting enabled', async ({ page, context }) => {
  await enableClipboard(context);
  await firstParagraph(page).click();
  await page.evaluate(async () => {
    await navigator.clipboard.write([new ClipboardItem({
      'text/html': new Blob(['<p><b>Plain paste</b></p>'], { type: 'text/html' }),
      'text/plain': new Blob(['Plain paste'], { type: 'text/plain' }),
    })]);
  });
  await page.keyboard.press('Control+Shift+V');
  await expect(sheet(page).getByText('Plain paste', { exact: true })).toHaveCSS('font-weight', '400');
});

test('selects and removes a whole table with Delete, then restores it with undo', async ({ page }, testInfo) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  const table = sheet(page).getByRole('table', { name: 'Document table' });
  await table.getByText('Milestone', { exact: true }).hover();
  await sheet(page).getByRole('button', { name: 'Select entire table' }).click();
  await expect(sheet(page).locator('[data-table-selected="true"]')).toHaveCount(1);
  await page.keyboard.press('Delete');
  await expect(table).toHaveCount(0);
  await expect(sheet(page).getByText('Production schedule', { exact: true })).toBeVisible();
  await expect(sheet(page).getByText('Review and handoff', { exact: true })).toHaveCount(1);
  await page.keyboard.press('Control+Z');
  await expect(table.getByRole('cell')).toHaveCount(12);
  await expect(sheet(page).locator('[data-table-selected="true"]')).toHaveCount(1);
  await page.keyboard.press('Control+Y');
  await expect(table).toHaveCount(0);

  const bytes = await exportBytes(page);
  const parts = await readZip(bytes);
  expect(Buffer.from(parts.get('word/document.xml')!).toString()).not.toContain('<w:tbl>');
  await writeFile(testInfo.outputPath('table-deleted.docx'), bytes);
  expect(await open(page, bytes)).toBe('OK');
  await expect(table).toHaveCount(0);
});

test('deletes a table from a full text selection or the table tools without deleting adjacent content', async ({ page }) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  const table = sheet(page).getByRole('table', { name: 'Document table' });
  await select(page, 'Milestone', 0, 'Production');
  await expect(page.getByText('Table selected', { exact: true })).toBeVisible();
  await page.keyboard.press('Backspace');
  await expect(table).toHaveCount(0);

  await page.keyboard.press('Control+Z');
  await expect(table).toHaveCount(1);
  const tools = page.getByRole('group', { name: 'Table tools' });
  await tools.getByRole('button', { name: 'Delete table', exact: true }).click();
  const confirm = page.getByRole('dialog', { name: 'Delete entire table?' });
  await confirm.getByRole('button', { name: 'Delete', exact: true }).click();
  await expect(table).toHaveCount(0);
  await expect(sheet(page).getByText('Production schedule', { exact: true })).toBeVisible();
});

test('keeps an ordinary cell text selection distinct from a whole-table selection', async ({ page }) => {
  await firstParagraph(page).click();
  await page.getByRole('tab', { name: 'Insert' }).click();
  await page.getByRole('button', { name: 'Insert table' }).click();
  await page.getByLabel('Rows').fill('1');
  await page.getByLabel('Columns').fill('1');
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.keyboard.insertText('Cell text');
  await select(page, 'Cell text', 0);
  await page.keyboard.press('Backspace');
  await expect(sheet(page).getByRole('cell')).toHaveCount(1);
  await expect(sheet(page).getByRole('cell')).toHaveText('');
  await page.keyboard.press('Alt+Numpad5');
  await expect(page.getByRole('button', { name: 'Select table', exact: true })).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('Delete');
  await expect(sheet(page).getByRole('table')).toHaveCount(0);
});

test('deletes every fragment of a table spread across pages', async ({ page }) => {
  const fixture = await patchFixture(xml => xml.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/, table => {
    const row = /<w:tr>[\s\S]*?<\/w:tr>/.exec(table)?.[0];
    if (!row) throw new Error('Missing fixture row');
    return table.replace('</w:tbl>', row.repeat(36) + '</w:tbl>');
  }));
  expect(await open(page, fixture)).toBe('OK');
  await expect.poll(() => sheet(page).getByRole('table').count()).toBeGreaterThan(1);
  await sheet(page).getByRole('cell', { name: 'Milestone', exact: true }).first().click();
  await page.getByRole('button', { name: 'Select table', exact: true }).click();
  await expect(sheet(page).locator('[data-table-selected="true"]')).toHaveCount(await sheet(page).getByRole('table').count());
  await page.keyboard.press('Backspace');
  await expect(sheet(page).getByRole('table')).toHaveCount(0);
});

test('deletes a selection that includes a table and surrounding paragraphs', async ({ page }) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  await select(page, 'Milestone', 0, 'Milestone', 0);
  await page.keyboard.press('Control+A');
  await page.keyboard.press('Delete');
  await expect(sheet(page).getByRole('table')).toHaveCount(0);
  await expect(firstParagraph(page)).toHaveText('');
  await page.keyboard.insertText('New text');
  await expect(firstParagraph(page)).toHaveText('New text');
});

test('formats table layout and scoped cells, undoes the change, and preserves it in DOCX', async ({ page }, testInfo) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  await expect(page.getByRole('tab', { name: 'Format Table', exact: true })).toHaveCount(0);
  await select(page, 'Milestone', 0, 'Milestone', 0);
  await expect(page.getByRole('tablist', { name: 'Editor ribbon' }).getByRole('tab')).toHaveText([
    'File', 'Home', 'Insert', 'Layout', 'Review', 'View', 'Format Table',
  ]);
  const formatTab = page.getByRole('tab', { name: 'Format Table', exact: true });
  await expect(formatTab).toHaveAttribute('aria-selected', 'true');
  await page.getByRole('tab', { name: 'View', exact: true }).click();
  await page.getByRole('tab', { name: 'View', exact: true }).press('ArrowRight');
  await expect(formatTab).toBeFocused();
  await page.getByRole('tabpanel', { name: 'Format Table commands', exact: true }).getByRole('button', { name: 'Format table' }).click();
  const dialog = page.getByRole('dialog', { name: 'Format table' });
  await dialog.getByLabel('Table width (%)').fill('78');
  await dialog.getByLabel('Table alignment').selectOption('center');
  await dialog.getByLabel('Borders', { exact: true }).selectOption('outline');
  await dialog.getByLabel('Border style').selectOption('dashed');
  await dialog.getByLabel('Border width (pt)').fill('1.5');
  await dialog.getByLabel('Border color').fill('#223344');
  await dialog.getByLabel('Row 1 minimum height (px)').fill('44');
  await dialog.getByLabel('Horizontal cell padding (px)').fill('12');
  await dialog.getByLabel('Vertical cell padding (px)').fill('7');
  await dialog.getByLabel('Distribute columns evenly').check();
  await dialog.getByLabel('Apply cell formatting to').selectOption('column');
  await dialog.getByLabel('Vertical alignment').selectOption('bottom');
  await dialog.getByLabel('Shading', { exact: true }).selectOption('color');
  await dialog.getByLabel('Shading color').fill('#fde0b0');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();

  const table = sheet(page).getByRole('table');
  await expect(table).toHaveAttribute('style', /width: 78%/);
  const cell = table.getByRole('cell', { name: 'Week 1', exact: true });
  await expect(cell).toHaveCSS('background-color', 'rgb(253, 224, 176)');
  await expect(cell).toHaveCSS('vertical-align', 'bottom');
  await expect(cell).toHaveCSS('padding-left', '12px');
  await expect(table.getByRole('cell', { name: 'Editorial', exact: true })).toHaveCSS('vertical-align', 'top');
  await page.keyboard.press('Control+Z');
  await expect(table).not.toHaveAttribute('style', /width: 78%/);
  await page.keyboard.press('Control+Y');
  await expect(table).toHaveAttribute('style', /width: 78%/);

  const bytes = await exportBytes(page);
  const parts = await readZip(bytes);
  const xml = Buffer.from(parts.get('word/document.xml')!).toString();
  expect(xml).toContain('w:w="3900" w:type="pct"');
  expect(xml).toContain('w:val="dashed" w:sz="12" w:color="223344"');
  expect(xml).toContain('w:val="660" w:hRule="atLeast"');
  expect(xml).toContain('<w:vAlign w:val="bottom"');
  await writeFile(testInfo.outputPath('formatted-table.docx'), bytes);
  expect(await open(page, bytes)).toBe('OK');
  await expect(formatTab).toHaveCount(0);
  await expect(table).toHaveAttribute('style', /width: 78%/);
  await expect(cell).toHaveCSS('background-color', 'rgb(253, 224, 176)');
  await expect(cell).toHaveCSS('border-left-style', 'dashed');
});

test('repeats the header of a multipage table and exports one source header row', async ({ page }, testInfo) => {
  const fixture = await patchFixture(xml => xml.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/, table => {
    const row = Array.from(table.matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g))[1]?.[0];
    if (!row) throw new Error('Missing fixture body row');
    return table.replace('</w:tbl>', row.repeat(52) + '</w:tbl>');
  }));
  expect(await open(page, fixture)).toBe('OK');
  const fragments = sheet(page).getByRole('table');
  await expect.poll(() => fragments.count()).toBeGreaterThan(2);
  await fragments.first().getByRole('cell', { name: 'Milestone', exact: true }).click();
  await page.getByRole('button', { name: 'Format table' }).click();
  const dialog = page.getByRole('dialog', { name: 'Format table' });
  await dialog.getByLabel('Repeat header row on each page').check();
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  const repeat = sheet(page).locator('[data-repeat-row]');
  await expect(repeat).toHaveCount(await fragments.count() - 1);
  await expect(repeat.first()).toContainText('Milestone');

  const bytes = await exportBytes(page);
  const parts = await readZip(bytes);
  const xml = Buffer.from(parts.get('word/document.xml')!).toString();
  expect((xml.match(/<w:tblHeader\b/g) ?? []).length).toBe(1);
  expect((xml.match(/>Milestone</g) ?? []).length).toBe(1);
  await writeFile(testInfo.outputPath('repeating-table.docx'), bytes);
  expect(await open(page, bytes)).toBe('OK');
  await expect(repeat).toHaveCount(await fragments.count() - 1);
});

test('reuses row measurements when shading a large table and refreshes them when its width changes', async ({ page }) => {
  const fixture = await patchFixture(xml => xml.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/, table => {
    const row = Array.from(table.matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g))[1]?.[0];
    if (!row) throw new Error('Missing fixture body row');
    return table.replace('</w:tbl>', row.repeat(150) + '</w:tbl>');
  }));
  expect(await open(page, fixture)).toBe('OK');
  const tables = sheet(page).getByRole('table');
  await expect.poll(() => tables.count()).toBeGreaterThan(3);
  await tables.first().getByRole('cell', { name: 'Milestone', exact: true }).click();
  await page.evaluate(() => {
    const measure = Element.prototype.getBoundingClientRect;
    const tally = { reads: 0 };
    Object.assign(window, { tableMeasurementTally: tally });
    Element.prototype.getBoundingClientRect = function() {
      if (this.hasAttribute('data-measure') && this.firstElementChild?.matches('table[data-table-id]')) tally.reads++;
      return measure.call(this);
    };
  });
  const readTally = () => page.evaluate(() => (window as unknown as { tableMeasurementTally: { reads: number } }).tableMeasurementTally.reads);
  await page.getByRole('button', { name: 'Format table' }).click();
  const dialog = page.getByRole('dialog', { name: 'Format table' });
  await dialog.getByLabel('Apply cell formatting to').selectOption('table');
  await dialog.getByLabel('Shading', { exact: true }).selectOption('color');
  await dialog.getByLabel('Shading color').fill('#fce9cf');
  await dialog.getByRole('button', { name: 'Apply' }).click();
  await expect(tables.last().getByRole('cell').last()).toHaveCSS('background-color', 'rgb(252, 233, 207)');
  expect(await readTally()).toBe(0);

  await page.getByRole('button', { name: 'Format table' }).click();
  await dialog.getByLabel('Table width (%)').fill('65');
  await dialog.getByRole('button', { name: 'Apply' }).click();
  await expect(tables.first()).toHaveAttribute('style', /width: 65%/);
  expect(await readTally()).toBeGreaterThanOrEqual(154);
});

test('draws two strokes for thin double table borders at reduced zoom and preserves the DOCX width', async ({ page }, testInfo) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  const cell = sheet(page).getByRole('cell', { name: 'Milestone', exact: true });
  await cell.click();
  await page.getByRole('button', { name: 'Format table' }).click();
  const dialog = page.getByRole('dialog', { name: 'Format table' });
  await expect(dialog.getByLabel('Border width (pt)')).toHaveValue('0.5');
  await dialog.getByLabel('Border style').selectOption('double');
  await dialog.getByLabel('Border color').fill('#17263b');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();

  const bytes = await exportBytes(page);
  const parts = await readZip(bytes);
  expect(Buffer.from(parts.get('word/document.xml')!).toString()).toContain('w:val="double" w:sz="4" w:color="17263B"');
  await writeFile(testInfo.outputPath('double-border.docx'), bytes);
  expect(await open(page, bytes)).toBe('OK');
  await expect(cell).toHaveCSS('border-top-style', 'double');

  for (const zoom of [100, 50, 25]) {
    await page.getByRole('slider', { name: 'Zoom', exact: true }).fill(String(zoom));
    await cell.scrollIntoViewIfNeeded();
    const box = (await cell.boundingBox())!;
    const screenshot = await page.screenshot({
      clip: { x: Math.round(box.x + box.width / 2), y: Math.floor(box.y), width: 1, height: 8 },
      scale: 'css',
    });
    const strokes = await page.evaluate(async data => {
      const bitmap = await createImageBitmap(new Blob([new Uint8Array(data)], { type: 'image/png' }));
      const canvas = document.createElement('canvas');
      canvas.width = bitmap.width;
      canvas.height = bitmap.height;
      const context = canvas.getContext('2d')!;
      context.drawImage(bitmap, 0, 0);
      const pixels = context.getImageData(0, 0, bitmap.width, bitmap.height).data;
      bitmap.close();
      let count = 0;
      let on = false;
      for (let pixel = 0; pixel < pixels.length; pixel += 4) {
        const ink = pixels[pixel] === 23 && pixels[pixel + 1] === 38 && pixels[pixel + 2] === 59;
        if (ink && !on) count++;
        on = ink;
      }
      return count;
    }, Array.from(screenshot));
    expect(strokes, `visible border strokes at ${zoom}%`).toBe(2);
  }
});

test('applies decimal border widths over imported overrides and saves the native weight', async ({ page }, testInfo) => {
  const fixture = await patchFixture(xml => xml.replace('<w:tr>', `<w:tr><w:tblPrEx>
    <w:tblBorders><w:top w:val="single" w:sz="4" w:color="CBD3DF"/></w:tblBorders>
    <w:tblCellMar><w:top w:w="75" w:type="dxa"/></w:tblCellMar>
  </w:tblPrEx>`).replace('<w:shd ', `<w:tcBorders>
    <w:top w:val="single" w:sz="4" w:color="CBD3DF"/>
    <w:left w:val="single" w:sz="4" w:color="CBD3DF"/>
    <w:tl2br w:val="single" w:sz="4" w:color="CBD3DF"/>
  </w:tcBorders><w:shd `));
  expect(await open(page, fixture)).toBe('OK');
  const cell = sheet(page).getByRole('cell', { name: 'Milestone', exact: true });
  await cell.click();
  const format = page.getByRole('button', { name: 'Format table' });
  const dialog = page.getByRole('dialog', { name: 'Format table' });
  const width = dialog.getByLabel('Border width (pt)');

  await format.click();
  await width.fill('1.2');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(dialog).toBeHidden();
  await format.click();
  await expect(width).toHaveValue('1.25');
  await width.fill('3');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(cell).toHaveCSS('border-top-width', '4px');
  await expect(cell).toHaveCSS('border-left-width', '4px');

  const bytes = await exportBytes(page);
  const parts = await readZip(bytes);
  const xml = Buffer.from(parts.get('word/document.xml')!).toString();
  expect(xml).toContain('w:val="single" w:sz="24"');
  expect(xml).toMatch(/<w:tcBorders>\s*<w:tl2br [^>]+\/>\s*<\/w:tcBorders>/);
  expect(xml).toMatch(/<w:tblPrEx>\s*<w:tblCellMar>/);
  await writeFile(testInfo.outputPath('border-width.docx'), bytes);
  expect(await open(page, bytes)).toBe('OK');
  await expect(cell).toHaveCSS('border-top-width', '4px');
});

test('changing a double border from half a point to two points thickens it on the canvas', async ({ page }) => {
  expect(await open(page, await readFile(fixturePath))).toBe('OK');
  const cell = sheet(page).getByRole('cell', { name: 'Milestone', exact: true });
  await cell.click();
  const format = page.getByRole('button', { name: 'Format table' });
  const dialog = page.getByRole('dialog', { name: 'Format table' });
  await format.click();
  await dialog.getByLabel('Border style').selectOption('double');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(cell).toHaveCSS('border-top-style', 'double');
  const thin = Number.parseFloat(await cell.evaluate(node => getComputedStyle(node).borderTopWidth));

  await format.click();
  await dialog.getByLabel('Border width (pt)').fill('2');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect.poll(async () => Number.parseFloat(await cell.evaluate(node => getComputedStyle(node).borderTopWidth)))
    .toBeGreaterThan(thin);
});

test('inherits Word clipboard wrapper formatting, merges CSS rules, and skips hidden blocks', async ({ page }) => {
  await firstParagraph(page).click();
  await pasteHtml(page, `<style>p {font-style:italic} p.WordBody {font-weight:bold} p.WordBody {color:#b00020}</style>
    <div style="font-family:Georgia;font-size:20pt;text-align:center"><p class="WordBody">Inherited format</p>
    <p hidden>Hidden text</p><p style="display:none">Also hidden</p></div>`);
  const run = sheet(page).getByText('Inherited format', { exact: true });
  await expect(run).toHaveCSS('font-family', /Georgia/);
  await expect(run).toHaveCSS('font-size', /26\.66\d+px/);
  await expect(run).toHaveCSS('font-weight', '700');
  await expect(run).toHaveCSS('font-style', 'italic');
  await expect(run).toHaveCSS('color', 'rgb(176, 0, 32)');
  await expect(firstParagraph(page)).toHaveCSS('text-align', 'center');
  await expect(sheet(page)).not.toContainText('Hidden text');
  await expect(sheet(page)).not.toContainText('Also hidden');
  expect(await open(page, await exportBytes(page))).toBe('OK');
  await expect(run).toHaveCSS('font-weight', '700');
  await expect(firstParagraph(page)).toHaveCSS('text-align', 'center');
});

for (const [rows, columns] of [[201, 1], [1, 21]]) {
  test(`preserves all plain text when a pasted table exceeds ${rows} by ${columns} rich-text limits`, async ({ page }) => {
    await firstParagraph(page).click();
    const values = Array.from({ length: rows }, (_, row) => Array.from({ length: columns }, (_, column) => `Cell ${row}:${column}`));
    await pasteHtml(page, '<table>' + values.map(row => '<tr>' + row.map(value => `<td>${value}</td>`).join('') + '</tr>').join('') + '</table>',
      values.map(row => row.join('\t')).join('\n'));
    await expect(sheet(page).getByRole('table')).toHaveCount(0);
    await expect(sheet(page)).toContainText(`Cell ${rows - 1}:${columns - 1}`);
    await expect(page.getByRole('status').filter({ hasText: 'rich-text limits' })).toBeVisible();
    const parts = await readZip(await exportBytes(page));
    const xml = Buffer.from(parts.get('word/document.xml')!).toString();
    for (const row of values) for (const value of row) expect(xml).toContain(value);
    await page.keyboard.press('Control+z');
    await expect(sheet(page)).not.toContainText('Cell 0:0');
  });
}

test('numbers nested pasted lists correctly before and after DOCX export', async ({ page }) => {
  await firstParagraph(page).click();
  await pasteHtml(page, '<ol><li>Parent<ol start="4"><li>Nested four</li><li>Nested five</li></ol></li><li>Next parent</li></ol>');
  await expect(sheet(page).locator('.rwe-list-label')).toHaveText(['1.', '4.', '5.', '2.']);
  expect(await open(page, await exportBytes(page))).toBe('OK');
  await expect(sheet(page).locator('.rwe-list-label')).toHaveText(['1.', '4.', '5.', '2.']);
});

test('copies table cells as tab-separated plain text and retains HTML structure', async ({ page, context }) => {
  await enableClipboard(context);
  await firstParagraph(page).click();
  await pasteHtml(page, '<table><tr><td>North</td><td>South</td></tr><tr><td>East</td><td>West</td></tr></table>');
  await sheet(page).getByRole('cell', { name: 'North', exact: true }).click();
  await page.getByRole('button', { name: 'Select table', exact: true }).click();
  await page.keyboard.press('Control+c');
  expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('North\tSouth\nEast\tWest');
  await page.keyboard.press('Control+x');
  expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('North\tSouth\nEast\tWest');
  await expect(sheet(page).getByRole('table')).toHaveCount(0);
  await page.keyboard.press('Control+z');
  await expect(sheet(page).getByRole('cell')).toHaveCount(4);
});

test('preserves picture selections when clipboard writing fails', async ({ page }) => {
  await firstParagraph(page).click();
  await pasteHtml(page, `<p>Before<img alt="Embedded sample" src="data:image/png;base64,${png.toString('base64')}">After</p>`);
  const image = sheet(page).getByRole('img', { name: 'Embedded sample' });
  await expect(image).toBeVisible();
  await select(page, 'Before', 0, 'After');
  const before = await readZip(await exportBytes(page));
  await page.evaluate(() => Object.defineProperty(navigator.clipboard, 'write', { configurable: true,
    value: () => Promise.reject(new DOMException('Permission denied', 'NotAllowedError')) }));
  await sheet(page).evaluate(element => element.dispatchEvent(new ClipboardEvent('cut', {
    clipboardData: new DataTransfer(), bubbles: true, cancelable: true,
  })));
  await expect(page.getByRole('alert')).toContainText('Your document was kept unchanged');
  await expect(image).toBeVisible();
  const after = await readZip(await exportBytes(page));
  expect(after).toEqual(before);
});

test('reports legacy and encrypted Word files clearly without replacing the current document', async ({ page }) => {
  await firstParagraph(page).click();
  await page.keyboard.type('Keep this draft');
  expect(await open(page, Buffer.from('legacy file'), 'Old document.doc')).toBe('UNSUPPORTED_FILE');
  await expect(page.getByRole('alert')).toContainText('Renaming the extension does not convert it');
  const ole = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  expect(await open(page, ole, 'Disguised.docx')).toBe('UNSUPPORTED_FILE');
  await expect(page.getByRole('alert')).toContainText('legacy binary or encrypted Word file');
  await expect(sheet(page)).toContainText('Keep this draft');
});

test('preserves exact line spacing and untouched indentation attributes when editing related properties', async ({ page }) => {
  const original = await patchFixture(xml => xml.replace('<w:body>', `<w:body><w:p><w:pPr>
    <w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="exact" w:afterAutospacing="1"/>
    <w:ind w:left="100" w:start="150" w:right="200" w:endChars="300" w:hanging="180" w:hangingChars="100"/>
    </w:pPr><w:r><w:t>Property preservation</w:t></w:r></w:p>`));
  const result = await page.evaluate(async bytes => {
    const readPath = '/src/docx-read.ts', writePath = '/src/docx-write.ts';
    const { importDocx } = await import(/* @vite-ignore */ readPath) as typeof import('../src/docx-read');
    const { exportDocx } = await import(/* @vite-ignore */ writePath) as typeof import('../src/docx-write');
    const { model, source } = await importDocx(new File([new Uint8Array(bytes)], 'preserve.docx'));
    const blocks = model.blocks.map(block => block.kind === 'paragraph' && block.inlines.some(run => run.kind === 'text' && run.text === 'Property preservation')
      ? { ...block, style: { ...block.style, before: 24, indent: 48 } } : block);
    return Array.from(new Uint8Array(await (await exportDocx({ ...model, blocks }, 'preserve.docx', source)).arrayBuffer()));
  }, Array.from(original));
  const parts = await readZip(new Uint8Array(result));
  const attributes = await page.evaluate(xml => {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const p = Array.from(doc.getElementsByTagNameNS(ns, 'p')).find(node => Array.from(node.getElementsByTagNameNS(ns, 't'))
      .some(text => text.textContent === 'Property preservation'))!;
    return Object.fromEntries(['spacing', 'ind'].map(name => [name, Object.fromEntries(Array.from(p.getElementsByTagNameNS(ns, name)[0].attributes)
      .map(attribute => [attribute.localName, attribute.value]))]));
  }, Buffer.from(parts.get('word/document.xml')!).toString());
  expect(attributes.spacing).toMatchObject({ before: '360', after: '240', line: '360', lineRule: 'exact', afterAutospacing: '1' });
  expect(attributes.ind).toMatchObject({ left: '720', right: '200', endChars: '300', hanging: '180', hangingChars: '100' });
  expect(attributes.ind).not.toHaveProperty('start');
  expect(await open(page, new Uint8Array(result))).toBe('OK');
  await expect(sheet(page)).toContainText('Property preservation');
});

test('registers a navigation guard only while there are unsaved edits', async ({ page }) => {
  const blocked = () => page.evaluate(() => !window.dispatchEvent(new Event('beforeunload', { cancelable: true })));
  expect(await blocked()).toBe(false);
  await firstParagraph(page).click();
  await page.keyboard.type('Unsaved draft');
  expect(await blocked()).toBe(true);
  await page.keyboard.press('Control+z');
  await expect(page.locator('.rwe-local-state')).toHaveText('Local document');
  expect(await blocked()).toBe(false);
});

test('pastes a styled Word text fragment into the current sentence without splitting it', async ({ page }) => {
  await firstParagraph(page).click();
  await page.keyboard.type('Before after');
  await select(page, 'Before after', 7, 'Before after', 7);
  await pasteHtml(page, '<p class="MsoNormal" style="margin-top:0pt;margin-bottom:8pt;text-align:right"><b>Word </b></p>', 'Word ');
  await expect(sheet(page).locator('[data-paragraph]')).toHaveCount(1);
  await expect(firstParagraph(page)).toHaveText('Before Word after');
  await expect(sheet(page).getByText('Word ', { exact: true })).toHaveCSS('font-weight', '700');
  await page.keyboard.press('Control+z');
  await expect(firstParagraph(page)).toHaveText('Before after');
});

test('merge formatting uses the destination font and size while keeping pasted emphasis', async ({ page }) => {
  await firstParagraph(page).click();
  await page.getByLabel('Font family', { exact: true }).selectOption('Georgia');
  await page.getByLabel('Font size', { exact: true }).selectOption('18');
  await page.keyboard.type('Destination ');
  await page.getByLabel('Paste formatting').selectOption('merge');
  await pasteHtml(page, '<p><b style="font-family:Arial;font-size:9pt;color:#b00020">Merged emphasis</b></p>', 'Merged emphasis');
  const inserted = sheet(page).getByText('Merged emphasis', { exact: true });
  await expect(inserted).toHaveCSS('font-family', /Georgia/);
  await expect(inserted).toHaveCSS('font-size', '24px');
  await expect(inserted).toHaveCSS('font-weight', '700');
  await expect(inserted).toHaveCSS('color', 'rgb(36, 39, 45)');
  expect(await open(page, await exportBytes(page))).toBe('OK');
  await expect(inserted).toHaveCSS('font-size', '24px');
});

test('copies and cuts embedded pictures through the system clipboard without losing their position', async ({ page, context }) => {
  await enableClipboard(context);
  await firstParagraph(page).click();
  await pasteHtml(page, `<p>Before<img alt="Round-trip image" width="64" height="32" src="data:image/png;base64,${png.toString('base64')}">After</p>`);
  await select(page, 'Before', 0, 'After');
  await page.keyboard.press('Control+x');
  await expect(sheet(page).getByRole('img')).toHaveCount(0);
  const html = await page.evaluate(async () => (await (await navigator.clipboard.read())[0].getType('text/html')).text());
  expect(html).toContain('data:image/png;base64,');
  expect(html.indexOf('Before')).toBeLessThan(html.indexOf('<img'));
  expect(html.indexOf('<img')).toBeLessThan(html.indexOf('After'));
  await page.keyboard.press('Control+v');
  await expect(sheet(page).getByRole('img', { name: 'Round-trip image' })).toBeVisible();
  await expect(firstParagraph(page)).toHaveText('BeforeAfter');
  await page.keyboard.press('Control+z');
  await expect(sheet(page).getByRole('img')).toHaveCount(0);
  await page.keyboard.press('Control+z');
  await expect(sheet(page).getByRole('img', { name: 'Round-trip image' })).toBeVisible();
});

test('writes an actual PNG representation for a picture-only Word selection', async ({ page, context }) => {
  await enableClipboard(context);
  await firstParagraph(page).click();
  await pasteHtml(page, `<p><img alt="Copy PNG" src="data:image/png;base64,${png.toString('base64')}"></p>`);
  await sheet(page).locator('[data-image]').evaluate(image => {
    const range = document.createRange(); range.selectNode(image);
    document.getSelection()!.removeAllRanges(); document.getSelection()!.addRange(range);
  });
  await page.keyboard.press('Control+c');
  await expect(page.getByRole('status').filter({ hasText: 'Copied with embedded pictures' })).toBeVisible();
  const bytes = await page.evaluate(async () => Array.from(new Uint8Array(await (await (await navigator.clipboard.read())[0].getType('image/png')).arrayBuffer())));
  expect(Buffer.from(bytes).subarray(0, 8)).toEqual(png.subarray(0, 8));
});

test('saves and restores a local recovery copy with images and dirty state after reload', async ({ page }) => {
  await page.goto('/tests/host.html?recovery=recover-one');
  await firstParagraph(page).click();
  await pasteHtml(page, `<p>Recover this draft<img alt="Recovered picture" src="data:image/png;base64,${png.toString('base64')}"></p>`);
  await expect(page.getByRole('status').filter({ hasText: 'Recovery copy saved' })).toBeVisible();
  await page.evaluate(() => window.editor!.flushRecovery());
  page.once('dialog', dialog => dialog.accept());
  await page.reload();
  await expect(page.getByRole('button', { name: 'Restore draft', exact: true })).toBeVisible();
  await expect(sheet(page)).not.toContainText('Recover this draft');
  await page.getByRole('button', { name: 'Restore draft', exact: true }).click();
  await expect(sheet(page)).toContainText('Recover this draft');
  await expect(sheet(page).getByRole('img', { name: 'Recovered picture' })).toBeVisible();
  await expect(page.locator('.rwe-local-state')).toHaveText('Unsaved changes');
  const event = page.waitForEvent('download'); await download(page).click(); await event;
  await expect(page.locator('.rwe-local-state')).toHaveText('Local document');
  await expect(page.getByRole('status').filter({ hasText: 'Recovery copy saved' })).toHaveCount(0);
  await page.reload();
  await expect(page.getByRole('button', { name: 'Restore draft', exact: true })).toHaveCount(0);
});

test('isolates recovery keys and requires explicit discard before replacing an existing recovery copy', async ({ page }) => {
  await page.goto('/tests/host.html?recovery=first-document');
  await firstParagraph(page).click(); await page.keyboard.type('First document draft');
  await expect(page.getByRole('status').filter({ hasText: 'Recovery copy saved' })).toBeVisible();
  page.once('dialog', dialog => dialog.accept());
  await page.goto('/tests/host.html?recovery=second-document');
  await expect(page.getByRole('button', { name: 'Restore draft', exact: true })).toHaveCount(0);
  await page.goto('/tests/host.html?recovery=first-document');
  await expect(page.getByRole('button', { name: 'Restore draft', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Discard recovery copy', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Restore draft', exact: true })).toHaveCount(0);
  await page.reload();
  await expect(page.getByRole('button', { name: 'Restore draft', exact: true })).toHaveCount(0);
});

test('preserves pasted table widths, borders, padding and alignment through DOCX', async ({ page }) => {
  await firstParagraph(page).click();
  await pasteHtml(page, `<table align="center" style="width:75%"><tr style="height:54px">
    <td style="width:25%;border:1.5pt solid #b00020;padding:9px 12px;vertical-align:middle;text-align:right">Quarter</td>
    <td style="width:75%;border:2pt dashed #0055aa;padding:6px;vertical-align:bottom">Rest</td></tr></table>`);
  const cell = sheet(page).getByRole('cell', { name: 'Quarter', exact: true });
  await expect(cell).toHaveCSS('border-top-color', 'rgb(176, 0, 32)');
  await expect(cell).toHaveCSS('border-top-width', '2px');
  await expect(cell).toHaveCSS('padding-left', '12px');
  await expect(cell).toHaveCSS('vertical-align', 'middle');
  await expect(cell.locator('[data-paragraph]')).toHaveCSS('text-align', 'right');
  const ratio = await cell.evaluate(node => node.getBoundingClientRect().width / node.closest('table')!.getBoundingClientRect().width);
  expect(ratio).toBeCloseTo(.25, 1);
  expect(await open(page, await exportBytes(page))).toBe('OK');
  await expect(cell).toHaveCSS('border-top-width', '2px');
  await expect(cell).toHaveCSS('padding-left', '12px');
  await expect(sheet(page).getByRole('cell', { name: 'Rest', exact: true })).toHaveCSS('border-top-style', 'dashed');
});

test('renders exact and minimum Word line spacing and right-to-left paragraphs', async ({ page }) => {
  await firstParagraph(page).click();
  await pasteHtml(page, `<p dir="rtl" style="line-height:24pt">مرحبا بالعالم</p>
    <p style="line-height:30pt;mso-line-height-rule:at-least">Minimum height</p>`);
  const rtl = sheet(page).locator('[data-paragraph]').filter({ hasText: 'مرحبا بالعالم' });
  await expect(rtl).toHaveAttribute('dir', 'rtl');
  await expect(rtl).toHaveCSS('line-height', '32px');
  await expect(sheet(page).locator('[data-paragraph]').filter({ hasText: 'Minimum height' })).toHaveCSS('line-height', '40px');
  expect(await open(page, await exportBytes(page))).toBe('OK');
  await expect(rtl).toHaveAttribute('dir', 'rtl');
  await expect(rtl).toHaveCSS('line-height', '32px');
  const parts = await readZip(await exportBytes(page));
  const xml = Buffer.from(parts.get('word/document.xml')!).toString();
  expect(xml).toContain('w:lineRule="exact"'); expect(xml).toContain('w:lineRule="atLeast"');
});

test('keeps a Word heading with the following paragraph at a page boundary', async ({ page }) => {
  await firstParagraph(page).click();
  await pasteHtml(page, `<p style="line-height:24px">${Array.from({ length: 33 }, () => 'Filler').join('<br>')}</p>
    <p style="line-height:24px;page-break-after:avoid">Kept heading</p>
    <p style="line-height:24px;page-break-inside:avoid">Following paragraph<br>Second line</p>`);
  const headingPage = await sheet(page).getByText('Kept heading', { exact: true }).evaluate(node => node.closest('[data-page]')!.getAttribute('data-page'));
  const followingPage = await sheet(page).getByText('Following paragraph\nSecond line', { exact: true }).evaluate(node => node.closest('[data-page]')!.getAttribute('data-page'));
  expect(headingPage).toBe(followingPage);
  expect(headingPage).toBe('2');
});

test('uses the configured author for tracked changes and the comment form', async ({ page }) => {
  await page.goto('/tests/host.html?author=Document%20Reviewer');
  await firstParagraph(page).click(); await page.keyboard.type('Review me');
  await select(page, 'Review me', 0, 'Review me', 6);
  await page.getByRole('tab', { name: 'Review', exact: true }).click();
  await page.getByRole('button', { name: 'New comment', exact: true }).click();
  await expect(page.getByRole('dialog').getByLabel('Author', { exact: true })).toHaveValue('Document Reviewer');
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  await page.getByRole('button', { name: 'Track changes', exact: true }).click();
  await select(page, 'Review me', 9, 'Review me', 9); await page.keyboard.type('!');
  const parts = await readZip(await exportBytes(page));
  expect(Buffer.from(parts.get('word/document.xml')!).toString()).toContain('w:author="Document Reviewer"');
});

test('converts real DOC files to editable DOCX and exports legacy DOC again', async ({ page, request }) => {
  test.skip(!process.env.WORD_CONVERTER_TESTS, 'Requires the local Word conversion companion.');
  const response = await request.post('http://127.0.0.1:4175/convert?input=docx&format=doc', {
    data: await readFile(fixturePath), headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  });
  expect(response.ok()).toBe(true);
  const legacy = await response.body();
  expect(legacy.subarray(0, 8).toString('hex')).toBe('d0cf11e0a1b11ae1');
  await page.goto('/tests/host.html?convert');
  expect(await open(page, legacy, 'Legacy plan.doc')).toBe('OK');
  await expect(sheet(page)).toHaveAttribute('aria-readonly', 'false');
  await expect(sheet(page)).toContainText('Overview');
  const untouched = await page.evaluate(async () => Array.from(new Uint8Array(await (await window.editor!.exportDoc()).arrayBuffer())));
  expect(Buffer.from(untouched)).toEqual(legacy);
  await select(page, 'Overview', 0, 'Overview', 0); await page.keyboard.type('Edited ');
  const edited = await page.evaluate(async () => Array.from(new Uint8Array(await (await window.editor!.exportDoc()).arrayBuffer())));
  expect(Buffer.from(edited).subarray(0, 8).toString('hex')).toBe('d0cf11e0a1b11ae1');
  expect(await open(page, new Uint8Array(edited), 'Edited plan.doc')).toBe('OK');
  await expect(sheet(page)).toContainText('Edited Overview');
  await page.getByRole('tab', { name: 'File', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Download DOC', exact: true })).toBeVisible();
});

test('edits horizontal merged cells while preserving their native Word grid', async ({ page }) => {
  await firstParagraph(page).click();
  await pasteHtml(page, '<table><tr><td colspan="2">Merged title</td><td>Last column</td></tr><tr><td>One</td><td>Two</td><td>Three</td></tr></table>');
  const merged = sheet(page).getByRole('cell', { name: 'Merged title', exact: true });
  await expect(merged).toHaveAttribute('colspan', '2');
  await expect(sheet(page).getByRole('cell')).toHaveCount(5);
  expect(await open(page, await exportBytes(page))).toBe('OK');
  await expect(merged).toHaveAttribute('colspan', '2');
  await select(page, 'Merged title', 0, 'Merged title', 0); await page.keyboard.type('Edited ');
  await expect(sheet(page).getByRole('cell', { name: 'Edited Merged title', exact: true })).toHaveAttribute('colspan', '2');
  await expect(page.getByRole('button', { name: 'Format table', exact: true })).toBeDisabled();
  await expect(page.getByLabel('Table action', { exact: true })).toBeDisabled();
  expect(await open(page, await exportBytes(page))).toBe('OK');
  await expect(sheet(page).getByRole('cell', { name: 'Edited Merged title', exact: true })).toHaveAttribute('colspan', '2');
  await expect(sheet(page).getByRole('cell')).toHaveCount(5);
});

test('preserves letter and Roman list numbering through Word export', async ({ page }) => {
  await firstParagraph(page).click();
  await pasteHtml(page, '<ol type="A" start="26"><li>Letter Z</li><li>Letter AA</li></ol><ol type="i" start="4"><li>Roman four</li><li>Roman five</li></ol>');
  await expect(sheet(page).locator('.rwe-list-label')).toHaveText(['Z.', 'AA.', 'iv.', 'v.']);
  expect(await open(page, await exportBytes(page))).toBe('OK');
  await expect(sheet(page).locator('.rwe-list-label')).toHaveText(['Z.', 'AA.', 'iv.', 'v.']);
});

test('reports conversion failures without replacing unsaved content', async ({ page }) => {
  await page.goto('/tests/host.html?convert');
  await firstParagraph(page).click(); await page.keyboard.type('Protected draft');
  await page.route('**/word-converter/convert?*', route => route.fulfill({ status: 503, contentType: 'application/json', body: JSON.stringify({ message: 'Converter unavailable' }) }));
  expect(await open(page, Buffer.from('a legacy file'), 'broken.doc')).toBe('CONVERSION_FAILED');
  await expect(sheet(page)).toContainText('Protected draft');
  await expect(page.getByRole('alert')).toContainText('Converter unavailable');
});

test('retains exact line spacing while modifying a Word style and supports minimum spacing', async ({ page }) => {
  await firstParagraph(page).click(); await page.keyboard.type('Spacing example');
  await page.getByRole('button', { name: 'Modify style', exact: true }).click();
  let dialog = page.getByRole('dialog');
  await dialog.getByLabel('Line spacing rule', { exact: true }).selectOption('exact');
  await dialog.getByLabel('Line spacing', { exact: true }).fill('24');
  await dialog.getByRole('button', { name: 'Save style', exact: true }).click();
  await expect(firstParagraph(page)).toHaveCSS('line-height', '32px');
  expect(await open(page, await exportBytes(page))).toBe('OK');
  await select(page, 'Spacing example', 0, 'Spacing example', 0);
  await page.getByRole('button', { name: 'Modify style', exact: true }).click();
  dialog = page.getByRole('dialog');
  await expect(dialog.getByLabel('Line spacing rule', { exact: true })).toHaveValue('exact');
  await expect(dialog.getByLabel('Line spacing', { exact: true })).toHaveValue('24');
  await dialog.getByLabel('Space after (pt)', { exact: true }).fill('15');
  await dialog.getByRole('button', { name: 'Save style', exact: true }).click();
  await expect(firstParagraph(page)).toHaveCSS('line-height', '32px');
  await page.getByRole('button', { name: 'Modify style', exact: true }).click();
  dialog = page.getByRole('dialog');
  await dialog.getByLabel('Line spacing rule', { exact: true }).selectOption('atLeast');
  await dialog.getByLabel('Line spacing', { exact: true }).fill('30');
  await dialog.getByRole('button', { name: 'Save style', exact: true }).click();
  await expect(firstParagraph(page)).toHaveCSS('line-height', '40px');
  expect(await open(page, await exportBytes(page))).toBe('OK');
  await expect(firstParagraph(page)).toHaveCSS('line-height', '40px');
});

test('round-trips pictures and horizontal merges through legacy DOC conversion', async ({ page }) => {
  test.skip(!process.env.WORD_CONVERTER_TESTS, 'Requires the local Word conversion companion.');
  await page.goto('/tests/host.html?convert');
  await firstParagraph(page).click();
  await pasteHtml(page, `<p>Document with picture<img alt="Legacy image" src="data:image/png;base64,${png.toString('base64')}"></p>
    <table><tr><td colspan="2">Merged DOC cell</td></tr><tr><td>Left</td><td>Right</td></tr></table>`);
  const bytes = await page.evaluate(async () => Array.from(new Uint8Array(await (await window.editor!.exportDoc()).arrayBuffer())));
  expect(await open(page, new Uint8Array(bytes), 'Pictures and tables.doc')).toBe('OK');
  await expect(sheet(page).getByRole('img')).toHaveCount(1);
  await expect(sheet(page).getByRole('cell', { name: 'Merged DOC cell', exact: true })).toHaveAttribute('colspan', '2');
  await expect(sheet(page).getByRole('cell', { name: 'Left', exact: true })).toBeVisible();
  await expect(sheet(page).getByRole('cell', { name: 'Right', exact: true })).toBeVisible();
});

test('rejects malformed conversions at the companion boundary', async ({ request }) => {
  test.skip(!process.env.WORD_CONVERTER_TESTS, 'Requires the local Word conversion companion.');
  const invalid = await request.post('http://127.0.0.1:4175/convert?input=doc&format=docx', {
    data: 'not a binary Word file', headers: { 'Content-Type': 'application/msword' },
  });
  expect(invalid.status()).toBe(422);
  const crossSite = await request.post('http://127.0.0.1:4175/convert?input=doc&format=docx', {
    data: 'payload', headers: { 'Content-Type': 'text/plain', 'Sec-Fetch-Site': 'cross-site' },
  });
  expect(crossSite.status()).toBe(403);
});
