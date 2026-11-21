// On a phone: the most used tools come first in the toolbar, and what just happened shows in the
// status bar. Share… (where the browser can share files) hands the saved file to the share sheet.
// Insert: a check box (Word's check box content control) or today's date at the caret.
import { test, expect, devices } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor, download } from './helpers.js';

const TEXT = '<w:p><w:r><w:t xml:space="preserve">Bring a packed lunch. </w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Date: </w:t></w:r></w:p>';
const status = (page) => page.locator('.statusbar [role=status]');

/** A fake share sheet: records what it was given, and answers as `result` says. */
function fakeShare(page, result = 'ok') {
  return page.addInitScript((r) => {
    window.__shared = [];
    navigator.canShare = (data) => !!data?.files?.length;
    navigator.share = async (data) => {
      const file = data.files[0];
      const bytes = Array.from(new Uint8Array(await file.arrayBuffer()));
      window.__shared.push({ name: file.name, type: file.type, title: data.title, bytes });
      if (window.__shareAnswer === 'denied-once') {
        window.__shareAnswer = 'ok';
        throw new DOMException('Not allowed', 'NotAllowedError');
      }
      if (r === 'abort') throw new DOMException('Cancelled', 'AbortError');
    };
    window.__shareAnswer = r;
  }, result);
}

/** Places the caret at the end of paragraph `i`. */
async function caretAtEnd(page, i) {
  await page.locator('.docx-page p').nth(i).click();
  await page.keyboard.press('End');
}

// ------------------------------------------------------------------ share

test('Share… goes through the review, then hands the .docx to the share sheet', async ({ page }) => {
  await fakeShare(page);
  await openInEditor(page, await buildDocx(TEXT), 'trip.docx');
  await caretAtEnd(page, 0);
  await page.keyboard.type('Thanks!');
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menuitem', { name: /Share…/ }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await review.getByRole('button', { name: 'Share…' }).click();
  await expect(status(page)).toContainText('Shared trip.docx');
  const shared = await page.evaluate(() => window.__shared);
  expect(shared).toHaveLength(1);
  expect(shared[0].name).toBe('trip.docx');
  expect(shared[0].type).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  const zip = await JSZip.loadAsync(Buffer.from(shared[0].bytes));
  expect(await zip.file('word/document.xml').async('string')).toContain('Bring a packed lunch. Thanks!');
  // Sent on: no longer "unsaved", as after Download.
  await expect(page.locator('.app-header .dirty')).toHaveCount(0);
});

test('when the browser wants another tap before sharing, one button finishes it', async ({ page }) => {
  await fakeShare(page, 'denied-once');
  await openInEditor(page, await buildDocx(TEXT), 'trip.docx');
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menuitem', { name: /Share…/ }).click();
  await page.getByRole('dialog', { name: 'Review your changes' }).getByRole('button', { name: 'Share…' }).click();
  const ready = page.getByRole('dialog', { name: 'Share' });
  await expect(ready).toContainText('trip.docx is ready.');
  await ready.getByRole('button', { name: 'Share…' }).click();
  await expect(ready).toBeHidden();
  await expect(status(page)).toContainText('Shared trip.docx');
  expect(await page.evaluate(() => window.__shared.length)).toBe(2);
});

test('cancelling the share sheet keeps the changes unsaved and says so', async ({ page }) => {
  await fakeShare(page, 'abort');
  await openInEditor(page, await buildDocx(TEXT), 'trip.docx');
  await caretAtEnd(page, 0);
  await page.keyboard.type('x');
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menuitem', { name: /Share…/ }).click();
  await page.getByRole('dialog', { name: 'Review your changes' }).getByRole('button', { name: 'Share…' }).click();
  await expect(status(page)).toContainText('trip.docx wasn’t shared');
  await expect(page.locator('.app-header .dirty')).toHaveCount(1);
});

test('without a share sheet for files, Share… isn’t offered', async ({ page }) => {
  await page.addInitScript(() => { navigator.canShare = () => false; navigator.share = async () => {}; });
  await openInEditor(page, await buildDocx(TEXT));
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await expect(page.getByRole('menuitem', { name: /PDF/ })).toBeVisible();
  await expect(page.getByRole('menuitem', { name: /Share/ })).toHaveCount(0);
});

// ------------------------------------------------------------------ insert

test('Insert puts a check box at the caret, which ticks with a click and saves as Word’s', async ({ page }) => {
  await openInEditor(page, await buildDocx(TEXT));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('Home');
  await page.getByRole('button', { name: 'Insert a check box or date' }).click();
  await page.getByRole('menuitem', { name: /Check box/ }).click();
  await expect(status(page)).toContainText('Check box inserted');
  const box = page.locator('.docx-page p').first().locator('.sdt-checkbox');
  await expect(box).toHaveText('☐');
  await box.click();
  await expect(box).toHaveText('☒');

  const { xml } = await download(page);
  expect(xml).toMatch(/<w:p><w:sdt><w:sdtPr><w:id w:val="-?\d+"\/><w14:checkbox[^>]*><w14:checked w14:val="1"\/><w14:checkedState w14:val="2612" w14:font="MS Gothic"\/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"\/><\/w14:checkbox><\/w:sdtPr><w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"\/><\/w:rPr><w:t[^>]*>☒<\/w:t><\/w:r><\/w:sdtContent><\/w:sdt><w:r><w:t xml:space="preserve">Bring a packed lunch. <\/w:t>/);

  await page.locator('.docx-page').focus();
  await page.keyboard.press('Control+z');
  await page.keyboard.press('Control+z');
  await expect(page.locator('.docx-page .sdt-checkbox')).toHaveCount(0);
});

test('a check box in a document that doesn’t declare Word 2010’s namespace still saves well-formed', async ({ page }) => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Milk</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  await openInEditor(page, await zip.generateAsync({ type: 'nodebuffer' }));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('Home');
  await page.getByRole('button', { name: 'Insert a check box or date' }).click();
  await page.getByRole('menuitem', { name: /Check box/ }).click();
  const { xml } = await download(page);
  expect(xml).toMatch(/<w14:checkbox[^>]*xmlns:w14="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordml"|xmlns:w14="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordml"[^>]*>(?:(?!<w14:checkbox).)*<w14:checkbox/s);
  expect(await page.evaluate((x) => !new DOMParser().parseFromString(x, 'application/xml').querySelector('parsererror'), xml)).toBe(true);
});

test('Insert offers today’s date in the text’s language and types it at the caret', async ({ page }) => {
  await page.clock.setFixedTime(new Date('2026-11-19T10:00:00'));
  await openInEditor(page, await buildDocx(TEXT));
  await caretAtEnd(page, 1);
  await page.getByRole('button', { name: 'Insert a check box or date' }).click();
  const menu = page.getByRole('menu', { name: 'Insert' });
  await expect(menu.getByRole('menuitem')).toHaveText([/Check box/, 'November 19, 2026', 'Thursday, November 19, 2026', '11/19/2026', '2026-11-19',
    /^Answer line/, /^Line/, /^Arrow/, /^Text box/, /^Table of contents/, /^Caption/, /^Cross-reference/, /^Equation/, /^Citation/, /^Bibliography/]);
  await menu.getByRole('menuitem', { name: 'November 19, 2026', exact: true }).click();
  await expect(page.locator('.docx-page p').nth(1)).toHaveText('Date: November 19, 2026');
  const { xml } = await download(page);
  expect(xml).toMatch(/Date: November 19, 2026<\/w:t>/);
});

test('a date follows the text’s language', async ({ page }) => {
  await page.clock.setFixedTime(new Date('2026-11-19T10:00:00'));
  const body = '<w:p><w:r><w:rPr><w:lang w:val="de-DE"/></w:rPr><w:t xml:space="preserve">Datum: </w:t></w:r></w:p>';
  await openInEditor(page, await buildDocx(body));
  await caretAtEnd(page, 0);
  await page.getByRole('button', { name: 'Insert a check box or date' }).click();
  await expect(page.getByRole('menu', { name: 'Insert' }).getByRole('menuitem').nth(1)).toHaveText('19. November 2026');
});

// ------------------------------------------------------------------ phone

test.describe('on a phone', () => {
  test.use({ viewport: devices['iPhone 13'].viewport, hasTouch: true });

  test('the toolbar starts with the tools a phone is used for, and messages come first in the status bar', async ({ page }) => {
    await openInEditor(page, await buildDocx(TEXT));
    const shown = await page.evaluate(() => [...document.querySelectorAll('.toolbar button, .toolbar select')]
      .filter((e) => { const r = e.getBoundingClientRect(); return r.width > 0 && r.right <= innerWidth; })
      .sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left)
      .map((e) => e.getAttribute('aria-label') || e.title));
    expect(shown.slice(0, 5)).toEqual(['Undo', 'Redo', 'Bold', 'Italic', 'Underline']);
    // The arrow keys follow what is shown.
    await page.locator('.docx-page p').first().click();
    await page.keyboard.press('Alt+F10');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowRight');
    await expect(page.getByRole('button', { name: 'Bold' })).toBeFocused();
    await page.keyboard.press('Escape');

    await expect(page.locator('.statusbar .zoom')).toBeHidden();
    await page.getByRole('button', { name: 'Insert a check box or date' }).click();
    await page.getByRole('menuitem', { name: /Check box/ }).click();
    const message = status(page);
    await expect(message).toContainText('Check box inserted');
    const box = await message.boundingBox();
    const words = await page.locator('.statusbar > span').first().boundingBox();
    expect(box.x).toBeLessThan(words.x); // first in the line, in view
    expect(box.x + box.width).toBeLessThanOrEqual(390);
    expect(await page.locator('.statusbar').evaluate((el) => el.offsetHeight)).toBeLessThan(30); // still one line
    expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(390);
  });

  // Approving a document between meetings: with tracked changes (or comments) to go through, the
  // review's tools come first, after Undo and Redo, with Accept and Reject named; and it says so.
  test('a document with tracked changes puts Accept and Reject first, and says so', async ({ page }) => {
    const tracked = '<w:p><w:r><w:t xml:space="preserve">The fee is </w:t></w:r>'
      + '<w:del w:id="1" w:author="Legal" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>200</w:delText></w:r></w:del>'
      + '<w:ins w:id="2" w:author="Legal" w:date="2026-01-01T00:00:00Z"><w:r><w:t>250</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> euros.</w:t></w:r></w:p>';
    await openInEditor(page, await buildDocx(TEXT + tracked), 'Contract.docx');
    await expect(status(page)).toHaveText('2 tracked changes to review: Accept and Reject are first in the toolbar');
    const shown = await page.evaluate(() => [...document.querySelectorAll('.toolbar button, .toolbar select')]
      .filter((e) => { const r = e.getBoundingClientRect(); return r.width > 0 && r.right <= innerWidth; })
      .sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left)
      .map((e) => e.getAttribute('aria-label') || e.title));
    expect(shown.slice(0, 5)).toEqual(['Undo', 'Redo', 'Previous change', 'Next change', 'Accept the change (or the changes selected) and go to the next']);
    const accept = page.locator('.changes-group .labelled').first();
    await expect(accept).toContainText('Accept'); // named, not just a tick
    // The arrow keys follow what is shown.
    await page.locator('.docx-page p').first().click();
    await page.keyboard.press('Alt+F10');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowRight');
    await expect(page.getByRole('button', { name: 'Previous change' })).toBeFocused();
    await page.keyboard.press('Escape');
    // Next, then Accept: through the changes with a thumb.
    await page.getByRole('button', { name: 'Next change' }).click();
    await accept.click();
    await accept.click();
    await expect(page.locator('.changes-group .badge')).toHaveText('0');
    await expect(page.locator('.docx-page')).toContainText('The fee is 250 euros.');
    // Nothing left to review: the editing tools come first again.
    const first = await page.evaluate(() => [...document.querySelectorAll('.toolbar button')]
      .filter((e) => { const r = e.getBoundingClientRect(); return r.width > 0 && r.right <= innerWidth; })
      .sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left)
      .map((e) => e.getAttribute('aria-label') || e.title).slice(0, 3));
    expect(first).toEqual(['Undo', 'Redo', 'Bold']);
  });

  test('the header leaves the file’s name room: Open and Close are icons, named for screen readers', async ({ page }) => {
    await openInEditor(page, await buildDocx(TEXT), 'Quarterly report.docx');
    await expect(page.locator('.app-header').getByRole('button', { name: 'Close' })).toBeVisible();
    await expect(page.locator('.app-header .close-doc .label')).toBeHidden();
    const name = await page.locator('.app-header .filename').boundingBox();
    expect(name.width).toBeGreaterThan(110); // was about 60 pixels: "Quart"
    expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(390);
  });
});
