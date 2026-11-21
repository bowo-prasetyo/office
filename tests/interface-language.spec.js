// The interface in Indonesian (Bahasa Indonesia) or English: the browser's language chooses it, the
// start page's Language menu changes it (remembered in this browser), and only the interface
// changes: a document's text and the files saved are the same in both. A blank document started in
// Indonesian is in Indonesian (w:lang), so spelling is checked in it.
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import fs from 'node:fs';
import { buildDocx } from './helpers.js';

const RIBBON = { cookies: [], origins: [{ origin: 'http://localhost:4173', localStorage: [{ name: 'intact.ribbon', value: '1' }] }] };

test.describe('in English', () => {
  test.use({ locale: 'en-GB' });
  test('the interface is in English, as before', async ({ page }) => {
    await page.goto('./');
    await expect(page.getByRole('heading', { level: 1 })).toHaveText('Change what you need. Keep the rest intact.');
    await expect(page.getByRole('combobox', { name: 'Language' })).toHaveValue('en');
    await expect(page.locator('html')).toHaveAttribute('lang', 'en');
    await expect(page.getByRole('region', { name: 'Naskah dinas (Tata Naskah Dinas)' })).toHaveCount(0);
  });
});

test.describe('in Indonesian', () => {
  test.use({ locale: 'id-ID', storageState: RIBBON });

  test('the start page is in Indonesian; the Language menu switches to English, remembered', async ({ page }) => {
    await page.goto('./');
    await expect(page.getByRole('heading', { level: 1 })).toHaveText('Ubah yang perlu. Sisanya tetap utuh.');
    await expect(page.locator('html')).toHaveAttribute('lang', 'id');
    await expect(page.getByRole('button', { name: /Dokumen kosong/ })).toBeVisible();
    await page.getByRole('combobox', { name: 'Bahasa' }).selectOption('en');
    await expect(page.getByRole('heading', { level: 1 })).toHaveText('Change what you need. Keep the rest intact.');
    await page.reload();
    await expect(page.getByRole('heading', { level: 1 })).toHaveText('Change what you need. Keep the rest intact.');
    await expect(page.locator('html')).toHaveAttribute('lang', 'en');
    await page.getByRole('combobox', { name: 'Language' }).selectOption('id');
    await expect(page.getByRole('heading', { level: 1 })).toHaveText('Ubah yang perlu. Sisanya tetap utuh.');
  });

  test('the ribbon, its tools and the page dialogs are in Indonesian, with Word\'s Indonesian names', async ({ page }) => {
    await page.goto('./');
    await page.setInputFiles('input[type=file]', { name: 'surat.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await buildDocx('<w:p><w:r><w:t>Isi surat.</w:t></w:r></w:p>') });
    await expect(page.locator('.docx-page')).toBeVisible();
    await expect(page.getByRole('tab')).toHaveText(['File', 'Beranda', 'Sisipkan', 'Tata Letak', 'Referensi', 'Tinjau', 'Tampilan']);
    await expect(page.getByRole('button', { name: 'Tebal' })).toBeVisible();
    await expect(page.locator('.app-header').getByRole('button', { name: 'Unduh' })).toBeVisible();
    await expect(page.locator('.statusbar')).toContainText('2 kata · 10 karakter');
    await page.locator('.docx-page p').first().click();
    await page.getByRole('tab', { name: 'Tata Letak' }).click();
    await page.getByRole('button', { name: 'Pengaturan halaman' }).click();
    const setup = page.getByRole('dialog', { name: 'Pengaturan halaman' });
    await expect(setup.getByRole('combobox', { name: /^Ukuran/ })).toHaveValue('A4');
    await expect(setup.getByRole('option', { name: 'F4 (21.5 × 33 cm)' })).toHaveCount(1);
    await expect(setup.getByRole('option', { name: 'Ukuran kustom' })).toHaveCount(1);
    await setup.getByRole('button', { name: 'Batal' }).click();
    await page.locator('.docx-page p').first().click();
    await page.getByRole('button', { name: 'Pilihan nomor halaman dan header' }).click();
    await page.getByRole('menuitem', { name: /^Format nomor halaman/ }).click();
    const format = page.getByRole('dialog', { name: 'Format nomor halaman' });
    await expect(format.getByLabel('Lanjutkan dari bagian sebelumnya')).toBeChecked();
    await format.getByRole('combobox', { name: 'Format angka', exact: true }).selectOption('lowerRoman');
    await format.getByRole('button', { name: 'OK' }).click();
    await expect(page.locator('.statusbar [role=status]')).toHaveText('Nomor halaman: i, ii, iii untuk seluruh dokumen');
  });

  test('a blank document started in Indonesian is in Indonesian, named "Tanpa judul"', async ({ page }) => {
    await page.goto('./');
    await page.getByRole('button', { name: /Dokumen kosong/ }).click();
    await expect(page.getByRole('textbox', { name: 'Nama berkas' })).toHaveValue('Tanpa judul');
    await page.locator('.docx-page p').first().click();
    await page.keyboard.type('Surat ini dibuat di Intact.');
    await page.locator('.app-header').getByRole('button', { name: 'Unduh' }).click();
    const review = page.getByRole('dialog', { name: 'Review your changes' });
    const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download' }).click()]);
    expect(dl.suggestedFilename()).toBe('Tanpa judul.docx');
    const zip = await JSZip.loadAsync(fs.readFileSync(await dl.path()));
    expect(await zip.file('word/styles.xml').async('string')).toContain('<w:lang w:val="id-ID"/>');
  });
});
