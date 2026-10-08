import { expect, test } from '@playwright/test';
import {
  DROPDOWN_SOURCE,
  DROPDOWN_SLOTS,
  POPUP_SELECTOR,
} from '../packages/vue/test/helpers/dropdown-document';

test.use({ viewport: { width: 2200, height: 1200 } });

for (const [adapter, port] of [
  ['react', 5273],
  ['vue', 5274],
] as const) {
  test(`${adapter}: mouse-opened dropdowns close on document Escape`, async ({
    page,
  }, testInfo) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.route('**/dropdown-escape.docx', (route) =>
      route.fulfill({
        body: Buffer.from(DROPDOWN_SOURCE),
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      })
    );
    await page.goto(`http://localhost:${port}/?fixture=dropdown-escape.docx`);
    const pages = page.locator('.docx-pages').first();
    try {
      await expect(pages).toContainText('First cell', { timeout: 45_000 });
    } finally {
      if (errors.length)
        await testInfo.attach('browser-errors', {
          body: errors.join('\n'),
          contentType: 'text/plain',
        });
    }
    await pages.getByText('First cell', { exact: true }).click();
    const before = await pages.textContent();
    for (const slot of DROPDOWN_SLOTS) {
      await pages.getByText('First cell', { exact: true }).click();
      const root = page.locator(`.docx-toolbar [data-slot="${slot}"]`).first();
      const trigger = root.locator('[aria-haspopup]').first();
      await expect(trigger).toBeEnabled();
      await trigger.click();
      await expect(root.locator(POPUP_SELECTOR)).toBeVisible();
      await pages.focus();
      await pages.dispatchEvent('keydown', { key: 'Escape', isComposing: true, bubbles: true });
      await expect(root.locator(POPUP_SELECTOR)).toBeVisible();
      await page.keyboard.press('Escape');
      await expect(root.locator(POPUP_SELECTOR)).toHaveCount(0);
      await expect(pages).toBeFocused();
    }
    expect(await pages.textContent()).toBe(before);
  });
}
