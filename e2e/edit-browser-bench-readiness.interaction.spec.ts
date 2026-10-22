import { expect, test, type Page } from '@playwright/test';
import { loadHarness, twoFrames } from './edit-browser-bench-harness.js';

async function serveOpeningState(page: Page, attributes: string): Promise<void> {
  await page.route('http://localhost:5275/**', (route) =>
    route.fulfill({
      contentType: 'text/html',
      body: `<main ${attributes}><div class="docx-page" data-materialized="true" style="width:100px;height:100px"></div></main>`,
    })
  );
  await page.addInitScript(() => {
    // Only harness readiness is under test. The fixture does not replace the editor engine.
    window.__DOCX_EDITOR_E2E__ = {
      ready: () => true,
      fontMeasurer: () => 'shaped',
    } as NonNullable<typeof window.__DOCX_EDITOR_E2E__>;
  });
}

test('keeps measurements outside partially published pagination', async ({ page }) => {
  await serveOpeningState(page, 'data-docx-opening-preview data-docx-layout-pending');
  let installed = 0;
  const loading = loadHarness(page, async () => {
    installed += 1;
  });
  await expect(page.locator('.docx-page[data-materialized="true"]')).toHaveCount(1);
  await twoFrames(page);
  expect(installed).toBe(0);

  await page.evaluate(() =>
    document.querySelector('main')!.removeAttribute('data-docx-opening-preview')
  );
  await twoFrames(page);
  expect(installed).toBe(0);

  await page.evaluate(() =>
    document.querySelector('main')!.removeAttribute('data-docx-layout-pending')
  );
  await loading;
  expect(installed).toBe(1);
});

test('requires a materialized page after pagination settles', async ({ page }) => {
  await serveOpeningState(page, 'data-docx-layout-pending');
  let installed = 0;
  const loading = loadHarness(page, async () => {
    installed += 1;
  });
  await expect(page.locator('[data-docx-layout-pending]')).toHaveCount(1);
  await page.evaluate(() => {
    document.querySelector('.docx-page')!.removeAttribute('data-materialized');
    document.querySelector('main')!.removeAttribute('data-docx-layout-pending');
  });
  await twoFrames(page);
  expect(installed).toBe(0);

  await page.evaluate(() =>
    document.querySelector('.docx-page')!.setAttribute('data-materialized', 'true')
  );
  await loading;
  expect(installed).toBe(1);
});

test('also waits for complete layout without measurement installation', async ({ page }) => {
  await serveOpeningState(page, 'data-docx-opening-preview');
  let completed = false;
  let installed = 0;
  const loading = loadHarness(
    page,
    async () => {
      installed += 1;
    },
    false
  ).then(() => {
    completed = true;
  });
  await expect(page.locator('[data-docx-opening-preview]')).toHaveCount(1);
  await twoFrames(page);
  expect(completed).toBe(false);
  expect(installed).toBe(0);

  await page.evaluate(() =>
    document.querySelector('main')!.removeAttribute('data-docx-opening-preview')
  );
  await loading;
  expect(completed).toBe(true);
  expect(installed).toBe(0);
});

test('keeps a hidden materialized page outside measurements', async ({ page }) => {
  await serveOpeningState(page, 'data-docx-layout-pending style="visibility:hidden"');
  let installed = 0;
  const loading = loadHarness(page, async () => {
    installed += 1;
  });
  await expect(page.locator('[data-docx-layout-pending]')).toHaveCount(1);
  await page.evaluate(() => {
    const materialized = document.querySelector<HTMLElement>('.docx-page')!;
    materialized.style.visibility = 'hidden';
    document.querySelector('main')!.removeAttribute('data-docx-layout-pending');
  });
  await twoFrames(page);
  expect(installed).toBe(0);
  await page.evaluate(() => {
    document.querySelector<HTMLElement>('.docx-page')!.style.visibility = 'visible';
  });
  await loading;
  expect(installed).toBe(1);
});
