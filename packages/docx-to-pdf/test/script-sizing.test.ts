/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/docx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { exportPdf } from '../src/index.ts';
import { docx } from './fixture.ts';
import { readFileSync } from 'node:fs';
import { createFontSource } from '@docx-editor.dev/core/editor';

// Saved Word 16.113 demo: 11pt super/subscript emits 6.96pt; 10pt note markers emit
// 6.48pt. Allow one 300-DPI device step (0.24pt), because Core retains fractional sizes.
for (const [size, wordSize] of [
  [11, 6.96],
  [10, 6.48],
] as const) {
  for (const verticalAlign of size === 11 ? ['superscript', 'subscript'] : ['superscript']) {
    test(`${size}pt ${verticalAlign} stays within one device step of the saved Word sizing`, async () => {
      const run = (text: string, extra = '') =>
        `<w:r><w:rPr><w:sz w:val="${size * 2}"/>${extra}</w:rPr><w:t>${text}</w:t></w:r>`;
      const input = docx(
        `<w:p>${run('BASE')}${run('SCRIPT', `<w:vertAlign w:val="${verticalAlign}"/>`)}${run('END')}</w:p>`
      );
      const result = await exportPdf(input, { useSystemFonts: false });
      const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
      try {
        const content = await (await pdf.getPage(1)).getTextContent();
        const items = content.items.filter((item) => 'str' in item);
        const script = items.find((item) => item.str === 'SCRIPT')!;
        const end = items.find((item) => item.str === 'END')!;
        expect(script).toBeDefined();
        expect(end).toBeDefined();
        expect(Math.abs(script.transform[3] - wordSize)).toBeLessThanOrEqual(0.24);
        // The next run begins where layout put it, within one device step of the painted
        // script end. It cannot be exact: the drawn size is rounded onto the 0.24pt grid the
        // reference emits sizes on, while the advances keep the size layout measured with.
        expect(Math.abs(end.transform[4] - (script.transform[4] + script.width))).toBeLessThan(
          0.24
        );
      } finally {
        await pdf.destroy();
      }
    });
  }
}

test('font-derived script size reaches PDF text state without changing nominal size or baseline placement', async () => {
  const bytes = new Uint8Array(
    readFileSync(
      new URL('../../core/src/layout/__tests__/fixtures/fonts/DejaVuSans.ttf', import.meta.url)
    )
  );
  const admitted = createFontSource(bytes, { family: 'DejaVu Sans', weight: 400, style: 'normal' });
  if ('failure' in admitted) throw new Error(JSON.stringify(admitted.failure));
  const run = (text: string, script = false) =>
    `<w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="28"/>${script ? '<w:vertAlign w:val="subscript"/>' : ''}</w:rPr><w:t>${text}</w:t></w:r>`;
  const input = docx(`<w:p>${run('BASE')}${run('123', true)}${run('END')}</w:p>`);
  const original = input.slice();
  const states = [];
  for (const mode of ['legacy', 'font-metrics'] as const) {
    const result = await exportPdf(input, {
      useSystemFonts: false,
      fonts: {
        sources: [admitted.source],
        defaultFont: { family: 'DejaVu Sans', sizeHalfPoints: 28 },
      },
      fontShaping: { scriptGlyphSizing: mode },
    });
    expect(result.diagnostics.filter((entry) => entry.severity !== 'information')).toEqual([]);
    const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
    try {
      const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(
        (item) => 'str' in item
      );
      const script = items.find((item) => item.str === '123')!;
      const end = items.find((item) => item.str === 'END')!;
      expect(script).toBeDefined();
      expect(end).toBeDefined();
      // The fixture has UPEM 2048 and script Y size 1433. Keep the existing half-point and 0.24pt device rounding.
      expect(script.transform[3]).toBeCloseTo(mode === 'legacy' ? 9.12 : 10.08, 6);
      expect(Math.abs(end.transform[4] - (script.transform[4] + script.width))).toBeLessThan(0.24);
      states.push({ baseline: script.transform[5], end: end.transform[4] });
    } finally {
      await pdf.destroy();
    }
  }
  expect(states[0]!.baseline).toBe(states[1]!.baseline);
  expect(states[0]!.end).not.toBe(states[1]!.end);
  expect(input).toEqual(original);
});
