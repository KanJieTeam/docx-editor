/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/docx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { exportPdf, exportPdfFrom, openDocumentForExport } from '../src/index.ts';
import { snapToGrid } from '../src/text.ts';
import { docx } from './fixture.ts';

// An exact half-unit on the 0.24pt device grid rounds toward the page TOP.
//
// Ties are not rare: any exact line spacing that is an odd multiple of 0.12pt hits one on
// every other line. Here `w:top="2880"` puts the story origin at 144pt — 600 units exactly —
// and 15pt lines advance 62.5 units. The in-line baseline is gridded first, so every
// other painted baseline lands on an exact .5 regardless of the font's ascent.
// A captured control of this shape shows the reference taking the LOWER unit for all six of
// them; `Math.round` took the upper one every time, 0.24pt low on the page.
// See `.cache/pdf/claude-linerule/`.
const GRID = 0.24;
const LINES = 12;

test('an exact half-unit baseline rounds toward the page top', async () => {
  const runs = Array.from(
    { length: LINES },
    (_, i) => `<w:r>${i ? '<w:br/>' : ''}<w:t xml:space="preserve">T${i}</w:t></w:r>`
  ).join('');
  const input = docx(
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/></w:pPr>${runs}</w:p>` +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="2880" w:right="1440" w:bottom="1440" w:left="1440"' +
      ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
  );
  const result = await exportPdf(input, { useSystemFonts: false });
  const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
  try {
    const page = await pdf.getPage(1);
    const height = page.view[3]!;
    const content = await page.getTextContent();
    const tops = content.items
      .filter((item): item is Extract<typeof item, { str: string }> => 'str' in item)
      .filter((item) => /^T\d+$/.test(item.str))
      .map((item) => Number((height - item.transform[5]!).toFixed(4)))
      .sort((a, b) => a - b);
    expect(tops).toHaveLength(LINES);
    // Every painted baseline sits on the grid, and the ties took the lower unit.
    for (const top of tops)
      expect(Math.abs(top / GRID - Math.round(top / GRID))).toBeLessThan(1e-6);
    // Exact spacing centres the face's natural band; it does not prescribe a 12pt
    // baseline. Test the tie direction independently of that in-line offset.
    const firstUnits = tops[0]! / GRID;
    expect(firstUnits).toBeCloseTo(Math.round(firstUnits), 8);
    const advances = tops.map((top) => Number((top - tops[0]!).toFixed(4)));
    const expected = Array.from({ length: LINES }, (_, i) =>
      Number((Math.ceil((i * 15) / GRID - 0.5) * GRID).toFixed(4))
    );
    expect(advances).toEqual(expected);
    expect(advances[1]).toBe(14.88);
    // Half of them ARE ties, or the fixture stopped exercising the rule.
    const ties = expected.filter((_, i) => Math.abs((((i * 15) / GRID) % 1) - 0.5) < 1e-9);
    expect(ties).toHaveLength(6);
  } finally {
    await pdf.destroy();
  }
});

test('font-backed exact lines centre the natural band and export that baseline', async () => {
  const paragraph = (label: string, line: string) =>
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0" ${line}/></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>` +
    `<w:sz w:val="24"/></w:rPr><w:t>${label}</w:t></w:r></w:p>`;
  const input = docx(
    paragraph('Natural', 'w:line="240" w:lineRule="auto"') +
      paragraph('Exact', 'w:line="480" w:lineRule="exact"')
  );
  const opened = await openDocumentForExport(input, { useSystemFonts: false });
  if (!opened.ok) throw new Error('Cannot open the exact-line fixture');
  try {
    const layout = await opened.session.layout();
    const page = layout.pages[0]!;
    const paragraphs = page.fragments.filter((fragment) => fragment.kind === 'paragraph');
    expect(paragraphs).toHaveLength(2);
    const natural = paragraphs[0]!.lines[0]!;
    const exact = paragraphs[1]!.lines[0]!;
    expect(natural.box.height).toBeLessThan(24);
    expect(exact.box.height).toBe(24);
    expect(exact.baseline - natural.baseline).toBeCloseTo((24 - natural.box.height) / 2, 8);
    expect(exact.box.y).toBeCloseTo(natural.box.y + natural.box.height, 8);
    const result = await exportPdfFrom(opened.session);
    const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
    try {
      const pdfPage = await pdf.getPage(1);
      const content = await pdfPage.getTextContent();
      const item = content.items.find((item) => 'str' in item && item.str === 'Exact');
      if (!item || !('str' in item)) throw new Error('The exact line is missing from the PDF');
      // This taller line rounds its in-line baseline, then grids the page position.
      const lineTop = page.contentBox.y + exact.box.y;
      const inLine = Math.round(exact.baseline / GRID) * GRID;
      const expected = Math.ceil((lineTop + inLine) / GRID - 0.5) * GRID;
      expect(pdfPage.view[3]! - item.transform[5]!).toBeCloseTo(expected, 6);
    } finally {
      await pdf.destroy();
    }
  } finally {
    opened.session.dispose();
  }
});

// The same rule, on a tie the layout reaches by accumulation rather than by arithmetic.
//
// A position arrives here as a sum of points that are exact in decimal and not in binary. In
// `issue-483-firstline-marker.docx` the thirty-sixth body baseline is mathematically 2149.5
// units, and the sum delivers 2149.5000000000014. `ceil(x - 1/2)` then takes 2150 instead of
// 2149 and paints that whole line 0.24pt low, which is most of that page's measured error.
// The drift is a part in 1e15, so nothing but a tie can see it.
test('a tie reached by accumulation resolves like an exact one', () => {
  const drifted = 515.88000000000033651 / GRID;
  expect(drifted).toBeGreaterThan(2149.5);
  expect(Math.ceil(drifted - 0.5)).toBe(2150);
  expect(Math.ceil(snapToGrid(drifted) - 0.5)).toBe(2149);
});

// Snapping must not move a value that is genuinely between two units, or every ordinary
// baseline would shift. Only a whole or half unit within a part in 1e9 is pulled back.
test('snapping leaves a position that is not near a tie alone', () => {
  for (const units of [2099.1916666666679, 2158.8083333333352, 0.3, -7.77, 1812.4]) {
    expect(snapToGrid(units)).toBe(units);
  }
  expect(snapToGrid(902.4999999999999)).toBe(902.5);
  expect(snapToGrid(340.00000000000006)).toBe(340);
});

// And the same rule end to end, on ties the layout REACHES rather than computes.
//
// `w:line="201"` advances 10.05pt. Twelve of those from a 72pt margin is mathematically a
// half unit, and the sum delivers 836.4999999999999; twenty is 1171.5000000000005. Drift in
// both directions, five times on one page. The expectation below stays in twips, where a twip
// is exactly 5/24 of a device unit, so the test cannot drift the way the code did.
const DRIFT_LINE_TWIPS = 201;
const DRIFT_TOP_TWIPS = 1440;
const DRIFT_LINES = 46;

test('ties reached by accumulation resolve like exact ones', async () => {
  const runs = Array.from(
    { length: DRIFT_LINES },
    (_, i) => `<w:r>${i ? '<w:br/>' : ''}<w:t xml:space="preserve">L${i}</w:t></w:r>`
  ).join('');
  const input = docx(
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${DRIFT_LINE_TWIPS}"` +
      ` w:lineRule="exact"/></w:pPr>${runs}</w:p>` +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      `<w:pgMar w:top="${DRIFT_TOP_TWIPS}" w:right="1440" w:bottom="1440" w:left="1440"` +
      ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
  );
  const result = await exportPdf(input, { useSystemFonts: false });
  const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
  try {
    const page = await pdf.getPage(1);
    const height = page.view[3]!;
    const content = await page.getTextContent();
    const tops = content.items
      .filter((item): item is Extract<typeof item, { str: string }> => 'str' in item)
      .filter((item) => /^L\d+$/.test(item.str))
      .map((item) => Number((height - item.transform[5]!).toFixed(4)))
      .sort((a, b) => a - b);
    expect(tops).toHaveLength(DRIFT_LINES);
    // Each line adds the SAME whole number of units for its in-line baseline, so subtracting
    // the tie rule applied to its own rule position must leave one shared offset.
    const offsets = tops.map((top, i) => {
      const twips = DRIFT_TOP_TWIPS + i * DRIFT_LINE_TWIPS;
      return Math.round(top / GRID) - Math.ceil((5 * twips - 12) / 24);
    });
    expect(new Set(offsets)).toHaveLength(1);
    // The fixture has to keep landing on ties, or it stops testing anything.
    const ties = offsets.filter(
      (_, i) => ((5 * (DRIFT_TOP_TWIPS + i * DRIFT_LINE_TWIPS)) / 24) % 1 === 0.5
    );
    expect(ties.length).toBeGreaterThanOrEqual(5);
  } finally {
    await pdf.destroy();
  }
});
