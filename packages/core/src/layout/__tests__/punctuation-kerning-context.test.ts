import { afterAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { createLayoutShaping, disposeLayoutShaping } from '../layout-shaping.ts';
import { createShapedMeasurer } from '../shaped-measurer.ts';
import { createFixedMeasurer } from '../fixed-measurer.ts';
import { sha256FontBytes } from '../font-resource.ts';
import { resolveRunStyle } from '../run-style.ts';
import { shapeLayoutStyleRun } from '../layout-run-shape.ts';
import {
  punctuationKerningSegments,
  withPunctuationKerningPolicy,
} from '../punctuation-kerning.ts';

const bytes = new Uint8Array(
  readFileSync(new URL('./fixtures/fonts/DejaVuSans.ttf', import.meta.url))
);
const shaping = await createLayoutShaping({
  epoch: 1,
  maxFontBytes: 1_000_000,
  sources: [
    {
      request: { family: 'DejaVu Sans', weight: 400, style: 'normal' },
      id: 'dejavu',
      bytes,
      hash: sha256FontBytes(bytes),
      faceIndex: 0,
    },
  ],
  defaultFont: { family: 'DejaVu Sans', sizeHalfPoints: 28 },
});
const face = shaping.fonts.resolve({ family: 'DejaVu Sans', weight: 400, style: 'normal' });
if (face instanceof Error) throw face;
const measurer = createShapedMeasurer({
  shaper: shaping.shaper,
  environment: shaping.environment,
  resolveFont: () => face,
  fallback: createFixedMeasurer(),
});
const base = resolveRunStyle([
  { localName: 'rFonts', attributes: { ascii: 'DejaVu Sans', cs: 'DejaVu Sans' } },
  { localName: 'sz', attributes: { val: '28' } },
  { localName: 'kern', attributes: { val: '2' } },
]);

for (const { text, script, direction } of [
  { text: 'office.T.', script: 'Latn', direction: 'ltr' as const },
  { text: '😀office.\u0301T.', script: 'Latn', direction: 'ltr' as const },
  { text: 'سلام،', script: 'Arab', direction: 'rtl' as const },
  { text: '،سلام', script: 'Arab', direction: 'rtl' as const },
])
  test(`punctuation policy retains native glyph substitutions and UTF-16 clusters for ${text}`, () => {
    const original = {
      ...base,
      shaping: {
        script,
        direction,
        level: direction === 'rtl' ? 1 : 0,
        baseLevel: direction === 'rtl' ? 1 : 0,
      },
    };
    const marked = withPunctuationKerningPolicy(
      [{ text, start: 0, end: text.length, style: original, props: [] }],
      true
    )[0]!.style;
    const natural = shapeLayoutStyleRun(shaping.shaper, shaping.environment, face, original, text);
    const disabled = shapeLayoutStyleRun(shaping.shaper, shaping.environment, face, marked, text);
    expect(disabled.glyphs.map((glyph) => [glyph.id, glyph.cluster])).toEqual(
      natural.glyphs.map((glyph) => [glyph.id, glyph.cluster])
    );
    expect(disabled.clusters.map((cluster) => [cluster.textStart, cluster.textEnd])).toEqual(
      natural.clusters.map((cluster) => [cluster.textStart, cluster.textEnd])
    );
    const segmented = punctuationKerningSegments(text).reduce(
      (sum, segment) =>
        sum +
        measurer.measure(
          segment.text,
          segment.punctuation ? { ...original, kerningMinPt: 0 } : original
        ),
      0
    );
    expect(measurer.measure(text, marked)).toBeCloseTo(segmented, 2);
    expect(disabled.glyphs.every((glyph) => glyph.id !== 0)).toBe(true);
    if (text.includes('office'))
      expect(disabled.glyphs.some((glyph) => glyph.id === 5044)).toBe(true);
    if (text.startsWith('😀')) expect(disabled.clusters[0]!.textEnd).toBe(2);
    expect(
      withPunctuationKerningPolicy(
        [{ text, start: 0, end: text.length, style: original, props: [] }],
        false
      )[0]!.style
    ).toBe(original);
  });

afterAll(() => disposeLayoutShaping(shaping));
