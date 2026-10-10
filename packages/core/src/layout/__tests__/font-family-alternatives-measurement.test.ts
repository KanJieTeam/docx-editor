import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolvedFontMeasurement } from '../../editor/resolved-font-measurement.js';
import { tryCreateCanvasMeasurer, type CanvasTextContext } from '../canvas-measurer.js';
import { createFontResourceSnapshot, FontResolutionError } from '../font-resource.js';
import { DEFAULT_RUN_STYLE, harfBuzzFontValidator, sha256FontBytes } from '../index.js';
import { createHarfBuzzLayoutOptions } from './fixtures/layout-shaping.js';

test('missing styled faces use the configured first family in Canvas fallback', () => {
  const regularBytes = new Uint8Array(
    readFileSync(new URL('./fixtures/fonts/DejaVuSans.ttf', import.meta.url))
  );
  const boldBytes = new Uint8Array(
    readFileSync(new URL('./fixtures/fonts/DejaVuSans-Bold.ttf', import.meta.url))
  );
  const fonts = createFontResourceSnapshot({
    resources: [
      {
        request: { family: 'First', weight: 400, style: 'normal' },
        id: 'first',
        bytes: regularBytes,
        hash: sha256FontBytes(regularBytes),
        faceIndex: 0,
      },
      {
        request: { family: 'Other', weight: 700, style: 'italic' },
        id: 'alternate',
        bytes: boldBytes,
        hash: sha256FontBytes(boldBytes),
        faceIndex: 0,
      },
    ],
    epoch: 1,
    maxFontBytes: 2_000_000,
    validateFont: harfBuzzFontValidator,
  });
  expect(fonts.resolve({ family: 'First', weight: 400, style: 'normal' })).not.toBeInstanceOf(
    FontResolutionError
  );
  expect(fonts.resolve({ family: 'Other', weight: 700, style: 'italic' })).not.toBeInstanceOf(
    FontResolutionError
  );
  const shaping = { ...createHarfBuzzLayoutOptions().shaping!, fonts };
  const reference = 'First;Other';
  const unavailable = shaping.fonts.resolve({
    family: reference,
    weight: 700,
    style: 'italic',
  });
  expect(unavailable).toBeInstanceOf(FontResolutionError);
  expect((unavailable as FontResolutionError).code).toBe('missing');

  const measuredFonts: string[] = [];
  const context = {
    font: '',
    measureText: () => {
      measuredFonts.push(context.font);
      return { width: 71 };
    },
  } as CanvasTextContext;
  const fontAlias = (family: string) => {
    if (family === 'First') return 'PrivateFirstFace';
    if (family === 'Other') return 'PrivateOtherFace';
    return undefined;
  };
  const canvas = tryCreateCanvasMeasurer({ context, fontAlias })!;
  const resolved = resolvedFontMeasurement(
    shaping,
    { sources: [] },
    { measurer: canvas, producer: 'canvas-measurer' },
    1,
    1
  );
  const measured = resolved.measurer.measure('PUBLIC', {
    ...DEFAULT_RUN_STYLE,
    fontFamily: reference,
    bold: true,
    italic: true,
  });
  expect(measured).toBe(71);
  expect(measuredFonts.length).toBeGreaterThan(0);
  const familyStack = measuredFonts.at(-1)!;
  expect(familyStack).toContain('italic');
  expect(familyStack).toMatch(/\bbold\b/);
  expect(familyStack.indexOf('"PrivateFirstFace"')).toBeGreaterThanOrEqual(0);
  expect(familyStack.indexOf('"PrivateOtherFace"')).toBeGreaterThan(
    familyStack.indexOf('"PrivateFirstFace"')
  );
  expect(familyStack.indexOf('"First"')).toBeGreaterThan(familyStack.indexOf('"PrivateFirstFace"'));
});
