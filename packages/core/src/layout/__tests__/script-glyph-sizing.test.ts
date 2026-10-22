import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { readOoxmlPart, serializeOoxmlPart } from '../../store/index.ts';
import { paintSemanticLayout } from '../../output/semantic-paint.ts';
import { bindExportLaidOutText } from '../../export/export-laid-out-text.ts';
import { acquireSharedExportShaping } from '../../export/shared-export-shaping.ts';
import {
  DEFAULT_RUN_STYLE,
  glyphSizeFactorOf,
  runStylesEqual,
  baselineShiftPtOf,
} from '../run-style.ts';
import { createShapedMeasurer } from '../shaped-measurer.ts';
import { createFixedMeasurer } from '../fixed-measurer.ts';
import { projectScriptGlyphSizing } from '../paragraph-script-glyph-sizing.ts';
import { styleForFontSlot } from '../script-itemization.ts';
import { growScriptLineMetrics } from '../paragraph-mark-metrics.ts';
import { scriptGlyphStyleForEditing, ScriptGlyphSizingError } from '../layout-glyph-sizing.ts';
import {
  createFontResourceSnapshot,
  boundedStructuralFontValidator,
  sha256FontBytes,
  FontResolutionError,
  type ResolvedFont,
} from '../font-resource.ts';
import {
  markFacePreservingShaper,
  fixedPoint,
  type ShapeInput,
  type ShapedRun,
} from '../shaped-run.ts';
import { HARFBUZZ_SHAPING_LIBRARY } from '../harfbuzz-shaper.ts';
import { prepareLayoutFontConfiguration, withScriptGlyphSizing } from '../layout-shaping.ts';
import { layoutSemanticDocument } from '../semantic-layout.ts';
import { createParagraphLayoutCache } from '../layout-cache.ts';
import { sfntScriptSizeEm } from '../sfnt-script-size.ts';
import { linesOf, type StyleSpanRecord } from '../semantic-records.ts';
import type { FieldAwarePiece } from '../field-pieces.ts';
import type { LayoutShapingOptions } from '../shaped-measurer.ts';
import {
  syntheticScriptFont,
  syntheticScriptCollection,
} from './fixtures/synthetic-script-font.ts';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const half = () => syntheticScriptFont({ em: 256, subscriptY: 128, superscriptY: 160 });
function admitted(bytes = half(), faceIndex = 0, family = 'Script Face') {
  const request = { family, weight: 400, style: 'normal' as const };
  const fonts = createFontResourceSnapshot({
    epoch: 1,
    maxFontBytes: 2048,
    resources: [{ request, id: family, bytes, hash: sha256FontBytes(bytes), faceIndex }],
    validateFont: boundedStructuralFontValidator,
  });
  const font = fonts.resolve(request);
  if (font instanceof FontResolutionError) throw font;
  return { font, fonts };
}
const primary = admitted();
const style = Object.freeze({
  ...DEFAULT_RUN_STYLE,
  fontFamily: 'Script Face',
  fontSizePt: 14,
  verticalAlign: 'subscript' as const,
});
function run(input: ShapeInput): ShapedRun {
  const advance = input.fontSizeHalfPoints * 500;
  return {
    text: input.text,
    direction: 'ltr',
    bidiLevel: 0,
    glyphs: Array.from(input.text, (_, index) => ({
      id: index + 1,
      cluster: index,
      originX: fixedPoint(index * advance),
      originY: fixedPoint(0),
      advanceX: fixedPoint(advance),
      advanceY: fixedPoint(0),
      offsetX: fixedPoint(0),
      offsetY: fixedPoint(0),
      outline: { path: 'M0 0L256 0L256 256L0 256Z', unitsPerEm: 256 },
    })),
    clusters: Array.from(input.text, (_, index) => ({
      textStart: index,
      textEnd: index + 1,
      glyphStart: index,
      glyphEnd: index + 1,
      advance: fixedPoint(advance),
      caretEdges: [fixedPoint(0), fixedPoint(advance)],
      fontSpan: 0,
    })),
    fontSpans: input.text
      ? [
          {
            glyphStart: 0,
            glyphEnd: input.text.length,
            font: input.environment.font,
            fallbackIndex: null,
          },
        ]
      : [],
    metrics: {
      ascent: fixedPoint(advance * 0.8),
      descent: fixedPoint(advance * 0.2),
      lineGap: fixedPoint(0),
    },
  };
}
function measurer(
  mode: 'legacy' | 'font-metrics' = 'font-metrics',
  shape = run,
  resolve: (style: typeof DEFAULT_RUN_STYLE) => ResolvedFont | null = () => primary.font
) {
  return createShapedMeasurer({
    shaper: markFacePreservingShaper({ shape }),
    resolveFont: resolve,
    fallback: createFixedMeasurer(14, 14),
    shapingLibrary: HARFBUZZ_SHAPING_LIBRARY,
    unicodeDataVersion: '16.0.0',
    scriptGlyphSizing: mode,
  });
}
function piece(text = '12', overrides: Partial<FieldAwarePiece> = {}): FieldAwarePiece {
  return { text, start: 0, end: text.length, props: [], style, ...overrides };
}
function projected(text = '12', measure = measurer()) {
  return projectScriptGlyphSizing([piece(text)], 'p', measure)[0]!.style;
}

test('default mode keeps source identity, legacy factor, and baseline offset', () => {
  const ordinary = [piece()];
  const legacy = measurer('legacy');
  expect(legacy.scriptGlyphSizing).toBeUndefined();
  expect(projectScriptGlyphSizing(ordinary, 'p', legacy)).toBe(ordinary);
  expect(glyphSizeFactorOf(style)).toBe(0.65);
  expect(legacy.measure('12', style)).toBe(18.2);
  expect(baselineShiftPtOf(projected())).toBe(baselineShiftPtOf(style));
});

test('one qualified face controls width, prefix carets, ink, and complete outline origins', () => {
  const measure = measurer();
  const derived = projected('12', measure);
  expect(derived.fontSizePt).toBe(14);
  expect(style.layoutGlyphSizing).toBeUndefined();
  expect(glyphSizeFactorOf(derived)).toBe(0.5);
  expect(measure.measure('12', derived)).toBe(14);
  expect(measure.measure('1', derived)).toBe(7);
  expect(measure.caretAdvances!('12', derived)).toEqual([0, 7, 14]);
  expect(measure.inkBounds!('1', derived)).toEqual({ left: 0, right: 7 });
  expect(measure.textOutline!('12', derived)!.ink.right).toBe(14);
  expect(Object.isFrozen(derived)).toBe(true);
  expect(Object.isFrozen(derived.layoutGlyphSizing)).toBe(true);
});

test('collection face selection, slot choice, and independent proof scope stay distinct', () => {
  const bytes = syntheticScriptCollection([
    half(),
    syntheticScriptFont({ em: 256, subscriptY: 192, superscriptY: 128 }),
  ]);
  const other = admitted(bytes, 1, 'Other Face').font;
  const measure = measurer('font-metrics', run, () => other);
  expect(glyphSizeFactorOf(projected('12', measure))).toBe(0.75);
  const mixed = piece('土', {
    style: { ...style, fontFamilyEastAsia: 'East Face' },
    fontSlot: 'eastAsia',
  });
  const derived = projectScriptGlyphSizing([mixed], 'p', measurer())[0]!;
  expect(derived.style.fontFamily).toBe('Script Face');
  expect(glyphSizeFactorOf(styleForFontSlot(derived.style, 'eastAsia'))).toBe(0.5);
  const first = projected();
  const split = { ...first };
  expect(runStylesEqual(first, split)).toBe(true);
  const second = projectScriptGlyphSizing([piece('12', { start: 2, end: 4 })], 'p', measurer())[0]!
    .style;
  expect(runStylesEqual(first, second)).toBe(false);
});

test('generated text, invalid extents, and unavailable metrics keep the entire source piece legacy', () => {
  const measure = measurer();
  for (const entry of [
    piece('12', { projected: true }),
    piece('12', { measureText: '123' }),
    piece('12', { end: 3 }),
    piece('12', { style: { ...style, caps: true } }),
    piece('12', { style: { ...style, smallCaps: true } }),
  ]) {
    expect(projectScriptGlyphSizing([entry], 'p', measure)[0]).toBe(entry);
  }
  const missing = admitted(syntheticScriptFont({ em: 256, subscriptY: 0, superscriptY: 128 })).font;
  expect(
    projected(
      '12',
      measurer('font-metrics', run, () => missing)
    )
  ).toBe(style);
});

test('a rejected whole piece never qualifies its supported prefix during line measurement', () => {
  let qualifications = 0;
  const measure = measurer('font-metrics', (input) => {
    if (input.text === '12') qualifications++;
    const shaped = run(input);
    return input.text === '12'
      ? { ...shaped, glyphs: shaped.glyphs.map((g, i) => (i ? { ...g, id: 0 } : g)) }
      : shaped;
  });
  const derived = projected('12', measure);
  expect(derived).toBe(style);
  expect(qualifications).toBe(1);
  expect(measure.measure('1', derived)).toBe(9.1);
  expect(qualifications).toBe(1);
});

test('unexpected prefix fallback refuses geometry instead of changing a published factor', () => {
  const measure = measurer('font-metrics', (input) => {
    const shaped = run(input);
    return input.text === '1'
      ? { ...shaped, glyphs: shaped.glyphs.map((g) => ({ ...g, id: 0 })) }
      : shaped;
  });
  const derived = projected('12', measure);
  expect(() => measure.measure('1', derived)).toThrow(ScriptGlyphSizingError);
  expect(() => measure.caretAdvances!('1', derived)).toThrow(ScriptGlyphSizingError);
  expect(() => measure.inkBounds!('1', derived)).toThrow(ScriptGlyphSizingError);
  expect(() => measure.textOutline!('1', derived)).toThrow(ScriptGlyphSizingError);
  expect(
    measurer('legacy', (input) => ({
      ...run(input),
      glyphs: run(input).glyphs.map((g) => ({ ...g, id: 0 })),
    })).measure('1', style)
  ).toBe(createFixedMeasurer(14, 14).measure('1', style));
});

test('stale published carrier copies refuse; editing and nominal line metrics strip the carrier', () => {
  const measure = measurer();
  const derived = projected('12', measure);
  for (const changed of [
    { caps: true },
    { smallCaps: true },
    { fontSizePt: 16 },
    { italic: true },
    { verticalAlign: 'superscript' as const },
    { verticalAlign: 'baseline' as const },
    { fontFamily: 'Other' },
  ]) {
    expect(() => glyphSizeFactorOf({ ...derived, ...changed })).toThrow(ScriptGlyphSizingError);
  }
  expect(scriptGlyphStyleForEditing(derived)).toEqual(style);
  const line = { height: 7, baseline: 5.6, spans: [{ text: '12', style: derived }] };
  growScriptLineMetrics(line, 'baseline', measure);
  expect(line.height).toBe(14);
  expect(line.baseline).toBeCloseTo(11.2);
});

test('PDF binding uses the same selected face and refuses a changed font resource', () => {
  const measure = measurer();
  const derived = projected('12', measure);
  const span = {
    text: '12',
    style: derived,
    props: [],
    range: { paragraphId: 'p', start: 0, end: 2 },
    box: { x: 0, y: 0, width: 14, height: 14 },
  } satisfies StyleSpanRecord;
  const shaping = {
    fonts: primary.fonts,
    shaper: markFacePreservingShaper({ shape: run }),
    defaultFont: { family: 'Script Face', sizeHalfPoints: 28 },
    environment: {
      shapingLibrary: HARFBUZZ_SHAPING_LIBRARY,
      unicodeDataVersion: '16.0.0',
      script: 'Latn',
      fixedPointScale: 1000,
      normalization: 'none',
      features: {},
      roundingMode: 'halfToEven',
      language: 'en',
      variationAxes: {},
    },
    ligatureCaretPolicy: 'cluster-edges-only',
    operation: {
      resourceEpoch: 1,
      configEpoch: 1,
      extensionFingerprint: 'test',
      shapingHash: 'test',
      producerVersion: 1,
    },
  } satisfies LayoutShapingOptions;
  expect(bindExportLaidOutText(shaping)(span)).not.toBeNull();
  const other = admitted(syntheticScriptFont({ em: 256, subscriptY: 192, superscriptY: 128 }));
  expect(bindExportLaidOutText({ ...shaping, fonts: other.fonts })(span)).toBeNull();
  const metric = withScriptGlyphSizing(shaping, 'font-metrics');
  expect(withScriptGlyphSizing(metric, 'font-metrics')).toBe(metric);
  expect(withScriptGlyphSizing(metric, 'legacy').operation.shapingHash).toBe(
    shaping.operation.shapingHash
  );
});

test('ordinary body and cell text publish the same factor and paint while source OOXML stays nominal', () => {
  const body =
    '<w:p><w:r><w:rPr><w:rFonts w:ascii="Script Face"/><w:sz w:val="28"/><w:vertAlign w:val="subscript"/></w:rPr><w:t>12</w:t></w:r></w:p>';
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${body}</w:tc></w:tr></w:tbl></w:body></w:document>`;
  const read = readOoxmlPart(xml, { name: '/word/document.xml', contentType: 'app/xml' });
  if (!read.ok) throw new Error(read.reason);
  const before = serializeOoxmlPart(read.part);
  const metric = layoutSemanticDocument(read.part, 1, { measurer: measurer() });
  const spans = linesOf(metric)
    .flatMap((line) => line.spans)
    .filter((span) => span.text === '12');
  expect(spans).toHaveLength(2);
  for (const span of spans) {
    expect(span.style.fontSizePt).toBe(14);
    expect(span.box.width).toBe(14);
  }
  const container = document.createElement('div');
  paintSemanticLayout(container, metric, { scale: 1 });
  const painted = [...container.querySelectorAll<HTMLElement>('.layout-run-text')].filter(
    (span) => span.textContent === '12'
  );
  expect(painted).toHaveLength(2);
  for (const span of painted) expect(span.style.fontSize).toBe('7px');
  expect(serializeOoxmlPart(read.part)).toBe(before);
  container.remove();
});

test('real HarfBuzz prefixes and shared export mode caches retain one admitted face policy', async () => {
  const bytes = new Uint8Array(
    readFileSync(new URL('./fixtures/fonts/DejaVuSans.ttf', import.meta.url))
  );
  const prepared = prepareLayoutFontConfiguration({
    epoch: 1,
    maxFontBytes: 2_000_000,
    sources: [
      {
        request: { family: 'DejaVu Sans', weight: 400, style: 'normal' },
        id: 'real',
        bytes,
        hash: sha256FontBytes(bytes),
        faceIndex: 0,
      },
    ],
    defaultFont: { family: 'DejaVu Sans', sizeHalfPoints: 28 },
  });
  const legacy = await acquireSharedExportShaping(prepared);
  const metric = await acquireSharedExportShaping(prepared, undefined, {
    scriptGlyphSizing: 'font-metrics',
  });
  expect(await acquireSharedExportShaping(prepared)).toBe(legacy);
  expect(
    await acquireSharedExportShaping(prepared, undefined, { scriptGlyphSizing: 'font-metrics' })
  ).toBe(metric);
  expect(metric.producer).not.toBe(legacy.producer);
  const measure = metric.createMeasurer();
  const source = { ...style, fontFamily: 'DejaVu Sans' };
  const derived = projectScriptGlyphSizing([piece('123', { style: source })], 'real', measure)[0]!
    .style;
  expect(derived.layoutGlyphSizing).toBeDefined();
  const base = { ...source, verticalAlign: 'baseline' as const };
  expect(measure.measure('123', derived)).toBeCloseTo(
    measure.measure('123', base) * glyphSizeFactorOf(derived),
    8
  );
  const edges = measure.caretAdvances!('123', derived)!;
  expect(edges.at(-1)).toBeCloseTo(measure.measure('123', derived), 8);
  expect(measure.measure('1', derived)).toBeCloseTo(
    measure.measure('1', base) * glyphSizeFactorOf(derived),
    8
  );
});

test('warm body and table caches match fresh geometry across legacy, metric, and legacy views', async () => {
  const bytes = new Uint8Array(
    readFileSync(new URL('./fixtures/fonts/DejaVuSans.ttf', import.meta.url))
  );
  const prepared = prepareLayoutFontConfiguration({
    epoch: 2,
    maxFontBytes: 2_000_000,
    sources: [
      {
        request: { family: 'DejaVu Sans', weight: 400, style: 'normal' },
        id: 'cache',
        bytes,
        hash: sha256FontBytes(bytes),
        faceIndex: 0,
      },
    ],
    defaultFont: { family: 'DejaVu Sans', sizeHalfPoints: 28 },
  });
  const legacy = await acquireSharedExportShaping(prepared);
  const metric = await acquireSharedExportShaping(prepared, undefined, {
    scriptGlyphSizing: 'font-metrics',
  });
  const p =
    '<w:p><w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans"/><w:sz w:val="28"/><w:vertAlign w:val="subscript"/></w:rPr><w:t>123 456 789</w:t></w:r></w:p>';
  const read = readOoxmlPart(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${p}<w:tbl><w:tblPr><w:tblW w:w="1600" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1600"/></w:tblGrid><w:tr><w:tc>${p}</w:tc></w:tr></w:tbl></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  const cache = createParagraphLayoutCache<never>();
  const geometry = (layout: ReturnType<typeof layoutSemanticDocument>) =>
    linesOf(layout).map((line) => ({
      box: line.box,
      spans: line.spans.map((span) => ({
        text: span.text,
        box: span.box,
        factor: glyphSizeFactorOf(styleForFontSlot(span.style, span.fontSlot)),
        nominal: span.style.fontSizePt,
      })),
    }));
  const views = [legacy, metric, legacy].map((view) => ({ view, measurer: view.createMeasurer() }));
  for (const { view, measurer } of views) {
    const warm = layoutSemanticDocument(read.part, 1, {
      measurer,
      producer: view.producer,
      cache: cache as never,
    });
    const fresh = layoutSemanticDocument(read.part, 1, { measurer, producer: view.producer });
    expect(geometry(warm)).toEqual(geometry(fresh));
  }
  expect(cache.stats.hits).toBeGreaterThan(0);
});

test('script metadata reads independent Y sizes and the selected collection face', () => {
  expect(sfntScriptSizeEm(half(), 0)).toEqual({ subscript: 0.5, superscript: 0.625 });
  expect(
    sfntScriptSizeEm(
      syntheticScriptCollection([
        half(),
        syntheticScriptFont({ em: 256, subscriptY: 192, superscriptY: 128 }),
      ]),
      1
    )
  ).toEqual({ subscript: 0.75, superscript: 0.5 });
});

test('script metadata refuses invalid indexes, truncated directories, and excessive table counts', () => {
  for (const index of [-1, 0.5, 1, Number.NaN]) expect(sfntScriptSizeEm(half(), index)).toBeNull();
  expect(sfntScriptSizeEm(half().slice(0, 24), 0)).toBeNull();
  const bytes = half();
  new DataView(bytes.buffer).setUint16(4, 4097);
  expect(sfntScriptSizeEm(bytes, 0)).toBeNull();
});

test('script metadata refuses absent, zero, negative, and excessive recommendations', () => {
  expect(
    sfntScriptSizeEm(
      syntheticScriptFont({ em: 256, subscriptY: 128, superscriptY: 128, omitOs2: true }),
      0
    )
  ).toBeNull();
  expect(
    sfntScriptSizeEm(syntheticScriptFont({ em: 256, subscriptY: 0, superscriptY: 300 }), 0)
  ).toEqual({ subscript: null, superscript: null });
  expect(
    sfntScriptSizeEm(syntheticScriptFont({ em: 256, subscriptY: -1, superscriptY: 128 }), 0)
  ).toEqual({ subscript: null, superscript: 0.5 });
});

test('script metadata refuses duplicate table records without changing strikeout parsing', () => {
  const bytes = half();
  const view = new DataView(bytes.buffer);
  view.setUint32(28, 0x68656164);
  expect(sfntScriptSizeEm(bytes, 0)).toBeNull();
});

test('different admitted Latin and East Asian faces retain their factors through body, cell, paint, and export', () => {
  const latin = syntheticScriptFont({ em: 256, subscriptY: 192, superscriptY: 192 });
  const east = half();
  const fonts = createFontResourceSnapshot({
    epoch: 1,
    maxFontBytes: 2048,
    resources: [
      {
        request: { family: 'Latin Face', weight: 400, style: 'normal' },
        id: 'latin',
        bytes: latin,
        hash: sha256FontBytes(latin),
        faceIndex: 0,
      },
      {
        request: { family: 'East Face', weight: 400, style: 'normal' },
        id: 'east',
        bytes: east,
        hash: sha256FontBytes(east),
        faceIndex: 0,
      },
    ],
    validateFont: boundedStructuralFontValidator,
  });
  const resolve = (style: typeof DEFAULT_RUN_STYLE) => {
    const font = fonts.resolve({
      family: style.fontFamily ?? 'Latin Face',
      weight: 400,
      style: 'normal',
    });
    if (font instanceof FontResolutionError) return null;
    return font;
  };
  const measure = measurer('font-metrics', run, resolve);
  const p =
    '<w:p><w:r><w:rPr><w:rFonts w:ascii="Latin Face" w:hAnsi="Latin Face" w:eastAsia="East Face"/><w:sz w:val="28"/><w:vertAlign w:val="subscript"/></w:rPr><w:t>Q3土</w:t></w:r></w:p>';
  const read = readOoxmlPart(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${p}<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${p}</w:tc></w:tr></w:tbl></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  const layout = layoutSemanticDocument(read.part, 1, { measurer: measure });
  const spans = linesOf(layout).flatMap((line) => line.spans);
  const source = spans.filter((span) => span.text === 'Q3' || span.text === '土');
  expect(source).toHaveLength(4);
  const admittedHashes = new Set(source.map((span) => span.style.layoutGlyphSizing!.fontHash));
  expect(admittedHashes.size).toBe(2);
  for (const span of source) {
    const faceStyle = styleForFontSlot(span.style, span.fontSlot);
    expect(glyphSizeFactorOf(faceStyle)).toBe(span.text === '土' ? 0.5 : 0.75);
    expect(measure.caretAdvances!(span.text, faceStyle)!.at(-1)).toBeCloseTo(span.box.width);
    expect(span.style.fontSizePt).toBe(14);
    expect(span.style.fontFamily).toBe('Latin Face');
    expect(span.style.fontFamilyEastAsia).toBe('East Face');
  }
  const shaping = {
    fonts,
    shaper: markFacePreservingShaper({ shape: run }),
    defaultFont: { family: 'Latin Face', sizeHalfPoints: 28 },
    environment: {
      shapingLibrary: HARFBUZZ_SHAPING_LIBRARY,
      unicodeDataVersion: '16.0.0',
      script: 'Latn',
      fixedPointScale: 1000,
      normalization: 'none',
      features: {},
      roundingMode: 'halfToEven',
      language: 'en',
      variationAxes: {},
    },
    ligatureCaretPolicy: 'cluster-edges-only',
    operation: {
      resourceEpoch: 1,
      configEpoch: 1,
      extensionFingerprint: 'mixed',
      shapingHash: 'mixed',
      producerVersion: 1,
    },
  } satisfies LayoutShapingOptions;
  for (const span of source) expect(bindExportLaidOutText(shaping)(span)).not.toBeNull();
  const container = document.createElement('div');
  paintSemanticLayout(container, layout, { scale: 1 });
  const painted = [...container.querySelectorAll<HTMLElement>('.layout-run-text')].filter(
    (span) => span.textContent === 'Q3' || span.textContent === '土'
  );
  expect(painted).toHaveLength(4);
  for (const span of painted)
    expect(span.style.fontSize).toBe(span.textContent === '土' ? '7px' : '10.5px');
  container.remove();
});

test('long ordinary script pieces split in body and cell with stable source ranges, carets, and paint', () => {
  const text = '1234567890'.repeat(7);
  const p = `<w:p><w:r><w:rPr><w:rFonts w:ascii="Script Face"/><w:sz w:val="28"/><w:vertAlign w:val="subscript"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const read = readOoxmlPart(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${p}<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/><w:tblW w:w="1600" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1600"/></w:tblGrid><w:tr><w:tc>${p}</w:tc></w:tr></w:tbl></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  const measure = measurer();
  const layout = layoutSemanticDocument(read.part, 1, {
    measurer: measure,
    geometry: { width: 100, height: 1000, margin: { top: 10, right: 10, bottom: 10, left: 10 } },
  });
  const groups = new Map<string, StyleSpanRecord[]>();
  for (const line of linesOf(layout))
    for (const span of line.spans) {
      const group = groups.get(span.range.paragraphId) ?? [];
      group.push(span);
      groups.set(span.range.paragraphId, group);
    }
  expect(groups.size).toBe(2);
  for (const spans of groups.values()) {
    expect(spans.length).toBeGreaterThan(1);
    expect(spans.map((span) => span.text).join('')).toBe(text);
    let end = 0;
    const scopes = new Set<string>();
    for (const span of spans) {
      expect(span.range.start).toBe(end);
      end = span.range.end;
      expect(span.range.end - span.range.start).toBe(span.text.length);
      expect(glyphSizeFactorOf(styleForFontSlot(span.style, span.fontSlot))).toBe(0.5);
      expect(
        measure.caretAdvances!(span.text, styleForFontSlot(span.style, span.fontSlot))!.at(-1)
      ).toBeCloseTo(span.box.width);
      scopes.add(span.style.layoutGlyphSizing!.scope);
    }
    expect(end).toBe(text.length);
    expect(scopes.size).toBe(1);
  }
  const container = document.createElement('div');
  paintSemanticLayout(container, layout, { scale: 1 });
  const painted = [...container.querySelectorAll<HTMLElement>('.layout-run-text')];
  expect(painted.length).toBeGreaterThan(2);
  for (const span of painted) expect(span.style.fontSize).toBe('7px');
  expect(painted.map((span) => span.textContent).join('')).toBe(text + text);
  container.remove();
});

test('real joined RTL text keeps its qualified factor across normal line splitting', async () => {
  const bytes = new Uint8Array(
    readFileSync(new URL('./fixtures/fonts/DejaVuSans.ttf', import.meta.url))
  );
  const prepared = prepareLayoutFontConfiguration({
    epoch: 3,
    maxFontBytes: 2_000_000,
    sources: [
      {
        request: { family: 'DejaVu Sans', weight: 400, style: 'normal' },
        id: 'rtl',
        bytes,
        hash: sha256FontBytes(bytes),
        faceIndex: 0,
      },
    ],
    defaultFont: { family: 'DejaVu Sans', sizeHalfPoints: 28 },
  });
  const view = await acquireSharedExportShaping(prepared, undefined, {
    scriptGlyphSizing: 'font-metrics',
  });
  const text = 'السَّلَامُعَلَيْكُمْ'.repeat(10);
  const read = readOoxmlPart(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans" w:cs="DejaVu Sans"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:rtl/><w:vertAlign w:val="subscript"/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  const measure = view.createMeasurer();
  const layout = layoutSemanticDocument(read.part, 1, {
    measurer: measure,
    producer: view.producer,
    geometry: { width: 140, height: 1000, margin: { top: 10, right: 10, bottom: 10, left: 10 } },
  });
  const lines = linesOf(layout);
  expect(lines.length).toBeGreaterThan(1);
  const spans = lines.flatMap((line) => line.spans).sort((a, b) => a.range.start - b.range.start);
  expect(spans.map((span) => span.text).join('')).toBe(text);
  const factors = new Set<number>();
  const scopes = new Set<string>();
  for (const span of spans) {
    const face = styleForFontSlot(span.style, span.fontSlot);
    expect(span.style.layoutGlyphSizing).toBeDefined();
    factors.add(glyphSizeFactorOf(face));
    scopes.add(span.style.layoutGlyphSizing!.scope);
    expect(measure.caretAdvances!(span.text, face)!.every(Number.isFinite)).toBe(true);
    expect(view.shapeLaidOutText(span)).not.toBeNull();
  }
  expect(factors.size).toBe(1);
  expect(scopes.size).toBe(1);
});

test('qualified text without supported outlines never borrows another face outline', () => {
  let fallbacks = 0;
  const substitute = { path: 'M0 0L1 1Z', ink: { left: 0, top: 0, right: 1, bottom: 1 } };
  const create = (mode: 'legacy' | 'font-metrics') =>
    createShapedMeasurer({
      shaper: markFacePreservingShaper({
        shape(input) {
          const shaped = run(input);
          return {
            ...shaped,
            glyphs: shaped.glyphs.map((glyph) => ({
              ...glyph,
              outline: { ...glyph.outline, path: 'unsupported' },
            })),
          };
        },
      }),
      resolveFont: () => primary.font,
      fallback: {
        ...createFixedMeasurer(),
        textOutline() {
          fallbacks++;
          return substitute;
        },
      },
      shapingLibrary: HARFBUZZ_SHAPING_LIBRARY,
      unicodeDataVersion: '16.0.0',
      scriptGlyphSizing: mode,
    });
  const metric = create('font-metrics');
  const derived = projected('12', metric);
  expect(metric.textOutline!('12', derived)).toBeUndefined();
  expect(fallbacks).toBe(0);
  expect(create('legacy').textOutline!('12', style)).toBe(substitute);
  expect(fallbacks).toBe(1);
});

test('an old qualified span cannot reuse metrics from a different admitted resource', () => {
  const derived = projected('12');
  const replacement = admitted(
    syntheticScriptFont({ em: 256, subscriptY: 192, superscriptY: 128 })
  ).font;
  const next = measurer('font-metrics', run, () => replacement);
  expect(() => next.lineMetrics(derived, '12')).toThrow(ScriptGlyphSizingError);
  expect(() => next.strikeoutMetrics!(derived)).toThrow(ScriptGlyphSizingError);
  expect(() => next.measure('12', derived)).toThrow(ScriptGlyphSizingError);
  expect(() => next.caretAdvances!('12', derived)).toThrow(ScriptGlyphSizingError);
  expect(() => next.inkBounds!('1', derived)).toThrow(ScriptGlyphSizingError);
  expect(() => next.textOutline!('12', derived)).toThrow(ScriptGlyphSizingError);
  expect(next.measure('12', style)).toBe(18.2);
});
