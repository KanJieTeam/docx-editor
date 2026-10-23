import { expect, test } from 'bun:test';
import { readOoxmlPart } from '../../store/index.ts';
import { DEFAULT_RUN_STYLE, type ResolvedRunStyle } from '../run-style.ts';
import { createParagraphCjkOpticalFitter } from '../cjk-optical-fit.ts';
import { cjkParagraphBreaks } from '../cjk-paragraph-breaks.ts';
import { compressCjkPieces } from '../cjk-spacing.ts';
import { cjkTypographyFromSettings, resolveCjkTypography } from '../cjk-typography.ts';
import { withLayoutDerivedSpacing } from '../layout-derived-spacing.ts';
import type { FieldAwarePiece } from '../field-pieces.ts';
import type { PendingLine } from '../pending-line.ts';
import type { StyleSpanRecord, TextMeasurer } from '../semantic-records.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
interface Options {
  closingSize?: number;
  openingSize?: number;
  deficit?: number;
  noPunc?: boolean;
  sourceEA?: boolean;
  sourceGap?: boolean;
  slot?: 'ascii';
  unresolved?: boolean;
  closingReduced?: number;
  openingReduced?: number;
  baselineDerived?: boolean;
  narrowInk?: boolean;
  openingStyle?: Partial<ResolvedRunStyle>;
  pair?: '》（' | '」「' | '）（';
  demand?: boolean;
}
// These bearings are synthetic and scale with the requested font size.
// They are not the reviewer's SimSun/MS Mincho measurements.
function attempt(o: Options = {}) {
  const openingSize = o.openingSize ?? 12;
  const pair = o.pair ?? '》（';
  const base = {
    ...DEFAULT_RUN_STYLE,
    fontFamily: 'Synthetic Latin',
    fontFamilyEastAsia: 'Synthetic EA',
    fontSizePt: openingSize,
  };
  const m: TextMeasurer = {
    measure: (text, style) =>
      text.length *
      ((style.fontSizePt * style.horizontalScalePercent) / 100 + style.characterSpacingPt),
    lineMetrics: () => ({ height: 24, baseline: 19.2 }),
    hasResolvedFont: () => !o.unresolved,
    inkBounds: (text, style) =>
      text === pair[0]
        ? { left: style.fontSizePt * 0.1, right: style.fontSizePt * (o.narrowInk ? 0.2 : 0.6) }
        : text === pair[1]
          ? { left: style.fontSizePt * (o.narrowInk ? 0.9 : 0.4), right: style.fontSizePt * 0.95 }
          : { left: 0, right: style.fontSizePt },
  };
  const sourceSpans: StyleSpanRecord[] = [];
  let x = 0;
  for (const [index, text] of Array.from('甲' + pair).entries()) {
    const style =
      index === 1
        ? { ...base, fontSizePt: o.closingSize ?? openingSize }
        : index === 2
          ? { ...base, ...o.openingStyle }
          : base;
    const start = index + (o.sourceGap && index === 2 ? 1 : 0);
    const width = m.measure(text, style);
    sourceSpans.push({
      text,
      style,
      props: [],
      fontSlot: o.slot ?? 'eastAsia',
      range: { paragraphId: 'synthetic-pair', start, end: start + 1 },
      box: { x, y: 0, width, height: 24 },
    });
    x += width;
  }
  const end = sourceSpans.at(-1)!.range.end;
  const pieces: FieldAwarePiece[] = [
    ...sourceSpans.map((span) => ({
      text: span.text,
      start: span.range.start,
      end: span.range.end,
      style: span.style,
      props: span.props,
      fontSlot: span.fontSlot,
    })),
    { text: '乙丙', start: end, end: end + 2, style: base, props: [], fontSlot: 'eastAsia' },
  ];
  const spans = sourceSpans.map((span) => ({ ...span, box: { ...span.box } }));
  for (const [index, reduction] of [
    [1, o.closingReduced ?? 0],
    [2, o.openingReduced ?? 0],
  ] as const) {
    const span = spans[index]!;
    if (reduction) span.style = withLayoutDerivedSpacing(span.style, -reduction);
    span.box.width -= reduction;
    if (index === 2 && reduction) span.glyphOffsetPt = -reduction;
    if (o.baselineDerived) pieces[index] = { ...pieces[index]!, style: span.style };
  }
  x = 0;
  for (const span of spans) {
    span.box.x = x;
    x += span.box.width;
  }
  const line: PendingLine = {
    spans,
    drawings: [],
    start: 0,
    end,
    width: x,
    height: 24,
    baseline: 19.2,
    leading: 0,
    trailingSpacing: 0,
  };
  const before = JSON.stringify(line);
  const read = readOoxmlPart(
    `<w:settings xmlns:w="${W}"><w:characterSpacingControl w:val="compressPunctuation"/>${o.noPunc ? '<w:noPunctuationKerning/>' : ''}</w:settings>`,
    { name: '/word/settings.xml', contentType: 'application/xml' }
  );
  if (!read.ok) throw Error(read.reason);
  const typography = resolveCjkTypography(
    [{ localName: 'jc', namespaceUri: W, attributes: { val: 'both' } }],
    { ...cjkTypographyFromSettings(read.part.root), sourceEastAsianFont: o.sourceEA !== false }
  );
  const breaks = cjkParagraphBreaks(pieces, typography);
  const factory = createParagraphCjkOpticalFitter(
    pieces,
    pieces,
    breaks,
    m,
    typography,
    false,
    false,
    true
  );
  const deficit = o.deficit ?? openingSize * 0.75;
  const suffixWidth = m.measure('乙丙', base);
  const available = x + suffixWidth - deficit;
  const fitted = factory.fitCjkOptically(
    line,
    pieces.length - 1,
    '乙丙',
    { paragraphId: 'synthetic-pair', start: end, end: end + 2 },
    suffixWidth,
    x,
    available,
    o.demand !== false
  );
  return { line, before, fitted, available, deficit, openingSize };
}
function fitted(o: Options) {
  const a = attempt(o);
  expect(a.fitted).toBeDefined();
  expect(a.line.width + a.fitted!.width).toBeCloseTo(a.available, 8);
  const spans = [...a.line.spans, ...a.fitted!.candidate];
  expect(spans.map((span) => span.text).join('')).toBe('甲' + (o.pair ?? '》（') + '乙丙');
  expect(spans.map((span) => [span.range.start, span.range.end])).toEqual([
    [0, 1],
    [1, 2],
    [2, 3],
    [3, 4],
    [4, 5],
  ]);
  const close = spans.find((span) => span.text === (o.pair ?? '》（')[0])!;
  const open = spans.find((span) => span.text === (o.pair ?? '》（')[1])!;
  expect(close.box.width).toBeGreaterThanOrEqual(0);
  expect(open.box.width).toBeGreaterThanOrEqual(0);
  return a;
}
function rejected(o: Options) {
  const a = attempt(o);
  expect(a.fitted).toBeUndefined();
  expect(JSON.stringify(a.line)).toBe(a.before);
}
for (const openingSize of [10.5, 12]) {
  for (const closingSize of [openingSize, 20]) {
    for (const noPunc of [false, true]) {
      test(`synthetic ${closingSize}/${openingSize} pt pair spends both measured sides; noPunc=${noPunc}`, () => {
        const a = fitted({ openingSize, closingSize, noPunc, deficit: openingSize * 0.8 });
        expect(a.deficit).toBeCloseTo(openingSize * 0.8, 8);
      });
    }
  }
  test(`synthetic ${openingSize} pt opening cannot spend more than its two measured sides`, () =>
    rejected({ openingSize, closingSize: 20, deficit: openingSize * 0.8 + 0.01 }));
}
test('a smaller closing run still permits a safe mixed-size pair', () => {
  fitted({ closingSize: 10.5, openingSize: 12, deficit: 7.8 });
  fitted({ closingSize: 10.5, openingSize: 12, deficit: 3 });
});

// Each selected pair receives explicit synthetic closing/opening ink below.
// Closing right=.6em and opening left=.4em give two .4em budgets.
// These values do not represent measurements from either reviewed font.
for (const pair of ['」「', '）（'] as const)
  for (const openingSize of [10.5, 12]) {
    for (const noPunc of [false, true])
      test(`synthetic ${pair} at 20/${openingSize} pt measures both sides; noPunc=${noPunc}`, () => {
        const a = fitted({
          pair,
          closingSize: 20,
          openingSize,
          noPunc,
          deficit: openingSize * 0.8,
        });
        expect(a.deficit).toBeCloseTo(openingSize * 0.4 + openingSize * 0.4, 8);
      });
  }
test('an actual smaller shared ink gap caps the opening-size budget', () =>
  rejected({ closingSize: 10.5, openingSize: 12, deficit: 9.1 }));
for (const [label, options] of [
  ['missing EA defaults', { sourceEA: false }],
  ['different declared face', { openingStyle: { fontFamilyEastAsia: 'Other EA' } }],
  ['unresolved face', { unresolved: true }],
  ['ASCII slot', { slot: 'ascii' as const }],
  ['source gap', { sourceGap: true }],
  ['authored tracking', { openingStyle: { characterSpacingPt: 0.1 } }],
  ['authored scale', { openingStyle: { horizontalScalePercent: 90 } }],
  ['different bold face', { openingStyle: { bold: true } }],
  ['complex shaping', { openingStyle: { shaping: { direction: 'ltr' as const, script: 'Hani' } } }],
  ['no demand', { demand: false }],
] as const) {
  test(`${label} cannot borrow even the old one-sided low-deficit budget`, () =>
    rejected({ ...options, deficit: 3 }));
}
test('a derived closing budget is subtracted when baseline pieces are also derived', () => {
  fitted({ closingReduced: 4.8, baselineDerived: true, deficit: 4.8 });
  // An earlier legacy seam may already exceed this pair's closing-side budget.
  // Only the remaining opening side and the actual shared ink gap are available.
  for (const closingReduced of [6, 7]) {
    const sharedGap = 12 - closingReduced + 4.8 - 7.2;
    const a = fitted({ closingReduced, baselineDerived: true, deficit: sharedGap });
    expect(a.line.spans[1]!.box.width).toBeCloseTo(12 - closingReduced, 8);
    expect(a.line.spans[2]!.box.width).toBeCloseTo(12 - sharedGap, 8);
  }
});

for (const [closingReduced, openingReduced] of [
  [2, 3],
  [0, 4.8],
] as const) {
  test(`derived ${closingReduced}/${openingReduced} sides spend only their remaining measured budgets`, () => {
    const a = fitted({
      closingReduced,
      openingReduced,
      baselineDerived: true,
      deficit: 9.6 - closingReduced - openingReduced,
    });
    expect(a.deficit).toBeCloseTo(9.6 - closingReduced - openingReduced, 8);
    rejected({
      closingReduced,
      openingReduced,
      baselineDerived: true,
      deficit: 9.6 - closingReduced - openingReduced + 0.01,
    });
    expect(a.line.spans[1]!.box.width).toBeGreaterThanOrEqual(7.2 - 1e-8);
    expect(a.line.spans[2]!.box.width).toBeGreaterThanOrEqual(7.2 - 1e-8);
  });
}
test('saturated derived sides with a positive ink gap cannot fall through to one-sided compression', () =>
  rejected({
    closingReduced: 6,
    openingReduced: 6,
    baselineDerived: true,
    narrowInk: true,
    deficit: 3,
  }));

test('a paragraph cache prices boundary clusters independently of coalesced source prefixes', () => {
  const style = {
    ...DEFAULT_RUN_STYLE,
    fontFamily: 'Synthetic Latin',
    fontFamilyEastAsia: 'Synthetic EA',
    fontSizePt: 12,
  };
  const measuredInk: string[] = [];
  const m: TextMeasurer = {
    measure: (text, s) => text.length * (s.fontSizePt + s.characterSpacingPt),
    lineMetrics: () => ({ height: 24, baseline: 19.2 }),
    hasResolvedFont: () => true,
    inkBounds: (text, s) => {
      measuredInk.push(text);
      return text === '》'
        ? { left: s.fontSizePt * 0.1, right: s.fontSizePt * 0.6 }
        : text === '（'
          ? { left: s.fontSizePt * 0.4, right: s.fontSizePt * 0.95 }
          : { left: 0, right: text.length * s.fontSizePt };
    },
  };
  const pieces: FieldAwarePiece[] = [
    { text: '甲乙》', start: 0, end: 3, style, props: [], fontSlot: 'eastAsia' },
    { text: '（', start: 3, end: 4, style, props: [], fontSlot: 'eastAsia' },
    { text: '丙丁', start: 4, end: 6, style, props: [], fontSlot: 'eastAsia' },
  ];
  const typography = resolveCjkTypography(
    [{ localName: 'jc', namespaceUri: W, attributes: { val: 'both' } }],
    {
      sourceEastAsianFont: true,
      strict: false,
      before: {},
      after: {},
      compression: 'compressPunctuation',
    }
  );
  const factory = createParagraphCjkOpticalFitter(
    pieces,
    pieces,
    cjkParagraphBreaks(pieces, typography),
    m,
    typography,
    false,
    false,
    true
  );
  for (const [prefix, start] of [
    ['甲乙》', 0],
    ['乙》', 1],
  ] as const) {
    const spans: StyleSpanRecord[] = [prefix, '（'].map((text, index) => ({
      text,
      style,
      props: [],
      fontSlot: 'eastAsia',
      range: {
        paragraphId: 'cache-pair',
        start: index === 0 ? start : 3,
        end: index === 0 ? 3 : 4,
      },
      box: { x: index === 0 ? 0 : prefix.length * 12, y: 0, width: text.length * 12, height: 24 },
    }));
    const width = (prefix.length + 1) * 12;
    const line: PendingLine = {
      spans,
      drawings: [],
      start,
      end: 4,
      width,
      height: 24,
      baseline: 19.2,
      leading: 0,
      trailingSpacing: 0,
    };
    const fitted = factory.fitCjkOptically(
      line,
      2,
      '丙丁',
      { paragraphId: 'cache-pair', start: 4, end: 6 },
      24,
      width,
      width + 24 - 9.6,
      true
    );
    expect(fitted).toBeDefined();
    expect(line.width + fitted!.width).toBeCloseTo(width + 24 - 9.6, 8);
    expect(line.spans.map((span) => span.text).join('')).toBe(prefix + '（');
    expect(fitted!.candidate.map((span) => span.range)).toEqual([
      { paragraphId: 'cache-pair', start: 4, end: 5 },
      { paragraphId: 'cache-pair', start: 5, end: 6 },
    ]);
  }
  expect(measuredInk).not.toContain('甲乙》');
  expect(measuredInk).not.toContain('乙》');
});

test('both saturated sides reject a proven positive ink gap without changing the pending line', () => {
  const a = attempt({
    closingReduced: 6,
    openingReduced: 6,
    baselineDerived: true,
    narrowInk: true,
    deficit: 3,
  });
  const close = a.line.spans[1]!,
    open = a.line.spans[2]!;
  expect(close.box.width + (open.glyphOffsetPt ?? 0) + 12 * 0.9 - 12 * 0.2).toBeCloseTo(8.4, 8);
  expect(a.available - a.line.width).toBe(21);
  expect(a.fitted).toBeUndefined();
  expect(JSON.stringify(a.line)).toBe(a.before);
});

test('explicit alignment selects measured or natural pairs while an omitted low-level argument retains its legacy contract', () => {
  const style = {
    ...DEFAULT_RUN_STYLE,
    fontFamily: 'Synthetic EA',
    fontFamilyEastAsia: 'Synthetic EA',
    fontSizePt: 12,
    kerningMinPt: 1,
  };
  const pieces: FieldAwarePiece[] = [
    { text: '甲》（乙', start: 0, end: 4, style, props: [], fontSlot: 'eastAsia' },
  ];
  const m: TextMeasurer = {
    measure: (t, s) => t.length * (12 + s.characterSpacingPt),
    lineMetrics: () => ({ height: 24, baseline: 19.2 }),
    hasResolvedFont: () => true,
    inkBounds: (t) =>
      t === '》'
        ? { left: 1.2, right: 7.2 }
        : t === '（'
          ? { left: 4.8, right: 11.4 }
          : { left: 0, right: 12 },
  };
  const typography = resolveCjkTypography([], {
    strict: false,
    before: {},
    after: {},
    compression: 'compressPunctuation',
    sourceEastAsianFont: true,
  });
  const width = (result: readonly FieldAwarePiece[]) =>
    result.reduce((sum, p) => sum + m.measure(p.text, p.style), 0);
  expect(width(compressCjkPieces(pieces, typography, m))).toBe(42);
  expect(width(compressCjkPieces(pieces, typography, m, undefined, Infinity, false))).toBe(48);
  expect(width(compressCjkPieces(pieces, typography, m, undefined, Infinity, true))).toBeCloseTo(
    38.4,
    8
  );
  const noInk = { ...m, inkBounds: undefined };
  expect(width(compressCjkPieces(pieces, typography, noInk, undefined, Infinity, true))).toBe(48);
});
