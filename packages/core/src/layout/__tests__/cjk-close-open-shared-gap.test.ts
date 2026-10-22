import { test, expect } from 'bun:test';
import { DEFAULT_RUN_STYLE } from '../run-style.ts';
import { createCjkOpticalFitter } from '../cjk-optical-fit.ts';
import { cjkParagraphBreaks } from '../cjk-paragraph-breaks.ts';
import { DEFAULT_CJK_TYPOGRAPHY } from '../cjk-typography.ts';
import {
  withLayoutDerivedSpacing,
  layoutDerivedSpacingStyleForEditing,
} from '../layout-derived-spacing.ts';
function attempt(o: any = {}) {
  const prefix = o.prefix ?? '《甲。乙》（丙',
    suffix = o.suffix ?? '戊己';
  const base = {
    ...DEFAULT_RUN_STYLE,
    fontFamily: 'Invented Latin',
    fontFamilyEastAsia: 'Invented EA',
    fontSizePt: 14,
  };
  const m: any = {
    measure: (t: string, s: any) =>
      (t.length * s.fontSizePt * s.horizontalScalePercent) / 100 + t.length * s.characterSpacingPt,
    lineMetrics: () => ({ height: 18.2, baseline: 14 }),
    hasResolvedFont: (s: any) => (o.unresolved ? s.fontFamily !== 'Invented EA' : true),
    inkBounds: (t: string) =>
      t === '》'
        ? { left: 1, right: o.saturatedInk ? 3 : o.tightGap ? 13.5 : 9.4 }
        : /^[《（]$/.test(t)
          ? { left: o.saturatedInk ? 12 : 7, right: 13 }
          : t === '。'
            ? { left: 0, right: 4 }
            : { left: 0, right: 12 },
  };
  let x = 0;
  const spans = Array.from(prefix).map((text, index) => {
    const style = index === prefix.indexOf('（') ? { ...base, ...o.openStyle } : base;
    const slot = o.asciiSlot ? 'ascii' : 'eastAsia';
    const start = index + (o.sourceGap && index >= prefix.indexOf('（') ? 1 : 0);
    const v: any = {
      text,
      style,
      props: [],
      fontSlot: slot,
      range: { paragraphId: 'generic-pair', start, end: start + 1 },
      box: { x, y: 0, width: m.measure(text, style), height: 18.2 },
    };
    x += v.box.width;
    return v;
  });
  // The selected prefix determines the closing and opening cell indices.
  const end = spans.at(-1)!.range.end;
  const pieces: any[] = [
    ...spans.map((s) => ({
      text: s.text,
      start: s.range.start,
      end: s.range.end,
      style: s.style,
      props: s.props,
      fontSlot: s.fontSlot,
    })),
    {
      text: suffix,
      start: end,
      end: end + suffix.length,
      style: base,
      props: [],
      fontSlot: 'eastAsia',
    },
  ];
  // Baseline pieces remain authored; only the actual PendingLine has derived spacing.
  for (const [index, reduction] of [
    [prefix.indexOf('》'), o.closeReduced ?? 0],
    [prefix.indexOf('（'), o.openReduced ?? 0],
  ]) {
    const span = spans[index]!;
    span.style = withLayoutDerivedSpacing(span.style, -reduction);
    span.box.width -= reduction;
    if (index === prefix.indexOf('（')) span.glyphOffsetPt = -reduction;
  }
  x = 0;
  for (const span of spans) {
    span.box.x = x;
    x += span.box.width;
  }
  if (o.baselineDerived)
    for (const index of [prefix.indexOf('》'), prefix.indexOf('（')])
      pieces[index].style = spans[index].style;
  const width = x;
  const line: any = {
    spans,
    drawings: [],
    start: 0,
    end,
    width,
    height: 18.2,
    baseline: 14,
    leading: 0,
    trailingSpacing: 0,
  };
  const before = {
    ...line,
    spans: line.spans.map((s) => ({
      ...s,
      style: { ...s.style },
      box: { ...s.box },
      range: { ...s.range },
      props: [...s.props],
    })),
  };
  const available =
    o.deficit !== undefined
      ? width + m.measure(suffix, base) - o.deficit
      : o.halfReject
        ? width + 6.9
        : width + m.measure(suffix, base) - 18.6;
  const typography: any = {
    ...DEFAULT_CJK_TYPOGRAPHY,
    settings: {
      before: {},
      after: {},
      compression: 'compressPunctuation',
      sourceEastAsianFont: o.sourceEA ?? true,
      punctuationKerning: o.noPunc ? false : true,
    },
  };
  const breaks = cjkParagraphBreaks(pieces, typography);
  const enabled = o.policy ? !o.noPunc && (o.sourceEA ?? true) : (o.explicit ?? true);
  const f = createCjkOpticalFitter(pieces, breaks, m, false, true, enabled);
  const result = f(
    line,
    pieces.length - 1,
    suffix,
    { paragraphId: 'generic-pair', start: end, end: end + suffix.length },
    m.measure(suffix, base),
    width,
    available,
    true
  );
  return { result, line, before, available, prefix, suffix };
}
function reject(o: any) {
  const a = attempt(o);
  expect(a.result).toBeUndefined();
  expect(a.line).toEqual(a.before);
}
test('same-face fullwidth close-open pair shares both bearings within one11.6pt ink gap', () => {
  const a = attempt();
  expect(a.result).toBeDefined();
  const all = [...a.line.spans, ...a.result!.candidate];
  expect(a.line.width + a.result!.width).toBeCloseTo(a.available, 7);
  expect(all.map((x) => x.text).join('')).toBe(a.prefix + a.suffix);
  const close = all.find((x) => x.text === '》')!,
    open = all.find((x) => x.text === '（')!;
  expect(
    open.box.x + (open.glyphOffsetPt ?? 0) - close.box.x - (close.glyphOffsetPt ?? 0)
  ).toBeCloseTo(2.4, 7);
  expect(close.box.width).toBeCloseTo(8.2, 7);
  expect(open.box.width).toBeCloseTo(8.2, 7);
  expect(all.map((x) => [x.range.start, x.range.end])).toEqual(
    Array.from(a.prefix + a.suffix, (_, i) => [i, i + 1])
  );
  for (const s of all.filter((x) => /^[甲乙丙戊己]$/.test(x.text))) expect(s.box.width).toBe(14);
});
test('disabled factory policy retains the baseline pair geometry', () =>
  reject({ policy: true, noPunc: true }));
test('unqualified factory policy retains the baseline pair geometry', () =>
  reject({ policy: true, sourceEA: false }));
test('without explicit demand the new double-bearing path is unavailable', () =>
  reject({ explicit: false }));
test('different authored EA faces cannot share a pair budget', () =>
  reject({ openStyle: { fontFamilyEastAsia: 'Other EA' } }));
test('different authored sizes cannot share a pair budget', () =>
  reject({ openStyle: { fontSizePt: 12 } }));
test('different authored bold face cannot share a pair budget', () =>
  reject({ openStyle: { bold: true } }));
test('authored tracking cannot supply pair capacity', () =>
  reject({ openStyle: { characterSpacingPt: 0.25 } }));
test('authored horizontal scaling cannot supply pair capacity', () =>
  reject({ openStyle: { horizontalScalePercent: 90 } }));
test('unresolved EA faces cannot supply pair capacity', () => reject({ unresolved: true }));
test('source range gap cannot supply pair capacity', () => reject({ sourceGap: true }));
test('a measured pair gap below the required total is not counted twice', () =>
  reject({ tightGap: true }));
test('first authored glyph half-width remains mandatory despite pair capacity', () =>
  reject({ halfReject: true, suffix: '戊' }));
test('ASCII-slot punctuation cannot enter the EA pair route', () => reject({ asciiSlot: true }));

for (const [closeReduced, openReduced] of [
  [2, 3],
  [6, 0],
  [0, 6],
  [7, 0],
]) {
  test(`derived close ${closeReduced}/open ${openReduced} use only remaining per-cell half and one remaining ink gap`, () => {
    const a = attempt({ closeReduced, openReduced, deficit: 18.6 - closeReduced - openReduced });
    expect(a.result).toBeDefined();
    const all = [...a.line.spans, ...a.result!.candidate];
    const close = all.find((s) => s.text === '》')!,
      open = all.find((s) => s.text === '（')!;
    const beforeClose = a.before.spans[4]!,
      beforeOpen = a.before.spans[5]!;
    const additional =
      beforeClose.box.width - close.box.width + beforeOpen.box.width - open.box.width;
    expect(additional).toBeCloseTo(11.6 - closeReduced - openReduced, 7);
    for (const s of [close, open]) {
      expect(14 - s.box.width).toBeLessThanOrEqual(7 + 1e-8);
      expect(layoutDerivedSpacingStyleForEditing(s.style).characterSpacingPt).toBe(0);
    }
    expect(
      open.box.x + (open.glyphOffsetPt ?? 0) - close.box.x - (close.glyphOffsetPt ?? 0)
    ).toBeCloseTo(2.4, 7);
    expect(a.line.width + a.result!.width).toBeCloseTo(a.available, 7);
    expect(all.map((s) => s.text).join('')).toBe(a.prefix + a.suffix);
  });
}
test('already exhausted half-bearings do not become fresh pair capacity', () =>
  reject({ closeReduced: 7, openReduced: 7, deficit: 7.1 }));
test('existing pair reductions cannot count the same measured ink gap again', () =>
  reject({ closeReduced: 2, openReduced: 3, deficit: 13.7 }));

test('qualified close-open pair supplies a small demand without unrelated punctuation capacity', () => {
  const a = attempt({ prefix: '甲》（乙', deficit: 3 });
  expect(a.result).toBeDefined();
  expect(a.line.width + a.result!.width).toBeCloseTo(a.available, 7);
});
for (const options of [
  { openStyle: { fontFamilyEastAsia: 'Other EA' } },
  { asciiSlot: true },
  { unresolved: true },
  { sourceGap: true },
  { openStyle: { characterSpacingPt: 0.25 } },
  { openStyle: { fontSizePt: 12 } },
  { openStyle: { horizontalScalePercent: 90 } },
  { openStyle: { bold: true } },
  { openStyle: { shaping: { direction: 'ltr', script: 'Hani' } } },
])
  test(`unqualified low-demand pair adds no seam: ${JSON.stringify(options)}`, () =>
    reject({ prefix: '甲》（乙', deficit: 3, ...options }));

for (const closeReduced of [6, 7])
  test(`a derived baseline also retains the authored half-width after ${closeReduced}pt prior closing reduction`, () => {
    const a = attempt({ closeReduced, baselineDerived: true, deficit: 18.6 - closeReduced });
    expect(a.result).toBeDefined();
    const all = [...a.line.spans, ...a.result!.candidate];
    const close = all.find((s) => s.text === '》')!,
      opening = all.find((s) => s.text === '（')!;
    expect(14 - close.box.width).toBeLessThanOrEqual(7 + 1e-8);
    expect(14 - opening.box.width).toBeLessThanOrEqual(7 + 1e-8);
    expect(layoutDerivedSpacingStyleForEditing(close.style).characterSpacingPt).toBe(0);
    expect(a.line.width + a.result!.width).toBeCloseTo(a.available, 7);
  });

test('two exhausted authored half-bearings retain a positive ink gap without fresh capacity', () => {
  const a = attempt({
    prefix: '甲》（',
    closeReduced: 7,
    openReduced: 7,
    baselineDerived: true,
    saturatedInk: true,
    deficit: 3,
    suffix: '乙',
  });
  const close = a.before.spans.find((s) => s.text === '》')!,
    open = a.before.spans.find((s) => s.text === '（')!;
  expect(close.box.width).toBe(7);
  expect(open.box.width).toBe(7);
  expect(open.glyphOffsetPt).toBe(-7);
  expect(close.box.width + open.glyphOffsetPt + 12 - 3).toBe(9);
  expect(a.available - a.before.width).toBe(11);
  expect(a.result).toBeUndefined();
  expect(a.line).toEqual(a.before);
});
