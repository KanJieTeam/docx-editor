import { expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart } from '../../store/index.ts';
import { layoutSemanticDocument } from '../semantic-layout.ts';
import { buildStyleCascadeTable } from '../style-cascade.ts';
import { NO_THEME_FONTS } from '../run-style.ts';
import { linesOf, type TextMeasurer } from '../semantic-records.ts';
import { automaticTextSpacingBefore } from '../cjk-automatic-spacing.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
// Proportional synthetic ink. No installed font or reference PDF is used here.
const measurer: TextMeasurer = {
  measure: (t, s) =>
    t.length * ((s.fontSizePt * s.horizontalScalePercent) / 100 + s.characterSpacingPt),
  lineMetrics: () => ({ height: 24, baseline: 19.2 }),
  hasResolvedFont: () => true,
  inkBounds: (t, s) =>
    t === '》'
      ? { left: s.fontSizePt * 0.1, right: s.fontSizePt * 0.6 }
      : t === '（'
        ? { left: s.fontSizePt * 0.4, right: s.fontSizePt * 0.95 }
        : { left: 0, right: s.fontSizePt },
};
function part(xml: string, name: string) {
  const p = readOoxmlPart(xml, { name, contentType: 'application/xml' });
  if (!p.ok) throw Error(p.reason);
  return p.part;
}
interface Options {
  alignment?: string;
  inherited?: boolean;
  noPunc?: boolean;
  ea?: boolean;
  openingSize?: number;
  closingSize?: number;
  wide?: boolean;
  kern?: boolean;
  ordinaryTail?: boolean;
  opaqueLatin?: boolean;
  unpairedOpening?: boolean;
  deficit?: number;
  noInk?: boolean;
  automaticSpacing?: boolean | 'default';
  reserveReverseGap?: boolean;
}
function layout(o: Options = {}) {
  const opening = o.openingSize ?? 12,
    closing = o.closingSize ?? opening;
  const alignment = o.alignment ?? 'both';
  const jc = `<w:jc w:val="${alignment}"/>`;
  // Pair-only controls disable automatic gaps explicitly. The 'default' controls
  // omit both flags and separately verify the source's default-on behavior.
  const automaticSpacing =
    o.automaticSpacing === 'default'
      ? ''
      : `<w:autoSpaceDE w:val="${o.automaticSpacing ? 1 : 0}"/><w:autoSpaceDN w:val="${o.automaticSpacing ? 1 : 0}"/>`;
  const run = (t: string, size: number) =>
    `<w:r><w:rPr><w:rFonts w:eastAsia="Synthetic EA" w:hint="eastAsia"/><w:sz w:val="${size * 2}"/>${o.kern ? '<w:kern w:val="2"/>' : ''}</w:rPr><w:t>${t}</w:t></w:r>`;
  const tail = o.opaqueLatin ? 'AB乙。' : o.ordinaryTail ? '乙丙' : '乙丙。';
  const source = part(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:pPr>${o.inherited ? '<w:pStyle w:val="SyntheticJustify"/>' : jc}<w:overflowPunct w:val="0"/>${automaticSpacing}</w:pPr>${run('甲', opening)}${o.unpairedOpening ? '' : run('》', closing)}${run('（', opening)}${run(tail, opening)}</w:p></w:body></w:document>`,
    '/word/document.xml'
  );
  const styles = part(
    `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>${o.ea === false ? '' : '<w:rFonts w:eastAsia="Synthetic EA"/>'}</w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="SyntheticJustify"><w:pPr>${jc}</w:pPr></w:style></w:styles>`,
    '/word/styles.xml'
  );
  const settings = part(
    `<w:settings xmlns:w="${W}"><w:characterSpacingControl w:val="compressPunctuation"/>${o.noPunc ? '<w:noPunctuationKerning/>' : ''}</w:settings>`,
    '/word/settings.xml'
  );
  const sourceBefore = serializeOoxmlPart(source);
  const natural = opening * (2 + tail.length) + (o.unpairedOpening ? 0 : closing);
  // Latin runs inherit Times New Roman. Its recorded regular average is 821/2048 em;
  // the reverse gap is half that value. This is metadata, not an installed-font test.
  const reverseGap = (opening * 821) / 4096;
  const width =
    (o.wide ? natural + 20 : natural - (o.deficit ?? opening * 0.75)) +
    (o.reserveReverseGap ? reverseGap : 0);
  const model = layoutSemanticDocument(source, 0, {
    measurer: o.noInk ? { ...measurer, inkBounds: undefined } : measurer,
    styleCascade: buildStyleCascadeTable(
      styles.root,
      { ...NO_THEME_FONTS, eastAsianFontTableFaces: ['synthetic ea'] },
      settings.root
    ),
    geometry: {
      width: width + 20,
      height: 500,
      margin: { top: 10, bottom: 10, left: 10, right: 10 },
    },
  });
  const lines = linesOf(model),
    spans = lines.flatMap((line) => line.spans);
  expect(spans.map((span) => span.text).join('')).toBe(
    '甲' + (o.unpairedOpening ? '' : '》') + '（' + tail
  );
  expect(serializeOoxmlPart(source)).toBe(sourceBefore);
  expect(sourceBefore).not.toContain('<w:spacing');
  expect(sourceBefore).not.toContain('<w:w ');
  // Equal-width synthetic clusters can share a natural span before any fit occurs.
  const character = (text: string) => {
    const span = spans.find((item) => item.text.includes(text));
    return span
      ? { ...span, box: { ...span.box, width: span.box.width / span.text.length } }
      : undefined;
  };
  const close = character('》');
  const open = character('（');
  // Only B and 乙 use this coordinate helper. These synthetic BMP glyphs have
  // equal advances within each uncompressed Latin or ordinary EA span.
  const interval = (text: 'B' | '乙') => {
    const span = spans.find((item) => item.text.includes(text))!;
    const width = span.box.width / span.text.length;
    const left = span.box.x + span.text.indexOf(text) * width;
    return { left, right: left + width };
  };
  return { lines, spans, close, open, closing, opening, natural, interval, reverseGap };
}
for (const openingSize of [10.5, 12])
  for (const closingSize of [openingSize, 20]) {
    for (const noPunc of [false, true])
      test(`parsed justified ${closingSize}/${openingSize} pt pair fits; noPunc=${noPunc}`, () => {
        const a = layout({ openingSize, closingSize, noPunc });
        expect(a.lines).toHaveLength(1);
        expect(a.close).toBeDefined();
        expect(a.open).toBeDefined();
        expect(a.close!.box.width).toBeLessThan(a.closing);
        expect(a.open!.box.width).toBeLessThan(a.opening);
      });
  }
test('parsed inherited justified alignment activates the measured pair', () =>
  expect(layout({ inherited: true, noPunc: true }).lines).toHaveLength(1));

for (const noPunc of [false, true])
  test(`a qualified pair on an opaque Latin line fits without borrowing Latin width; noPunc=${noPunc}`, () => {
    const a = layout({ opaqueLatin: true, noPunc });
    expect(a.lines).toHaveLength(1);
    expect(
      a.spans.filter((span) => /^[AB]+$/.test(span.text)).reduce((n, span) => n + span.box.width, 0)
    ).toBeCloseTo(24, 7);
    expect(a.close!.box.width).toBeLessThan(12);
    expect(a.open!.box.width).toBeLessThan(12);
  });

const latinAdvance = (a: ReturnType<typeof layout>) =>
  a.spans.filter((span) => /^[AB]+$/.test(span.text)).reduce((n, span) => n + span.box.width, 0);

for (const noPunc of [false, true]) {
  test(`an enabled or default reverse gap cannot rescue an overwide pair; noPunc=${noPunc}`, () => {
    for (const automaticSpacing of [true, 'default'] as const) {
      const a = layout({ opaqueLatin: true, noPunc, automaticSpacing });
      // 84pt text + 2.4052734375pt gap - 75pt measure exceeds the 9.6pt pair budget.
      expect(a.lines).toHaveLength(2);
      expect(a.reverseGap).toBeCloseTo(2.4052734375, 10);
      expect(latinAdvance(a)).toBeCloseTo(24, 7);
    }
  });

  test(`reserving the full reverse gap leaves the measured pair budget unchanged; noPunc=${noPunc}`, () => {
    for (const automaticSpacing of [true, 'default'] as const) {
      const a = layout({ opaqueLatin: true, noPunc, automaticSpacing, reserveReverseGap: true });
      expect(a.lines).toHaveLength(1);
      expect(latinAdvance(a)).toBeCloseTo(24, 7);
      const after = a.spans.find((span) => span.text.startsWith('乙'))!;
      expect(automaticTextSpacingBefore(after)).toBeCloseTo(a.reverseGap, 10);
      expect(a.interval('乙').left - a.interval('B').right).toBeCloseTo(a.reverseGap, 10);
      expect(a.closing - a.close!.box.width).toBeCloseTo(4.5, 7);
      expect(a.opening - a.open!.box.width).toBeCloseTo(4.5, 7);
    }
  });

  test(`wide left-aligned runs retain the whole reverse gap and Latin advance; noPunc=${noPunc}`, () => {
    const off = layout({ opaqueLatin: true, noPunc, alignment: 'left', wide: true });
    for (const automaticSpacing of [true, 'default'] as const) {
      const on = layout({
        opaqueLatin: true,
        noPunc,
        alignment: 'left',
        wide: true,
        automaticSpacing,
      });
      expect(on.lines).toHaveLength(1);
      expect(latinAdvance(on)).toBeCloseTo(24, 7);
      expect(on.spans.find((span) => span.text.includes('B'))!.style.fontFamily).toBe(
        'Times New Roman'
      );
      expect(on.interval('B')).toEqual(off.interval('B'));
      expect(on.interval('乙').left - off.interval('乙').left).toBeCloseTo(on.reverseGap, 10);
      expect(on.interval('乙').left - on.interval('B').right).toBeCloseTo(on.reverseGap, 10);
      expect(
        automaticTextSpacingBefore(on.spans.find((span) => span.text.startsWith('乙'))!)
      ).toBeCloseTo(on.reverseGap, 10);
    }
  });
}

test('an automatic reverse gap can also exhaust an unpaired opening budget', () => {
  const a = layout({
    opaqueLatin: true,
    unpairedOpening: true,
    deficit: 3,
    automaticSpacing: true,
  });
  // 3pt deficit + 2.4052734375pt gap exceeds the ordinary opening's 4.8pt ink budget.
  expect(a.lines).toHaveLength(2);
  expect(latinAdvance(a)).toBeCloseTo(24, 7);
});

for (const noPunc of [false, true])
  test(`an opaque line without a close-open pair keeps the old local noPunc policy; noPunc=${noPunc}`, () => {
    const a = layout({ opaqueLatin: true, unpairedOpening: true, noPunc, deficit: 3 });
    expect(a.lines).toHaveLength(noPunc ? 2 : 1);
    expect(
      a.spans.filter((span) => /^[AB]+$/.test(span.text)).reduce((n, span) => n + span.box.width, 0)
    ).toBeCloseTo(24, 7);
  });
test('parsed missing EA defaults retain natural pair advances and wrap', () => {
  const a = layout({ ea: false });
  expect(a.lines).toHaveLength(2);
  expect(a.close!.box.width).toBe(a.closing);
  expect(a.open!.box.width).toBe(a.opening);
});
for (const alignment of ['left', 'center', 'right'])
  for (const kern of [false, true]) {
    test(`parsed ${alignment} pair keeps natural widths even with kern=${kern}`, () => {
      for (const wide of [false, true]) {
        const a = layout({ alignment, kern, wide });
        expect(a.lines).toHaveLength(wide ? 1 : 2);
        expect(a.close!.box.width).toBe(a.closing);
        expect(a.open!.box.width).toBe(a.opening);
      }
    });
  }
for (const closingSize of [12, 20])
  test(`parsed prior pair kerning cannot overspend an opening-size budget; close=${closingSize}`, () => {
    const a = layout({ closingSize, kern: true });
    expect(a.lines).toHaveLength(1);
    expect(a.closing - a.close!.box.width).toBeLessThanOrEqual(a.opening * 0.4 + 0.001);
    expect(a.opening - a.open!.box.width).toBeLessThanOrEqual(a.opening * 0.4 + 0.001);
  });

test('the old five-character fixture still wraps when an ordinary glyph has less than half its advance', () => {
  const a = layout({ ordinaryTail: true });
  expect(a.natural).toBe(60);
  expect(a.lines).toHaveLength(2);
});

test('explicit justified flow without measured ink cannot fall back to the legacy closing half', () => {
  const a = layout({ kern: true, noInk: true });
  expect(a.lines).toHaveLength(2);
  expect(a.close!.box.width).toBe(12);
  expect(a.open!.box.width).toBe(12);
});
