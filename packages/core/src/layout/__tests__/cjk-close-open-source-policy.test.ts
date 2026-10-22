import { expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart } from '../../store/index.ts';
import { layoutSemanticDocument } from '../semantic-layout.ts';
import { buildStyleCascadeTable } from '../style-cascade.ts';
import { NO_THEME_FONTS } from '../run-style.ts';
import { linesOf, type TextMeasurer } from '../semantic-records.ts';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const text = '甲》（乙戊';
const m: TextMeasurer = {
  measure: (t, s) =>
    (t.length * s.fontSizePt * s.horizontalScalePercent) / 100 + t.length * s.characterSpacingPt,
  lineMetrics: () => ({ height: 18.2, baseline: 14 }),
  hasResolvedFont: () => true,
  inkBounds: (t) =>
    t === '》'
      ? { left: 1, right: 9.4 }
      : /^[《（]$/.test(t)
        ? { left: 7, right: 13 }
        : t === '。'
          ? { left: 0, right: 4 }
          : { left: 0, right: 12 },
};
function part(xml: string) {
  const r = readOoxmlPart(xml, { name: '/word/document.xml', contentType: 'application/xml' });
  if (!r.ok) throw Error(r.reason);
  return r.part;
}
function layout(noPunc = false, ea = true, kern?: number, width = 64) {
  const source = part(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:overflowPunct w:val="0"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="Invented EA" w:hint="eastAsia"/><w:sz w:val="28"/>${kern === undefined ? '' : `<w:kern w:val="${kern}"/>`}</w:rPr><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`
  );
  const settings = part(
    `<w:settings xmlns:w="${W}"><w:characterSpacingControl w:val="compressPunctuation"/>${noPunc ? '<w:noPunctuationKerning/>' : ''}</w:settings>`
  );
  const styles = part(
    `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>${ea ? '<w:rFonts w:eastAsia="Invented EA"/>' : ''}</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`
  );
  const before = serializeOoxmlPart(source);
  const lines = linesOf(
    layoutSemanticDocument(source, 0, {
      measurer: m,
      styleCascade: buildStyleCascadeTable(
        styles.root,
        { ...NO_THEME_FONTS, eastAsianFontTableFaces: ['invented ea'] },
        settings.root
      ),
      geometry: {
        width: width + 20,
        height: 500,
        margin: { top: 10, bottom: 10, left: 10, right: 10 },
      },
    })
  );
  expect(serializeOoxmlPart(source)).toBe(before);
  expect(
    lines
      .flatMap((l) => l.spans)
      .map((s) => s.text)
      .join('')
  ).toBe(text);
  return lines;
}
test('parsed EA defaults and enabled punctuation policy activate the pair in paragraph flow', () =>
  expect(layout()).toHaveLength(1));
test('parsed direct noPunctuationKerning prevents a new pair seam', () =>
  expect(layout(true)).toHaveLength(2));
test('missing source EA defaults prevent a new pair seam', () =>
  expect(layout(false, false)).toHaveLength(2));

test('parsed font pair kerning and demand fitting do not spend the closing half twice', () => {
  const lines = layout(false, true, 2, 58.4);
  expect(lines).toHaveLength(1);
  const spans = lines.flatMap((l) => l.spans);
  const close = spans.find((s) => s.text === '》')!,
    opening = spans.find((s) => s.text === '（')!;
  expect(close.box.width).toBeCloseTo(7, 7);
  expect(opening.box.width).toBeCloseTo(9.4, 7);
});
