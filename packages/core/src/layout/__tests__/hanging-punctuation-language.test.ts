import { expect, test } from 'bun:test';
import { loadFixture, measurer } from './punctuation-kerning-fixture.ts';
import { caretAt } from '../semantic-interaction.ts';
import { hangingCjkPunctuationAdvance } from '../cjk-spacing.ts';
import { piecesOfParagraph } from '../field-projection.ts';
import { serializeOoxmlPart } from '../../store/index.ts';

const H = '国'.repeat(19);
const properties = (jc = 'both', overflow = true) =>
  `<w:jc w:val="${jc}"/><w:overflowPunct w:val="${overflow ? 1 : 0}"/>`;
const texts = (result: ReturnType<typeof loadFixture>) =>
  result.lines.map((line) => line.spans.map((span) => span.text).join(''));

for (const jc of ['both', 'distribute']) {
  for (const mark of [...'〕》〉！？：；・']) {
    test(`${jc}: Japanese ${mark} wraps with its preceding ideograph`, () => {
      const result = loadFixture({
        text: H + mark + '国',
        available: 468,
        language: 'ja-JP',
        paragraphProperties: properties(jc),
      });
      expect(texts(result)).toEqual([H.slice(1), '国' + mark + '国']);
    });
  }
  for (const mark of ['。', '！', '.', ')']) {
    test(`${jc}: Korean ${mark} starts the next line`, () => {
      const result = loadFixture({
        text: H + mark + '国',
        available: 468,
        language: 'ko-KR',
        paragraphProperties: properties(jc),
      });
      expect(texts(result)).toEqual([H, mark + '国']);
    });
  }
  test(`${jc}: a hanging final mark leaves a full-width prefix`, () => {
    const result = loadFixture({
      text: H + '。',
      available: 468,
      language: 'ja-JP',
      paragraphProperties: properties(jc),
    });
    expect(texts(result)).toEqual([H + '。']);
    const line = result.lines[0]!;
    const id = line.range.paragraphId;
    expect(
      caretAt(result.layout, { paragraphId: id, offset: 19 }, measurer)!.x - line.box.x
    ).toBeCloseTo(468, 6);
    expect(
      caretAt(result.layout, { paragraphId: id, offset: 20 }, measurer)!.x - line.box.x
    ).toBeCloseTo(492, 6);
  });
}

test('the mark-owning run supplies its language across a source seam', () => {
  const japanese = loadFixture({
    text: H + '！国',
    split: 19,
    language: 'zh-CN',
    secondLanguage: 'ja-JP',
    available: 468,
    paragraphProperties: properties(),
  });
  expect(texts(japanese)).toEqual([H.slice(1), '国！国']);
  const chinese = loadFixture({
    text: H + '！国',
    split: 19,
    language: 'ja-JP',
    secondLanguage: 'zh-CN',
    available: 468,
    paragraphProperties: properties(),
  });
  expect(texts(chinese)).toEqual([H + '！', '国']);
});

test('an inherited Japanese language controls its punctuation', () => {
  const result = loadFixture({
    text: H + '！国',
    defaultLanguage: 'ja-JP',
    available: 468,
    paragraphProperties: properties(),
  });
  expect(texts(result)).toEqual([H.slice(1), '国！国']);
});

test('Chinese fullwidth dots are refused by the common hanging admission helper', () => {
  const result = loadFixture({
    text: '国．国',
    language: 'zh-CN',
    available: 468,
    paragraphProperties: properties(),
  });
  const pkg = result.view.currentPackage();
  const part = pkg.parts.get(pkg.mainDocumentPart)!;
  const paragraph = part.root.children
    .flatMap((node) => ('children' in node ? node.children : []))
    .find((node) => node.kind === 'paragraph')!;
  const piece = piecesOfParagraph(paragraph)[0]!;
  const before = serializeOoxmlPart(part);
  expect(hangingCjkPunctuationAdvance('国．', piece, measurer)).toBeUndefined();
  expect(serializeOoxmlPart(part)).toBe(before);
});

test('ordinary justification and disabled hanging retain their policies', () => {
  const off = loadFixture({
    text: H + '。国',
    language: 'ja-JP',
    available: 468,
    paragraphProperties: properties('both', false),
  });
  expect(texts(off)).toEqual([H.slice(1), '国。国']);
  const last = loadFixture({
    text: '国国',
    language: 'ja-JP',
    available: 468,
    paragraphProperties: properties(),
  });
  expect(last.width).toBe(48);
  const allowed = loadFixture({
    text: H + '。国',
    language: 'ja-JP',
    available: 468,
    paragraphProperties: properties(),
  });
  expect(texts(allowed)).toEqual([H + '。', '国']);
});

test('a Chinese final fullwidth dot wraps instead of hanging', () => {
  const result = loadFixture({
    text: H + '．',
    language: 'zh-CN',
    available: 468,
    paragraphProperties: properties(),
  });
  expect(texts(result)).toEqual([H.slice(1), '国．']);
});

test('Korean custom line-start prohibitions and nonjustified rules remain active', () => {
  const custom = loadFixture({
    text: H + '！国',
    language: 'ko-KR',
    available: 468,
    paragraphProperties: properties(),
    lineStartRule: { language: 'ko-kr', characters: '！' },
  });
  expect(texts(custom)).toEqual([H.slice(1), '国！国']);
  const left = loadFixture({
    text: H + '！国',
    language: 'ko-KR',
    available: 468,
    paragraphProperties: properties('left'),
  });
  expect(texts(left)).toEqual([H.slice(1), '国！国']);
});
