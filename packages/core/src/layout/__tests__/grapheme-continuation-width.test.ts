import { expect, test } from 'bun:test';
import { readOoxmlPart } from '../../store/index.ts';
import { breakParagraph } from '../paragraph-flow.ts';
import { createFixedMeasurer } from '../fixed-measurer.ts';
import { graphemeContinuationWidth } from '../grapheme-continuation-width.ts';
import { DEFAULT_RUN_STYLE } from '../run-style.ts';
import { resetGraphemeBoundary, setGraphemeBoundary } from '../grapheme.ts';

const measurer = createFixedMeasurer(6, 14);
function render(parts: readonly string[], width: number) {
  const read = readOoxmlPart(
    '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      parts
        .map(
          (text) =>
            '<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">' +
            text +
            '</w:t></w:r>'
        )
        .join('') +
      '</w:p>',
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!read.ok) throw Error(read.reason);
  return breakParagraph(read.part.root, 'p', 0, width, measurer, undefined, null).map((line) => ({
    text: line.spans.map((span) => span.text).join(''),
    width: line.width,
    cells: line.spans.flatMap((span) =>
      Array.from({ length: span.text.length }, (_, index) => [
        span.range.start + index,
        span.box.x + measurer.measure(span.text.slice(0, index), span.style),
      ])
    ),
  }));
}
for (const parts of [
  ['国abcd', 'efgh国'],
  ['e\u0301 AB', 'CD e\u0301'],
  ['👩‍💻 ABC', 'DEF 👩‍💻'],
]) {
  test(`ASCII run boundaries keep internal Unicode geometry: ${parts.join('|')}`, () => {
    for (const width of [18, 30, 54])
      expect(render(parts, width)).toEqual(render([parts.join('')], width));
  });
}
test('a custom grapheme strategy can join an ASCII run boundary', () => {
  setGraphemeBoundary({
    segment: (text) => [{ index: 0, text, utf16From: 0, utf16To: text.length }],
  });
  try {
    const style = { ...DEFAULT_RUN_STYLE, fontSizePt: 11 };
    const pieces = [
      { text: 'AB', start: 0, end: 2, style, props: [] },
      { text: 'CD', start: 2, end: 4, style, props: [] },
    ];
    expect(graphemeContinuationWidth(pieces, measurer)(0, 2)).toBe(12);
    expect(render(['国AB', 'CD国'], 12)).toEqual(render(['国ABCD国'], 12));
  } finally {
    resetGraphemeBoundary();
  }
});
