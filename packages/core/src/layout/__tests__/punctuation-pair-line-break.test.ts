import { describe, expect, test } from 'bun:test';
import { loadFixture } from './punctuation-kerning-fixture.ts';

// Fixed 24pt advances isolate source layout policy in a 468pt column.
// This fixture has no resolved-face or ink witness. Close-open pairs stay natural;
// other qualified legacy punctuation seams can still share a 12pt bearing.
const H = (count: number) => '国'.repeat(count);
const LEFT = '<w:jc w:val="left"/>';
const JUSTIFIED = '<w:jc w:val="both"/>';

function lay(text: string, paragraphProperties = LEFT) {
  const { lines } = loadFixture({
    text,
    kern: 2,
    available: 468,
    language: 'ja-JP',
    paragraphProperties,
  });
  return lines.map((line) => {
    const chars: { char: string; x: number; width: number }[] = [];
    for (const span of line.spans) {
      const count = [...span.text].length;
      [...span.text].forEach((char, index) =>
        chars.push({
          char,
          x: span.box.x + (index * span.box.width) / count,
          width: span.box.width / count,
        })
      );
    }
    const last = line.spans.at(-1)!;
    return { text: chars.map((c) => c.char).join(''), chars, end: last.box.x + last.box.width };
  });
}
const texts = (lines: ReturnType<typeof lay>) => lines.map((line) => line.text);
const ends = (lines: ReturnType<typeof lay>) => lines.map((line) => line.end);

describe('unmeasured close-open pairs retain natural advances and legal line breaks', () => {
  test.each([
    [15, [H(15) + '」「」', '「国国国'], [432, 96]],
    [16, [H(16) + '」「」', '「国国国'], [456, 96]],
    [17, [H(17) + '」', '「」「国国国'], [432, 144]],
    [18, [H(18) + '」', '「」「国国国'], [456, 144]],
    [19, [H(18), '国」「」「国国国'], [432, 192]],
  ])('%d ideographs before 」「」「国国国', (count, expected, widths) => {
    const lines = lay(H(count) + '」「」「国国国');
    expect(texts(lines)).toEqual(expected);
    expect(ends(lines)).toEqual(widths);
  });

  test('a closing bracket at the line end keeps its full advance', () => {
    const [first, second] = lay(H(18) + '」「」「国国国');
    expect(first!.chars.at(-1)).toEqual({ char: '」', x: 432, width: 24 });
    // Left alignment does not reduce the pair on the next line either.
    expect(second!.chars.slice(0, 3).map((c) => c.x)).toEqual([0, 24, 48]);
  });

  test.each([
    [H(18) + '。」「国', [H(18) + '。」', '「国'], [468, 48]],
    [H(17) + '」「国国', [H(17) + '」', '「国国'], [432, 72]],
    [H(18) + '」）国', [H(18) + '」）', '国'], [468, 24]],
    [H(19) + '。」国', [H(18), '国。」国'], [432, 84]],
  ])('%s', (text, expected, widths) => {
    const lines = lay(text);
    expect(texts(lines)).toEqual(expected);
    expect(ends(lines)).toEqual(widths);
  });

  test('a close-open pair stays natural after a colon retry moves it off the line end', () => {
    const lines = lay(H(13) + '：「'.repeat(4) + H(17) + '）（国');
    const last = lines.at(-1)!;
    expect(last.text).toBe('国）（国');
    expect(last.chars.map((c) => c.x)).toEqual([0, 24, 48, 72]);
  });

  test('right alignment places a line-final closing bracket at its full advance', () => {
    const [first] = lay(H(18) + '」「」「国国国', '<w:jc w:val="right"/>');
    expect(first!.text).toBe(H(18) + '」');
    expect(first!.chars[0]!.x).toBe(12);
    expect(first!.chars.at(-1)).toEqual({ char: '」', x: 444, width: 24 });
  });
});

describe('punctuation hangs past the margin only on justified lines', () => {
  test.each(['。', '、', '，', '．', '」', '）', '』', '！', '？', '：'])(
    'left-aligned %s moves with the previous character',
    (mark) => {
      expect(texts(lay(H(19) + mark + '国'))).toEqual([H(18), '国' + mark + '国']);
      expect(texts(lay(H(19) + mark + '国', '<w:overflowPunct/>'))).toEqual([
        H(18),
        '国' + mark + '国',
      ]);
    }
  );

  test.each(['center', 'right'])('%s alignment does not hang punctuation', (jc) => {
    expect(texts(lay(H(19) + '。国', `<w:jc w:val="${jc}"/>`))).toEqual([H(18), '国。国']);
  });

  test.each(['。', '、', '，', '．', '」', '）', '』', '】', '］', '｝'])(
    'justified %s hangs',
    (mark) => {
      const lines = lay(H(19) + mark + '国', JUSTIFIED);
      expect(texts(lines)).toEqual([H(19) + mark, '国']);
      expect(lines[0]!.chars.at(-1)!.x).toBeGreaterThanOrEqual(456);
    }
  );

  test('justified hanging does not authorize an inner pair without measured ink', () => {
    const lines = lay(H(18) + '」「」「国国国', JUSTIFIED);
    expect(texts(lines)).toEqual([H(18) + '」', '「」「国国国']);
    // Natural width is 19 * 24 = 456. Justification adds 468 - 456 = 12pt
    // before the closing glyph: its origin is 18 * 24 + 12 = 444pt.
    expect(lines[0]!.chars.at(-1)).toEqual({ char: '」', x: 444, width: 24 });
    expect(lines[0]!.end).toBe(468);
    expect(lines[1]!.chars.map((c) => c.x)).toEqual([0, 24, 48, 72, 96, 120]);
    expect(texts(lay(H(19) + '」「」「国国国', JUSTIFIED))).toEqual([H(19) + '」', '「」「国国国']);
  });

  test('a justified paragraph without overflow punctuation does not hang', () => {
    const off = JUSTIFIED + '<w:overflowPunct w:val="0"/>';
    expect(texts(lay(H(19) + '。国', off))).toEqual([H(18), '国。国']);
    expect(texts(lay(H(19) + '」「」「国国国', off))).toEqual([H(18), '国」「」「国国国']);
  });
});
