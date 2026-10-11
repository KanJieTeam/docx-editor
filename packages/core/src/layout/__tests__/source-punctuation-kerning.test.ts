import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { readOoxmlPart } from '../../store/index.ts';
import { DEFAULT_RUN_STYLE } from '../run-style.ts';
import { DEFAULT_CJK_TYPOGRAPHY, cjkTypographyFromSettings } from '../cjk-typography.ts';
import { kernSourcePunctuationPieces } from '../source-punctuation-kerning.ts';
import { withResolvedPunctuationFace } from '../resolved-punctuation-face.ts';
import { layoutDerivedSpacingStyleForEditing } from '../layout-derived-spacing.ts';
import type { FieldAwarePiece } from '../field-pieces.ts';
import type { TextMeasurer } from '../semantic-records.ts';
import { hasSourcePunctuationFace } from '../sfnt-source-punctuation-face.ts';
import {
  createFontResourceSnapshot,
  boundedStructuralFontValidator,
  sha256FontBytes,
} from '../font-resource.ts';

// Deterministic protocol tests. This is not native font admission evidence.
const face = {
  hash: 'sha256:' + 'a'.repeat(64),
  faceIndex: 0,
  family: 'SimSun',
  weight: 400,
  style: 'normal',
};
const base: TextMeasurer = {
  measure: (text, style) => text.length * (style.fontSizePt + style.characterSpacingPt),
  lineMetrics: (style) => ({ height: style.fontSizePt, baseline: style.fontSizePt }),
};
const measurer = withResolvedPunctuationFace(base, () => face);
const policy = DEFAULT_CJK_TYPOGRAPHY;
const pieces = (pair: string): FieldAwarePiece[] => [
  {
    text: '甲乙' + pair + '1）丙丁',
    props: [],
    start: 0,
    end: 8,
    fontSlot: 'eastAsia',
    style: {
      ...DEFAULT_RUN_STYLE,
      fontFamily: 'Times New Roman',
      fontFamilyEastAsia: 'SimSun',
      fontSizePt: 14,
      kerningMinPt: 1,
    },
  },
];

test.each(['》（', '）（'])('source kern keeps %s offsets and authored spacing', (pair) => {
  const input = pieces(pair),
    before = JSON.stringify(input);
  const result = kernSourcePunctuationPieces(input, policy, measurer);
  expect(result.map((p) => p.text).join('')).toBe(input[0]!.text);
  const left = result.find((p) => p.start === 2)!;
  expect(left.style.characterSpacingPt).toBe(-7);
  expect(left.end).toBe(3);
  expect(layoutDerivedSpacingStyleForEditing(left.style).characterSpacingPt).toBe(0);
  expect(JSON.stringify(input)).toBe(before);
  expect(kernSourcePunctuationPieces(result, policy, measurer)).toBe(result);
});

test.each(['》（', '）（'])(
  'source kern %s rejects threshold, tracking and unsupported paths',
  (pair) => {
    const input = pieces(pair);
    for (const style of [
      { kerningMinPt: 0 },
      { kerningMinPt: 30 },
      { characterSpacingPt: 1 },
      { horizontalScalePercent: 80 },
      { bold: true },
      { italic: true },
      { smallCaps: true },
    ]) {
      const changed = [{ ...input[0]!, style: { ...input[0]!.style, ...style } }];
      expect(kernSourcePunctuationPieces(changed, policy, measurer)).toBe(changed);
    }
    expect(kernSourcePunctuationPieces(input, policy, base)).toBe(input);
    expect(
      kernSourcePunctuationPieces(
        input,
        {
          ...policy,
          settings: {
            strict: false,
            compression: 'doNotCompress',
            before: {},
            after: {},
            punctuationKerning: false,
          },
        },
        measurer
      )
    ).toBe(input);
    for (const changedFace of [
      { ...face, family: 'Other' },
      { ...face, weight: 700 },
      { ...face, style: 'italic' },
    ]) {
      expect(
        kernSourcePunctuationPieces(
          input,
          policy,
          withResolvedPunctuationFace(base, () => changedFace)
        )
      ).toBe(input);
    }
    const projected = [{ ...input[0]!, projected: true }];
    expect(kernSourcePunctuationPieces(projected, policy, measurer)).toBe(projected);
  }
);

test('split runs use the same actual face and preserve the private measurer capability', () => {
  const input = pieces('）（'),
    p = input[0]!;
  const split = [
    { ...p, text: p.text.slice(0, 3), end: 3 },
    { ...p, text: p.text.slice(3), start: 3 },
  ];
  const result = kernSourcePunctuationPieces(split, policy, { ...measurer });
  expect(result.find((p) => p.start === 2)!.style.characterSpacingPt).toBe(-7);
  const different = withResolvedPunctuationFace(base, (_style, text) => ({
    ...face,
    faceIndex: text === '（' ? 1 : 0,
  }));
  expect(kernSourcePunctuationPieces(split, policy, different)).toBe(split);
});

test.each(['', ' w:val="1"', ' w:val="true"'])(
  'direct noPunctuationKerning%s disables source kern',
  (attrs) => {
    const parsed = readOoxmlPart(
      `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:noPunctuationKerning${attrs}/></w:settings>`,
      { name: '/word/settings.xml', contentType: 'application/xml' }
    );
    expect(parsed.ok).toBe(true);
    if (!parsed.ok) return;
    expect(cjkTypographyFromSettings(parsed.part.root).punctuationKerning).toBe(false);
  }
);

test('misplaced flags reject only this subset and leave canonical XML intact', () => {
  const parsed = readOoxmlPart(
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:noPunctuationKerning w:val="0"/></w:compat></w:settings>',
    { name: '/word/settings.xml', contentType: 'application/xml' }
  );
  expect(parsed.ok).toBe(true);
  if (!parsed.ok) return;
  const before = JSON.stringify(parsed.part);
  expect(cjkTypographyFromSettings(parsed.part.root).punctuationKerning).toBe(false);
  expect(JSON.stringify(parsed.part)).toBe(before);
});

test('an admitted public font renamed SimSun cannot claim the binary face', () => {
  const bytes = new Uint8Array(
    readFileSync(new URL('./fixtures/fonts/DejaVuSans.ttf', import.meta.url))
  );
  const fonts = createFontResourceSnapshot({
    epoch: 1,
    maxFontBytes: 2_000_000,
    resources: [
      {
        request: { family: 'SimSun', weight: 400, style: 'normal' },
        id: 'public-renamed-font',
        bytes,
        hash: sha256FontBytes(bytes),
        faceIndex: 0,
      },
    ],
    validateFont: boundedStructuralFontValidator,
  });
  const font = fonts.resolve({ family: 'SimSun', weight: 400, style: 'normal' });
  expect(font instanceof Error).toBe(false);
  if (font instanceof Error) return;
  expect(hasSourcePunctuationFace(font)).toBe(false);
  expect(hasSourcePunctuationFace(font)).toBe(false);
});
