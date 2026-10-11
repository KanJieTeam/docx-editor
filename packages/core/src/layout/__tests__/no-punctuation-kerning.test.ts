import { afterAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import {
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlPart,
} from '../../store/package/ooxml-tree.ts';
import { cjkTypographyFromSettings } from '../cjk-typography.ts';
import { buildStyleCascadeTable } from '../style-cascade.ts';
import { layoutSemanticDocument } from '../semantic-layout.ts';
import { createLayoutShaping, disposeLayoutShaping } from '../layout-shaping.ts';
import { createFixedMeasurer } from '../fixed-measurer.ts';
import { createShapedMeasurer } from '../shaped-measurer.ts';
import { resolveRunStyle, runStylesEqual } from '../run-style.ts';
import { shapeLayoutStyleRun } from '../layout-run-shape.ts';
import {
  punctuationKerningRanges,
  punctuationKerningSegments,
  withPunctuationKerningPolicy,
} from '../punctuation-kerning.ts';
import type { FieldAwarePiece } from '../field-pieces.ts';
import { sha256FontBytes } from '../font-resource.ts';
import { tryCreateCanvasMeasurer, type CanvasTextContext } from '../canvas-measurer.ts';

const WML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function parse(xml: string, name: string): OoxmlPart {
  const result = readOoxmlPart(xml, { name, contentType: 'application/xml' });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
function settings(content = '', namespace = WML) {
  return parse(`<w:settings xmlns:w="${namespace}">${content}</w:settings>`, '/word/settings.xml')
    .root;
}
const style = resolveRunStyle([
  { localName: 'rFonts', attributes: { ascii: 'Liberation Serif', hAnsi: 'Liberation Serif' } },
  { localName: 'sz', attributes: { val: '28' } },
  { localName: 'kern', attributes: { val: '2' } },
]);

describe('document punctuation kerning policy', () => {
  test('accepts resolved settings while keeping absent and false policies unchanged', () => {
    const pieces: FieldAwarePiece[] = [{ text: 'AV.', start: 0, end: 3, props: [], style }];
    expect(withPunctuationKerningPolicy(pieces, {})).toBe(pieces);
    expect(withPunctuationKerningPolicy(pieces, { noPunctuationKerning: false })).toBe(pieces);
    expect(withPunctuationKerningPolicy(pieces, false)).toBe(pieces);
    const marked = withPunctuationKerningPolicy(pieces, { noPunctuationKerning: true });
    expect(marked[0]?.style.shaping?.noPunctuationKerning).toBe(true);
    expect(marked[0]?.props).toBe(pieces[0]?.props);
    expect(pieces[0]?.style).toBe(style);
  });

  test('reads direct Word settings and distinguishes false, absent, and malformed values', () => {
    for (const value of ['', ' w:val="1"', ' w:val="true"', ' w:val="on"'])
      expect(
        cjkTypographyFromSettings(settings(`<w:noPunctuationKerning${value}/>`))
          .noPunctuationKerning
      ).toBe(true);
    for (const value of ['0', 'false', 'off', 'invalid'])
      expect(
        cjkTypographyFromSettings(settings(`<w:noPunctuationKerning w:val="${value}"/>`))
          .noPunctuationKerning
      ).toBeUndefined();
    expect(cjkTypographyFromSettings(settings()).noPunctuationKerning).toBeUndefined();
    expect(
      cjkTypographyFromSettings(settings('<x:noPunctuationKerning xmlns:x="urn:test"/>'))
        .noPunctuationKerning
    ).toBeUndefined();
    expect(
      cjkTypographyFromSettings(settings('<w:compat><w:noPunctuationKerning/></w:compat>'))
        .noPunctuationKerning
    ).toBeUndefined();
    expect(
      cjkTypographyFromSettings(settings('<w:noPunctuationKerning/>', 'urn:test'))
        .noPunctuationKerning
    ).toBeUndefined();
  });

  test('changes the cascade key only for an effective direct setting', () => {
    const original = buildStyleCascadeTable(null, undefined, settings());
    const disabled = buildStyleCascadeTable(null, undefined, settings('<w:noPunctuationKerning/>'));
    const explicitFalse = buildStyleCascadeTable(
      null,
      undefined,
      settings('<w:noPunctuationKerning w:val="false"/>')
    );
    expect(disabled.cacheToken).not.toBe(original.cacheToken);
    expect(explicitFalse.cacheToken).toBe(original.cacheToken);
  });

  test('uses the last direct setting without treating nested or foreign settings as overrides', () => {
    expect(
      cjkTypographyFromSettings(
        settings('<w:noPunctuationKerning/><w:noPunctuationKerning w:val="0"/>')
      ).noPunctuationKerning
    ).toBeUndefined();
    expect(
      cjkTypographyFromSettings(
        settings('<w:noPunctuationKerning w:val="0"/><w:noPunctuationKerning/>')
      ).noPunctuationKerning
    ).toBe(true);
    expect(
      cjkTypographyFromSettings(
        settings(
          '<w:noPunctuationKerning/><w:compat><w:noPunctuationKerning w:val="0"/></w:compat>'
        )
      ).noPunctuationKerning
    ).toBe(true);
    expect(
      cjkTypographyFromSettings(
        settings('<w:noPunctuationKerning/><x:noPunctuationKerning xmlns:x="urn:test" w:val="0"/>')
      ).noPunctuationKerning
    ).toBe(true);
  });

  test('keeps atomic field ranges, properties, and authored style values unchanged', () => {
    const atom = { formField: false };
    const props: FieldAwarePiece['props'] = [];
    const piece: FieldAwarePiece = {
      text: 'AV.',
      start: 9,
      end: 10,
      projected: true,
      fieldAtom: atom,
      props,
      style,
    };
    const marked = withPunctuationKerningPolicy([piece], true)[0]!;
    expect(marked.text).toBe('AV.');
    expect([marked.start, marked.end]).toEqual([9, 10]);
    expect(marked.fieldAtom).toBe(atom);
    expect(marked.props).toBe(props);
    expect(marked.style.characterSpacingPt).toBe(style.characterSpacingPt);
    expect(marked.style.fontFamily).toBe(style.fontFamily);
    expect(style.shaping?.noPunctuationKerning).toBeUndefined();
    expect(marked.style.shaping?.noPunctuationKerning).toBe(true);
    expect(runStylesEqual(marked.style, style)).toBe(false);
    expect(withPunctuationKerningPolicy([piece], false)[0]).toBe(piece);
  });

  test('does not split surrogate pairs or detach punctuation combining marks', () => {
    expect(punctuationKerningRanges('😀AV.\u0301', 3)).toEqual([{ start: 7, end: 9 }]);
    expect(punctuationKerningSegments('😀AV.\u0301')).toEqual([
      { text: '😀AV', punctuation: false },
      { text: '.\u0301', punctuation: true },
    ]);
  });

  test('keeps disabled and below-threshold source kerning disabled', () => {
    for (const original of [
      { ...style, kerningMinPt: 0 },
      { ...style, kerningMinPt: 18 },
    ]) {
      const piece: FieldAwarePiece = { text: 'AV.', start: 0, end: 3, props: [], style: original };
      expect(withPunctuationKerningPolicy([piece], true)[0]).toBe(piece);
    }
  });
});

const bytes = new Uint8Array(
  readFileSync(new URL('../../../../fonts/assets/LiberationSerif-Regular.ttf', import.meta.url))
);
const shaping = await createLayoutShaping({
  epoch: 1,
  maxFontBytes: 1_000_000,
  sources: [
    {
      request: { family: 'Liberation Serif', weight: 400, style: 'normal' },
      id: 'liberation-serif',
      bytes,
      hash: sha256FontBytes(bytes),
      faceIndex: 0,
    },
  ],
  defaultFont: { family: 'Liberation Serif', sizeHalfPoints: 28 },
});
const face = shaping.fonts.resolve({ family: 'Liberation Serif', weight: 400, style: 'normal' });
if (face instanceof Error) throw face;
const measurer = createShapedMeasurer({
  shaper: shaping.shaper,
  environment: shaping.environment,
  resolveFont: () => face,
  fallback: createFixedMeasurer(),
});
const marked = withPunctuationKerningPolicy(
  [{ text: 'AV.', start: 0, end: 3, props: [], style }],
  true
)[0]!.style;
const advance = (text: string, runStyle = style) => measurer.measure(text, runStyle);

describe('real font punctuation kerning', () => {
  test('disables punctuation pairs while retaining the AV pair and both cache directions', () => {
    const natural = advance('AV.');
    const expected = advance('AV') + advance('.');
    expect(expected).toBeGreaterThan(natural);
    expect(advance('AV.', marked)).toBeCloseTo(expected, 3);
    expect(advance('AV')).toBeLessThan(advance('A') + advance('V'));
    expect(advance('AV.', marked)).toBeLessThan(advance('A') + advance('V') + advance('.'));
    expect(advance('AV.')).toBe(natural);
    expect(advance('AV.', marked)).toBeCloseTo(expected, 3);
    for (const text of ['T.', 'T,'])
      expect(advance(text, marked)).toBeCloseTo(advance(text[0]!) + advance(text[1]!), 3);
  });

  test('uses source settings in native layout and preserves save/reopen canonical source', () => {
    const part = parse(
      `<w:document xmlns:w="${WML}"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Liberation Serif"/><w:sz w:val="28"/><w:kern w:val="2"/></w:rPr><w:t>AV.</w:t></w:r></w:p></w:body></w:document>`,
      '/word/document.xml'
    );
    const before = serializeOoxmlPart(part);
    const cascade = buildStyleCascadeTable(null, undefined, settings('<w:noPunctuationKerning/>'));
    const model = layoutSemanticDocument(part, 0, { measurer, styleCascade: cascade });
    const spans = model.pages
      .flatMap((page) => page.fragments)
      .flatMap((fragment) => ('lines' in fragment ? fragment.lines : []))
      .flatMap((line) => line.spans);
    expect(spans.map((span) => span.text).join('')).toBe('AV.');
    expect(spans.reduce((sum, span) => sum + span.box.width, 0)).toBeCloseTo(
      advance('AV') + advance('.'),
      3
    );
    expect(spans.every((span) => span.style.characterSpacingPt === 0)).toBe(true);
    const saved = serializeOoxmlPart(part);
    expect(saved).toBe(before);
    expect(saved).not.toContain('w:spacing');
    expect(serializeOoxmlPart(parse(saved, '/word/document.xml'))).toBe(before);
  });

  test('invalidates native layout when one source switches true, false, and true settings', () => {
    const part = parse(
      `<w:document xmlns:w="${WML}"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Liberation Serif"/><w:sz w:val="28"/><w:kern w:val="2"/></w:rPr><w:t>AV.</w:t></w:r></w:p></w:body></w:document>`,
      '/word/document.xml'
    );
    const before = serializeOoxmlPart(part);
    for (const disabled of [true, false, true, false]) {
      const model = layoutSemanticDocument(part, 0, {
        measurer,
        styleCascade: buildStyleCascadeTable(
          null,
          undefined,
          settings(disabled ? '<w:noPunctuationKerning/>' : '<w:noPunctuationKerning w:val="0"/>')
        ),
      });
      const spans = model.pages
        .flatMap((page) => page.fragments)
        .flatMap((fragment) => ('lines' in fragment ? fragment.lines : []))
        .flatMap((line) => line.spans);
      expect(spans.reduce((sum, span) => sum + span.box.width, 0)).toBeCloseTo(
        disabled ? advance('AV') + advance('.') : advance('AV.'),
        3
      );
      expect(spans.every((span) => span.style.shaping?.noPunctuationKerning === true)).toBe(
        disabled
      );
    }
    expect(serializeOoxmlPart(part)).toBe(before);
  });

  test('retains the existing fallback when an admitted font cannot resolve', () => {
    const fallback = createFixedMeasurer();
    const missing = createShapedMeasurer({
      shaper: shaping.shaper,
      environment: shaping.environment,
      resolveFont: () => null,
      fallback,
    });
    expect(missing.measure('AV.', marked)).toBe(fallback.measure('AV.', marked));
    expect(missing.measure('AV.', style)).toBe(fallback.measure('AV.', style));
    expect(marked.fontFamily).toBe(style.fontFamily);
  });

  test('retains UTF-16 clusters in the native shaper with leading context', () => {
    const contextual = {
      ...marked,
      shaping: { ...marked.shaping!, context: { before: '😀', after: '' } },
    };
    const run = shapeLayoutStyleRun(shaping.shaper, shaping.environment, face, contextual, 'AV.');
    expect(run.glyphs.map((glyph) => glyph.cluster)).toEqual([0, 1, 2]);
    expect(run.glyphs.reduce((sum, glyph) => sum + glyph.advanceX, 0) / 1000).toBeCloseTo(
      advance('AV') + advance('.'),
      3
    );
  });

  test('preserves atomic cached field source positions through native layout and reopen', () => {
    const part = parse(
      `<w:document xmlns:w="${WML}"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Liberation Serif"/><w:sz w:val="28"/><w:kern w:val="2"/></w:rPr><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> MERGEFIELD PublicName </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>AV.</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>`,
      '/word/document.xml'
    );
    const before = serializeOoxmlPart(part);
    const model = layoutSemanticDocument(part, 0, {
      measurer,
      styleCascade: buildStyleCascadeTable(null, undefined, settings('<w:noPunctuationKerning/>')),
    });
    const spans = model.pages
      .flatMap((page) => page.fragments)
      .flatMap((fragment) => ('lines' in fragment ? fragment.lines : []))
      .flatMap((line) => line.spans);
    expect(spans.map((span) => span.text).join('')).toBe('AV.');
    expect(
      spans.every(
        (span) => span.projected && span.fieldAtom && span.range.end - span.range.start === 1
      )
    ).toBe(true);
    expect(spans.reduce((sum, span) => sum + span.box.width, 0)).toBeCloseTo(
      advance('AV') + advance('.'),
      3
    );
    expect(serializeOoxmlPart(part)).toBe(before);
    expect(serializeOoxmlPart(parse(before, '/word/document.xml'))).toBe(before);
  });

  test('uses real font metrics in the Canvas port and separates permission cache entries', () => {
    const context: CanvasTextContext = {
      font: '',
      fontKerning: 'normal',
      measureText(text) {
        return {
          width: advance(text, this.fontKerning === 'none' ? { ...style, kerningMinPt: 0 } : style),
        };
      },
    };
    const canvas = tryCreateCanvasMeasurer({ context, scale: 1 })!;
    const natural = canvas.measure('AV.', style);
    expect(canvas.measure('AV.', marked)).toBeCloseTo(advance('AV') + advance('.'), 3);
    expect(canvas.measure('AV.', style)).toBe(natural);
    expect(canvas.measure('AV.', marked)).toBeGreaterThan(natural);
  });
});

afterAll(() => disposeLayoutShaping(shaping));
