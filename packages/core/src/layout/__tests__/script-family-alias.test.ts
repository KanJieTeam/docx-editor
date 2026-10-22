import { expect, test } from 'bun:test';
import { DEFAULT_RUN_STYLE, glyphSizeFactorOf } from '../run-style.ts';
import { createScriptFamilyAliasQualifier } from '../script-family-alias.ts';
import { createShapedMeasurer } from '../shaped-measurer.ts';
import { createFixedMeasurer } from '../fixed-measurer.ts';
import { projectScriptGlyphSizing } from '../paragraph-script-glyph-sizing.ts';
import {
  createFontResourceSnapshot,
  boundedStructuralFontValidator,
  FontResolutionError,
  sha256FontBytes,
  type ResolvedFont,
} from '../font-resource.ts';
import {
  markFacePreservingShaper,
  fixedPoint,
  type ShapeInput,
  type ShapedRun,
} from '../shaped-run.ts';
import { HARFBUZZ_SHAPING_LIBRARY } from '../harfbuzz-shaper.ts';
import {
  syntheticScriptFont,
  syntheticScriptCollection,
} from './fixtures/synthetic-script-font.ts';

const alias = 'Example Alias',
  family = 'Example Family';
const style = {
  ...DEFAULT_RUN_STYLE,
  fontFamily: alias,
  fontSizePt: 14,
  verticalAlign: 'subscript' as const,
};
interface NameRecord {
  value: string;
  language?: number;
  nameId?: number;
  platform?: number;
  encoding?: number;
}
function namedFont(
  records: readonly NameRecord[] = [{ value: family }, { value: alias, language: 2052 }],
  subscriptY = 128
) {
  const base = syntheticScriptFont({ em: 256, subscriptY, superscriptY: 160 });
  const strings = records.map((record) => {
    const bytes = new Uint8Array(record.value.length * 2);
    const view = new DataView(bytes.buffer);
    for (let i = 0; i < record.value.length; i++) view.setUint16(i * 2, record.value.charCodeAt(i));
    return bytes;
  });
  const storage = 6 + 12 * records.length,
    nameLength = storage + strings.reduce((sum, s) => sum + s.length, 0),
    nameOffset = base.length + 16;
  const bytes = new Uint8Array(nameOffset + nameLength),
    view = new DataView(bytes.buffer),
    original = new DataView(base.buffer);
  bytes.set(base.subarray(0, 12));
  view.setUint16(4, 3);
  for (let index = 0; index < 2; index++) {
    bytes.set(base.subarray(12 + index * 16, 28 + index * 16), 12 + index * 16);
    view.setUint32(20 + index * 16, original.getUint32(20 + index * 16) + 16);
  }
  bytes.set(base.subarray(44), 60);
  view.setUint32(44, 0x6e616d65);
  view.setUint32(52, nameOffset);
  view.setUint32(56, nameLength);
  view.setUint16(nameOffset, 0);
  view.setUint16(nameOffset + 2, records.length);
  view.setUint16(nameOffset + 4, storage);
  let relative = 0;
  records.forEach((record, index) => {
    const start = nameOffset + 6 + index * 12;
    view.setUint16(start, record.platform ?? 3);
    view.setUint16(start + 2, record.encoding ?? 1);
    view.setUint16(start + 4, record.language ?? 1033);
    view.setUint16(start + 6, record.nameId ?? 1);
    view.setUint16(start + 8, strings[index]!.length);
    view.setUint16(start + 10, relative);
    bytes.set(strings[index]!, nameOffset + storage + relative);
    relative += strings[index]!.length;
  });
  return { bytes, view, nameOffset, nameLength };
}
function admitted(
  bytes: Uint8Array = namedFont().bytes,
  options: {
    from?: string;
    to?: string;
    fromWeight?: number;
    toWeight?: number;
    fromStyle?: 'normal' | 'italic';
    toStyle?: 'normal' | 'italic';
    lineMetrics?: { heightEm: number; baselineEm: number };
    faceIndex?: number;
  } = {}
) {
  const target = {
    family: options.to ?? family,
    weight: options.toWeight ?? 400,
    style: options.toStyle ?? ('normal' as const),
  };
  const requested = {
    family: options.from ?? alias,
    weight: options.fromWeight ?? 400,
    style: options.fromStyle ?? ('normal' as const),
  };
  const fonts = createFontResourceSnapshot({
    epoch: 1,
    maxFontBytes: 65536,
    resources: [
      {
        request: target,
        id: 'synthetic',
        bytes,
        hash: sha256FontBytes(bytes),
        faceIndex: options.faceIndex ?? 0,
      },
    ],
    substitutions: [
      {
        from: requested,
        to: target,
        ...(options.lineMetrics ? { lineMetrics: options.lineMetrics } : {}),
      },
    ],
    validateFont: boundedStructuralFontValidator,
  });
  const font = fonts.resolve(requested);
  if (font instanceof FontResolutionError) throw font;
  return { font, fonts };
}
function shape(input: ShapeInput): ShapedRun {
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
function measure(font: ResolvedFont) {
  return createShapedMeasurer({
    shaper: markFacePreservingShaper({ shape }),
    resolveFont: () => font,
    fallback: createFixedMeasurer(14, 14),
    shapingLibrary: HARFBUZZ_SHAPING_LIBRARY,
    unicodeDataVersion: '16.0.0',
    scriptGlyphSizing: 'font-metrics',
  });
}

test('an admitted face family alias receives its script factor across measurement and caret consumers', () => {
  const { font } = admitted(),
    m = measure(font);
  const piece = { text: '12', start: 0, end: 2, props: [], style };
  const projected = projectScriptGlyphSizing([piece], 'p', m)[0]!;
  expect(font.substitution).not.toBeNull();
  expect(projected.style).not.toBe(style);
  expect(glyphSizeFactorOf(projected.style)).toBe(0.5);
  expect(projected.style.fontFamily).toBe(alias);
  expect(projected.style.fontSizePt).toBe(14);
  expect(projected.style.layoutGlyphSizing?.fontRequest).toBe(
    JSON.stringify([family, 400, 'normal'])
  );
  expect(m.measure('12', projected.style)).toBe(14);
  expect(m.caretAdvances!('12', projected.style)).toEqual([0, 7, 14]);
  expect(piece.style.layoutGlyphSizing).toBeUndefined();
});
test('a directly registered face keeps its existing qualification without a family-name table', () => {
  const bytes = syntheticScriptFont({ em: 256, subscriptY: 128, superscriptY: 160 });
  const { fonts } = admitted(bytes);
  const direct = fonts.resolve({ family, weight: 400, style: 'normal' });
  if (direct instanceof FontResolutionError) throw direct;
  expect(direct.substitution).toBeNull();
  expect(
    measure(direct).scriptGlyphSizing!('1', { ...style, fontFamily: family }, 'p')?.factor
  ).toBe(0.5);
});
test('a valid selected collection face supplies its own family aliases and factor', () => {
  const bytes = syntheticScriptCollection([
    namedFont([{ value: 'Unrelated' }]).bytes,
    namedFont(undefined, 96).bytes,
  ]);
  const { font } = admitted(bytes, { faceIndex: 1 });
  expect(createScriptFamilyAliasQualifier()(font, style)).toBe(true);
  expect(measure(font).scriptGlyphSizing!('1', style, 'p')?.factor).toBe(0.375);
  const wrong = admitted(bytes, { faceIndex: 0 }).font;
  expect(createScriptFamilyAliasQualifier()(wrong, style)).toBe(false);
});
test('a translated family does not depend on a hardcoded alias list', () => {
  const { font } = admitted(
    namedFont([{ value: family }, { value: '示例别名', language: 2052 }]).bytes,
    { from: '示例别名' }
  );
  expect(createScriptFamilyAliasQualifier()(font, { ...style, fontFamily: '示例别名' })).toBe(true);
});
test('an alias cache cannot qualify another request style or another admitted font', () => {
  const qualify = createScriptFamilyAliasQualifier(),
    a = admitted().font,
    b = admitted(namedFont([{ value: family }]).bytes).font;
  expect(qualify(a, style)).toBe(true);
  expect(qualify(a, { ...style, fontFamily: 'Unknown' })).toBe(false);
  expect(qualify(a, { ...style, bold: true })).toBe(false);
  expect(qualify(b, style)).toBe(false);
  expect(qualify(a, style)).toBe(true);
});
for (const [title, options, currentStyle] of [
  ['true fallback', { from: 'Unknown' }, { ...style, fontFamily: 'Unknown' }],
  ['weight replacement', { fromWeight: 700 }, { ...style, bold: true }],
  ['italic replacement', { fromStyle: 'italic' as const }, { ...style, italic: true }],
  ['line metric override', { lineMetrics: { heightEm: 1.2, baselineEm: 0.9 } }, style],
] as const)
  test(title, () => {
    const { font } = admitted(undefined, options);
    expect(createScriptFamilyAliasQualifier()(font, currentStyle)).toBe(false);
    expect(measure(font).scriptGlyphSizing!('1', currentStyle, 'p')).toBeUndefined();
  });
for (const [title, change] of [
  ['unknown name format', (f: ReturnType<typeof namedFont>) => f.view.setUint16(f.nameOffset, 1)],
  [
    'record count ceiling',
    (f: ReturnType<typeof namedFont>) => f.view.setUint16(f.nameOffset + 2, 257),
  ],
  [
    'storage overlaps records',
    (f: ReturnType<typeof namedFont>) => f.view.setUint16(f.nameOffset + 4, 6),
  ],
  [
    'out of table string',
    (f: ReturnType<typeof namedFont>) => f.view.setUint16(f.nameOffset + 6 + 10, 65535),
  ],
  [
    'odd UTF16 byte length',
    (f: ReturnType<typeof namedFont>) => f.view.setUint16(f.nameOffset + 6 + 8, 3),
  ],
  [
    'truncated family string',
    (f: ReturnType<typeof namedFont>) => f.view.setUint16(f.nameOffset + 6 + 8, 65534),
  ],
  [
    'unpaired high surrogate',
    (f: ReturnType<typeof namedFont>) =>
      f.view.setUint16(f.nameOffset + f.view.getUint16(f.nameOffset + 4), 0xd800),
  ],
  [
    'unpaired low surrogate',
    (f: ReturnType<typeof namedFont>) =>
      f.view.setUint16(f.nameOffset + f.view.getUint16(f.nameOffset + 4), 0xdc00),
  ],
  [
    'embedded control',
    (f: ReturnType<typeof namedFont>) =>
      f.view.setUint16(f.nameOffset + f.view.getUint16(f.nameOffset + 4), 0),
  ],
] as const)
  test(title, () => {
    const fixture = namedFont();
    change(fixture);
    expect(createScriptFamilyAliasQualifier()(admitted(fixture.bytes).font, style)).toBe(false);
  });
test('conflicting records and unsupported encodings cannot prove an alias', () => {
  for (const records of [
    [{ value: family }, { value: alias, language: 1033 }],
    [{ value: family }, { value: alias, language: 2052, platform: 1, encoding: 0 }],
    [{ value: family }, { value: alias, language: 2052, encoding: 0 }],
  ])
    expect(createScriptFamilyAliasQualifier()(admitted(namedFont(records).bytes).font, style)).toBe(
      false
    );
});
test('name strings cannot escape their declared table even when adjacent bytes exist', () => {
  const f = namedFont();
  f.view.setUint32(56, 6 + 24);
  expect(createScriptFamilyAliasQualifier()(admitted(f.bytes).font, style)).toBe(false);
});

test('a copied font metadata object cannot leak an alias byte-capability error', () => {
  const { font } = admitted();
  const untrusted = Object.create(font) as ResolvedFont;
  const m = measure(untrusted);
  expect(() => m.scriptGlyphSizing!('1', style, 'p')).not.toThrow();
  expect(m.scriptGlyphSizing!('1', style, 'p')).toBeUndefined();
});
