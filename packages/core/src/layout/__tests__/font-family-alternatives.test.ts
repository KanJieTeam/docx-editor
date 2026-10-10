import { fontFamilyStack } from '../font-family-stack.ts';
import { unzipSync, strFromU8 } from 'fflate';
import { openTreeSession } from '../../binding/tree-session.ts';
import { fontFamilyAlternativesDocx } from './fixtures/font-family-alternatives.ts';
import { expect, test } from 'bun:test';
import {
  fontFamilyAlternatives,
  validFontFamilyReference,
} from '../../store/package/font-family-reference.ts';
import {
  familyFromRFonts,
  eastAsiaFamilyFromRFonts,
  createRunDefaultsResolver,
} from '../../store/package/run-defaults.ts';
import { readOoxmlPart } from '../../store/package/ooxml-tree.ts';
import { collectDocumentFonts } from '../../binding/document-catalog.ts';
import { registerEmbeddedFontFaces } from '../../editor/embedded-font-faces.ts';
import { paintSemanticLayout } from '../../output/semantic-paint.ts';
import {
  boundedStructuralFontValidator,
  createFontResourceSnapshot,
  FontResolutionError,
  sha256FontBytes,
  DEFAULT_RUN_STYLE,
  tryCreateCanvasMeasurer,
  createFixedMeasurer,
  layoutSemanticDocument,
  type FontRequest,
  type CanvasTextContext,
} from '../index.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const reference = '宋体;SimSun';
const request = (family: string, weight = 400): FontRequest => ({
  family,
  weight,
  style: 'normal',
});
const source = (family: string, weight = 400) => {
  const bytes = new Uint8Array([0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
  return {
    request: request(family, weight),
    id: family,
    bytes,
    hash: sha256FontBytes(bytes),
    faceIndex: 0,
  };
};
const snapshot = (resources: ReturnType<typeof source>[]) =>
  createFontResourceSnapshot({
    resources,
    epoch: 1,
    maxFontBytes: 1024,
    validateFont: boundedStructuralFontValidator,
  });

test('alternatives keep source order and reject unsafe or excessive references', () => {
  expect(fontFamilyAlternatives(reference)).toEqual(['宋体', 'SimSun']);
  expect(fontFamilyAlternatives('黑体; SimHei')).toEqual(['黑体', 'SimHei']);
  expect(validFontFamilyReference(reference)).toBe(reference);
  for (const raw of [
    'Arial;',
    ';Arial',
    'Arial;;SimSun',
    'Arial; ',
    'Arial;bad()',
    'Arial;url(x)',
    'Arial;"x"',
    'Arial;../x',
    'Arial;C:\\x',
    'Arial;\tSimSun',
    'Arial;\nSimSun',
    'Arial,x',
    'x'.repeat(65),
    Array(9).fill('A').join(';'),
  ]) {
    expect(fontFamilyAlternatives(raw)).toEqual([]);
    expect(validFontFamilyReference(raw)).toBeNull();
  }
});

test('catalog discovery and inherited references keep independent candidates', () => {
  const xml =
    `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr>` +
    `<w:rFonts w:ascii="${reference}" w:eastAsia="黑体;SimHei"/>` +
    '</w:rPr><w:t>PUBLIC FONT TEST</w:t></w:r></w:p></w:body></w:document>';
  const parsed = readOoxmlPart(xml, { name: '/word/document.xml', contentType: 'app/xml' });
  if (!parsed.ok) throw new Error(parsed.reason);
  expect(collectDocumentFonts([parsed.part.root])).toEqual(['SimHei', 'SimSun', '宋体', '黑体']);
  const p = parsed.part.root.children[0];
  if (p.kind === 'textValue') throw new Error('body');
  const stack = [p];
  while (stack.length) {
    const node = stack.pop()!;
    if (node.localName === 'rFonts') {
      expect(familyFromRFonts(node, { major: null, minor: null })).toBe(reference);
      expect(eastAsiaFamilyFromRFonts(node, { major: null, minor: null })).toBe('黑体;SimHei');
    }
    for (const child of node.children) if (child.kind !== 'textValue') stack.push(child);
  }
});

test('resolution records the selected face without changing the authored request', () => {
  const fonts = snapshot([source('SimSun')]);
  const resolved = fonts.resolve(request(reference));
  if (resolved instanceof FontResolutionError) throw resolved;
  expect(resolved.request.family).toBe('SimSun');
  expect(resolved.substitution?.requested.family).toBe(reference);
  expect(resolved.substitution?.resolved.family).toBe('SimSun');
});

test('the first configured family controls variants and refusal', () => {
  const fonts = snapshot([source('宋体'), source('SimSun', 700)]);
  expect(fonts.resolve(request(reference, 700))).toBeInstanceOf(FontResolutionError);
  const first = fonts.resolve(request(reference));
  if (first instanceof FontResolutionError) throw first;
  expect(first.request.family).toBe('宋体');
  for (const availability of ['forbidden', 'available'] as const) {
    const blocked = source('宋体');
    const resources = [
      { ...blocked, availability, hash: availability === 'available' ? 'invalid' : blocked.hash },
      source('SimSun'),
    ];
    const result = snapshot(resources).resolve(request(reference));
    expect(result).toBeInstanceOf(FontResolutionError);
    expect((result as FontResolutionError).code).toBe(
      availability === 'forbidden' ? 'forbidden' : 'hashMismatch'
    );
  }
});

test('Canvas and semantic paint use the same safe stack and preserve emphasis', async () => {
  const registration = await registerEmbeddedFontFaces([source('SimSun')], {
    fontSet: { add() {}, delete() {} },
    createFontFace: () => ({ load: async () => undefined }),
  });
  expect(registration.alias(reference)).toBe(registration.alias('SimSun'));
  expect(registration.alias('Arial;bad()')).toBeUndefined();
  const context = { font: '', measureText: () => ({ width: 10 }) } as unknown as CanvasTextContext;
  const style = { ...DEFAULT_RUN_STYLE, fontFamily: reference, bold: true, italic: true };
  tryCreateCanvasMeasurer({ context, fontAlias: registration.alias })!.measure('PUBLIC', style);
  expect(context.font).toContain(`"${registration.alias(reference)}", "宋体", "SimSun"`);
  expect(context.font).toContain('italic normal bold');
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr>` +
      `<w:rFonts w:ascii="${reference}"/><w:b/><w:i/></w:rPr><w:t>PUBLIC</w:t>` +
      '</w:r></w:p></w:body></w:document>',
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!parsed.ok) throw new Error(parsed.reason);
  const layout = layoutSemanticDocument(parsed.part, 1, { measurer: createFixedMeasurer(6, 14) });
  const container = document.createElement('div');
  paintSemanticLayout(container, layout, { scale: 1, fontAlias: registration.alias });
  const span = container.querySelector<HTMLElement>('.docx-line span')!;
  expect(span.style.fontFamily.replaceAll('"', '')).toContain('宋体, SimSun');
  expect(span.style.fontWeight).toBe('bold');
  expect(span.style.fontStyle).toBe('italic');
  registration.dispose();
  expect(registration.alias(reference)).toBeUndefined();
});

test('save and reopen preserve authored alternatives and independent plain names', () => {
  const original = fontFamilyAlternativesDocx();
  const opened = openTreeSession(original);
  if (!opened.ok) throw new Error(opened.reason);
  const saved = opened.session.save();
  const reopened = openTreeSession(saved);
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(reopened.session.bodyText()).toBe(opened.session.bodyText());
  const xml = strFromU8(unzipSync(saved)['word/document.xml']!);
  expect(xml).toContain('w:ascii="宋体;SimSun"');
  expect(xml).toContain('w:hAnsi="宋体;SimSun"');
  expect(xml).toContain('w:eastAsia="黑体;SimHei"');
  expect(xml).toContain('w:ascii="SimSun"');
  expect(unzipSync(saved)['word/styles.xml']).toEqual(unzipSync(original)['word/styles.xml']);
  expect(reopened.session.documentFonts()).toEqual(['SimHei', 'SimSun', '宋体', '黑体']);
});

test('inherited defaults preserve the validated authored reference', () => {
  const stylesXml = strFromU8(unzipSync(fontFamilyAlternativesDocx())['word/styles.xml']!);
  const read = readOoxmlPart(stylesXml, { name: '/word/styles.xml', contentType: 'app/xml' });
  if (!read.ok) throw new Error(read.reason);
  const inherited = createRunDefaultsResolver(read.part.root, { major: null, minor: null })(null);
  expect(inherited.fontFamily).toBe('宋体;SimSun');
  expect(inherited.fontSizeHalfPoints).toBe(28);
});

test('every supplied alternative precedes platform names at both text sinks', () => {
  const aliases: Record<string, string> = { First: 'docx-first', Second: 'docx-second' };
  const alias = (family: string) => aliases[family];
  expect(fontFamilyStack('First;Second', alias)).toBe(
    '"docx-first", "docx-second", "First", "Second"'
  );
  expect(fontFamilyStack('First;Second', () => 'bad;name')).toBe('"First", "Second"');
  expect(fontFamilyStack('First;bad()', alias)).toBeNull();
  const context = { font: '', measureText: () => ({ width: 10 }) } as unknown as CanvasTextContext;
  tryCreateCanvasMeasurer({ context, fontAlias: alias })!.measure('PUBLIC', {
    ...DEFAULT_RUN_STYLE,
    fontFamily: 'First;Second',
  });
  expect(context.font).toContain('"docx-first", "docx-second", "First", "Second"');
});
