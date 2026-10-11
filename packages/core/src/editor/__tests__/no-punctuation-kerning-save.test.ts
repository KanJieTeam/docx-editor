import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
import { expect, test } from 'bun:test';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';
import { createFixedMeasurer } from '../../layout/fixed-measurer.ts';
import { serializeOoxmlPart, type OoxmlNode } from '../../store/index.ts';
import { linesOf } from '../../layout/semantic-records.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/package/2006/relationships';
const O = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
function fixture(spacing: number): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${R}"><Relationship Id="rId1" Type="${O}officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${R}"><Relationship Id="rId2" Type="${O}settings" Target="settings.xml"/><Relationship Id="rId3" Type="urn:public-fixture:opaque" Target="../opaque.bin"/></Relationships>`
    ),
    'word/settings.xml': strToU8(
      `<w:settings xmlns:w="${W}"><w:noPunctuationKerning/></w:settings>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:sz w:val="28"/><w:kern w:val="2"/><w:spacing w:val="${spacing}"/></w:rPr><w:t>AV.</w:t></w:r></w:p><w:p><w:r><w:rPr><w:kern w:val="0"/></w:rPr><w:t>target</w:t></w:r></w:p></w:body></w:document>`
    ),
    'opaque.bin': new Uint8Array([0, 7, 255, 0, 127]),
  });
}
function withSurface(bytes: Uint8Array, run: (surface: PaginatedSurface) => void): void {
  const container = document.createElement('div');
  document.body.append(container);
  const opened = mountPaginatedSurface(container, bytes, { measurer: createFixedMeasurer() });
  if (!opened.ok) throw new Error(opened.reason);
  try {
    run(opened.surface);
  } finally {
    opened.surface.destroy();
    container.remove();
  }
}
function propertyValue(node: OoxmlNode, localName: string): string | undefined {
  if (node.kind === 'textValue') return;
  if (node.localName === localName)
    return node.attributes.find((attribute) => attribute.localName === 'val')?.value;
  for (const child of node.children) {
    const value = propertyValue(child, localName);
    if (value !== undefined) return value;
  }
}
function paragraph(surface: PaginatedSurface, index: number): OoxmlNode {
  const id = surface.session.paragraphIds()[index]!;
  const body = surface.session
    .part()
    .root.children.find((node) => node.kind !== 'textValue' && node.localName === 'body')!;
  if (body.kind === 'textValue') throw new Error('Missing body');
  return body.children.find((node) => node.id === id)!;
}
function select(surface: PaginatedSurface, index: number, start: number, end: number): void {
  const paragraphId = surface.session.paragraphIds()[index]!;
  surface.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
}
for (const spacing of [0, -30])
  test(`real Format Painter, save, and cold reopen retain author tracking ${spacing}`, () => {
    const input = fixture(spacing);
    const originalOpaque = unzipSync(input)['opaque.bin']!;
    withSurface(input, (surface) => {
      const firstId = surface.session.paragraphIds()[0]!;
      const first = linesOf(surface.layout())
        .filter((line) => line.range.paragraphId === firstId)
        .flatMap((line) => line.spans);
      expect(first.every((span) => span.style.shaping?.noPunctuationKerning === true)).toBe(true);
      expect(first.every((span) => span.style.characterSpacingPt === spacing / 20)).toBe(true);
      const initial = serializeOoxmlPart(surface.session.part());
      withSurface(surface.session.save(), (cold) =>
        expect(serializeOoxmlPart(cold.session.part())).toBe(initial)
      );
      select(surface, 0, 2, 3);
      expect(surface.formatPainter.capture()).toBe(true);
      select(surface, 1, 0, 1);
      expect(surface.formatPainter.apply()).toBe('painted');
      expect(propertyValue(paragraph(surface, 0), 'spacing')).toBe(String(spacing));
      expect(propertyValue(paragraph(surface, 1), 'spacing')).toBe(String(spacing));
      expect(propertyValue(paragraph(surface, 1), 'kern')).toBe('2');
      const saved = surface.session.save();
      const parts = unzipSync(saved);
      expect(parts['opaque.bin']).toEqual(originalOpaque);
      const canonical = serializeOoxmlPart(surface.session.part());
      expect(canonical).not.toContain('noPunctuationKerning');
      withSurface(saved, (cold) => {
        expect(serializeOoxmlPart(cold.session.part())).toBe(canonical);
        expect(propertyValue(paragraph(cold, 1), 'spacing')).toBe(String(spacing));
        expect(propertyValue(paragraph(cold, 1), 'kern')).toBe('2');
        expect(unzipSync(cold.session.save())['opaque.bin']).toEqual(originalOpaque);
      });
    });
  });
