import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
import { expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart } from '../../store/package/ooxml-tree.ts';
import { createFixedMeasurer } from '../../layout/fixed-measurer.ts';
import { layoutSemanticDocument } from '../../layout/semantic-layout.ts';
import { buildStyleCascadeTable } from '../../layout/style-cascade.ts';
import { paintSemanticLayout } from '../../output/semantic-paint.ts';
import {
  applySelectionToDom,
  positionFromDomPoint,
  semanticSelectionFromDom,
} from '../dom-selection.ts';

const WML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function painted(text: string) {
  const part = readOoxmlPart(
    `<w:document xmlns:w="${WML}"><w:body><w:p><w:r><w:rPr><w:sz w:val="28"/><w:kern w:val="2"/><w:u/><w:strike/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  const settings = readOoxmlPart(
    `<w:settings xmlns:w="${WML}"><w:noPunctuationKerning/></w:settings>`,
    { name: '/word/settings.xml', contentType: 'application/xml' }
  );
  if (!part.ok || !settings.ok) throw new Error('Invalid source');
  const before = serializeOoxmlPart(part.part);
  const root = document.createElement('div');
  const layout = layoutSemanticDocument(part.part, 0, {
    measurer: createFixedMeasurer(),
    styleCascade: buildStyleCascadeTable(null, undefined, settings.part.root),
  });
  paintSemanticLayout(root, layout, { scale: 1 });
  document.body.append(root);
  const runs = Array.from(root.querySelectorAll<HTMLElement>('.layout-run-text'));
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Node[] = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  return { part: part.part, before, root, runs, nodes };
}

test('selection reads and writes all punctuation segments under nested decoration layers', () => {
  const fixture = painted('AV.f');
  try {
    const paragraphId = fixture.runs[0]!.dataset.paragraphId!;
    const punctuation = fixture.nodes.find((node) => node.textContent === '.')!;
    expect(positionFromDomPoint(punctuation, 0, fixture.root)).toEqual({ paragraphId, offset: 2 });
    expect(positionFromDomPoint(punctuation, 1, fixture.root)).toEqual({ paragraphId, offset: 3 });
    expect(positionFromDomPoint(punctuation.parentNode!, 0, fixture.root)).toEqual({
      paragraphId,
      offset: 2,
    });
    expect(positionFromDomPoint(punctuation.parentNode!, 1, fixture.root)).toEqual({
      paragraphId,
      offset: 3,
    });
    const selection = document.getSelection()!;
    for (let offset = 0; offset <= 4; offset += 1) {
      const position = { paragraphId, offset };
      expect(
        applySelectionToDom(fixture.root, { anchor: position, head: position }, selection)
      ).toBe(true);
      expect(semanticSelectionFromDom(fixture.root, selection)).toEqual({
        anchor: position,
        head: position,
      });
    }
    const range = { anchor: { paragraphId, offset: 1 }, head: { paragraphId, offset: 4 } };
    expect(applySelectionToDom(fixture.root, range, selection)).toBe(true);
    expect(selection.toString()).toBe('V.f');
    expect(semanticSelectionFromDom(fixture.root, selection)).toEqual(range);
    expect(serializeOoxmlPart(fixture.part)).toBe(fixture.before);
  } finally {
    document.getSelection()?.removeAllRanges();
    fixture.root.remove();
  }
});

test('segmented text retains UTF-16 positions for astral text and punctuation combining marks', () => {
  const fixture = painted('😀AV.\u0301f');
  try {
    const paragraphId = fixture.runs[0]!.dataset.paragraphId!;
    const punctuation = fixture.nodes.find((node) => node.textContent === '.\u0301')!;
    expect(positionFromDomPoint(punctuation, 0, fixture.root)).toEqual({ paragraphId, offset: 4 });
    expect(positionFromDomPoint(punctuation, 2, fixture.root)).toEqual({ paragraphId, offset: 6 });
    const selection = document.getSelection()!;
    for (const offset of [0, 2, 3, 4, 6, 7]) {
      const point = { paragraphId, offset };
      expect(applySelectionToDom(fixture.root, { anchor: point, head: point }, selection)).toBe(
        true
      );
      expect(semanticSelectionFromDom(fixture.root, selection)).toEqual({
        anchor: point,
        head: point,
      });
    }
    expect(serializeOoxmlPart(fixture.part)).toBe(fixture.before);
  } finally {
    document.getSelection()?.removeAllRanges();
    fixture.root.remove();
  }
});
