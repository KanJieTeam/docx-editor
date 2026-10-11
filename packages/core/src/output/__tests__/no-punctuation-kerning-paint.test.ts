import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
import { expect, test } from 'bun:test';
import { readOoxmlPart } from '../../store/package/ooxml-tree.ts';
import { createFixedMeasurer } from '../../layout/fixed-measurer.ts';
import { layoutSemanticDocument } from '../../layout/semantic-layout.ts';
import { buildStyleCascadeTable } from '../../layout/style-cascade.ts';
import { paintSemanticLayout } from '../semantic-paint.ts';

const WML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function paint(disabled: boolean): HTMLElement {
  const part = readOoxmlPart(
    `<w:document xmlns:w="${WML}"><w:body><w:p><w:r><w:rPr><w:sz w:val="28"/><w:kern w:val="2"/></w:rPr><w:t>AV.</w:t></w:r></w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  const settings = readOoxmlPart(
    `<w:settings xmlns:w="${WML}">${disabled ? '<w:noPunctuationKerning/>' : ''}</w:settings>`,
    { name: '/word/settings.xml', contentType: 'application/xml' }
  );
  if (!part.ok || !settings.ok) throw new Error('Invalid test source');
  const container = document.createElement('div');
  paintSemanticLayout(
    container,
    layoutSemanticDocument(part.part, 0, {
      measurer: createFixedMeasurer(),
      styleCascade: buildStyleCascadeTable(null, undefined, settings.part.root),
    }),
    { scale: 1 }
  );
  return container;
}

test('native text keeps Latin kerning while punctuation has its own disabled paint segment', () => {
  const run = paint(true).querySelector<HTMLElement>('.layout-run-text')!;
  expect(run.textContent).toBe('AV.');
  expect(
    Array.from(run.querySelectorAll('span')).map((span) => [
      span.textContent,
      span.style.fontKerning,
    ])
  ).toEqual([
    ['AV', 'normal'],
    ['.', 'none'],
  ]);
  expect(run.style.letterSpacing).toBe('');
  expect(run.dataset.start).toBe('0');
  const original = paint(false).querySelector<HTMLElement>('.layout-run-text')!;
  expect(original.textContent).toBe('AV.');
  expect(original.querySelectorAll('span').length).toBe(0);
  expect(original.style.fontKerning).toBe('normal');
});
