import { expect, test } from 'bun:test';
import type { StyleSpanRecord } from '@docx-editor.dev/core/layout';
import { escapeText, MarkdownInlineWriter } from '../src/markdown-inline.ts';
import { readFileSync } from 'node:fs';
import { readOoxmlPart } from '@docx-editor.dev/core/store';
import {
  acquireSharedExportShaping,
  type ExportSemanticLayout,
} from '@docx-editor.dev/core/export';
import {
  layoutSemanticDocument,
  prepareLayoutFontConfiguration,
  sha256FontBytes,
  linesOf,
} from '@docx-editor.dev/core/layout';
import { exportMarkdownLayout } from '../src/markdown.ts';

function styledToken(text: string, start: number, style: Partial<StyleSpanRecord['style']>) {
  return {
    paragraphId: 'p',
    sourceText: text,
    span: {
      range: { paragraphId: 'p', start, end: start + text.length },
      text,
      style: style as StyleSpanRecord['style'],
    } as StyleSpanRecord,
  };
}

function token(text: string, start: number, bold = false) {
  return styledToken(text, start, bold ? { bold: true } : {});
}

test('keeps ordinary in-word hyphens readable without opening Markdown syntax', () => {
  expect(escapeText('docx-editor.dev')).toBe('docx-editor\\.dev');
  expect(escapeText('multi-level and source-to-export')).toBe('multi-level and source-to-export');
  expect(escapeText('日本-語')).toBe('日本-語');
  expect(escapeText('𐐀-𐐁')).toBe('𐐀-𐐁');
});

test('keeps in-word hyphens readable across DOCX run and style boundaries', () => {
  const sameStyle = new MarkdownInlineWriter({
    tableCell: false,
    displayMode: 'all-markup',
    sourceScope: 'body',
  });
  sameStyle.writeText(token('source', 0));
  sameStyle.writeText(token('-', 6));
  sameStyle.writeText(token('to', 7));
  expect(sameStyle.finish()).toBe('source-to');

  const changedStyle = new MarkdownInlineWriter({
    tableCell: false,
    displayMode: 'all-markup',
    sourceScope: 'body',
  });
  changedStyle.writeText(token('source', 0));
  changedStyle.writeText(token('-', 6));
  changedStyle.writeText(token('to', 7, true));
  expect(changedStyle.finish()).toBe('source-**to**');

  const astral = new MarkdownInlineWriter({
    tableCell: false,
    displayMode: 'all-markup',
    sourceScope: 'body',
  });
  astral.writeText(token('𐐀', 0));
  astral.writeText(token('-', 2));
  astral.writeText(token('𐐁', 3));
  expect(astral.finish()).toBe('𐐀-𐐁');

  const structuralBoundary = new MarkdownInlineWriter({
    tableCell: false,
    displayMode: 'all-markup',
    sourceScope: 'body',
  });
  structuralBoundary.writeText(token('source', 0));
  structuralBoundary.writeBoundary('<br>');
  structuralBoundary.writeText(token('-to', 6));
  expect(structuralBoundary.finish()).toBe('source<br>\\-to');
});

test('keeps one style delimiter around Word runs split at spaces', () => {
  const writer = new MarkdownInlineWriter({
    tableCell: false,
    displayMode: 'all-markup',
    sourceScope: 'body',
  });
  writer.writeText(token('Table ', 0, true));
  writer.writeText(token('of ', 6, true));
  writer.writeText(token('Contents', 9, true));

  expect(writer.finish()).toBe('**Table of Contents**');
});

test('does not absorb an independently unstyled space into matching styled runs', () => {
  for (const [style, delimiter] of [
    [{ bold: true }, '**'],
    [{ italic: true }, '*'],
    [{ strike: true }, '~~'],
  ] as const) {
    const writer = new MarkdownInlineWriter({
      tableCell: false,
      displayMode: 'all-markup',
      sourceScope: 'body',
    });
    writer.writeText(styledToken('Old', 0, style));
    writer.writeText(styledToken(' ', 3, {}));
    writer.writeText(styledToken('New', 4, style));

    expect(writer.finish()).toBe(`${delimiter}Old${delimiter} ${delimiter}New${delimiter}`);
  }
});

test('does not bridge a mixed-provenance sequence of whitespace chunks', () => {
  for (const [style, delimiter] of [
    [{ bold: true }, '**'],
    [{ strike: true }, '~~'],
  ] as const) {
    const writer = new MarkdownInlineWriter({
      tableCell: false,
      displayMode: 'all-markup',
      sourceScope: 'body',
    });
    writer.writeText(styledToken('Old ', 0, style));
    writer.writeText(styledToken(' ', 4, {}));
    writer.writeText(styledToken('New', 5, style));

    expect(writer.finish()).toBe(`${delimiter}Old${delimiter}  ${delimiter}New${delimiter}`);
  }
});

test('still escapes hyphens that could be interpreted as Markdown punctuation', () => {
  expect(escapeText('- list item')).toBe('\\- list item');
  expect(escapeText('a - b')).toBe('a \\- b');
  expect(escapeText('---')).toBe('\\-\\-\\-');
});

test('real derived script carriers do not change Markdown or leak font proofs', async () => {
  const bytes = new Uint8Array(
    readFileSync(
      new URL('../../core/src/layout/__tests__/fixtures/fonts/DejaVuSans.ttf', import.meta.url)
    )
  );
  const prepared = prepareLayoutFontConfiguration({
    epoch: 1,
    maxFontBytes: 2_000_000,
    sources: [
      {
        request: { family: 'DejaVu Sans', weight: 400, style: 'normal' },
        id: 'markdown-script',
        bytes,
        hash: sha256FontBytes(bytes),
        faceIndex: 0,
      },
    ],
    defaultFont: { family: 'DejaVu Sans', sizeHalfPoints: 28 },
  });
  const read = readOoxmlPart(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="28"/><w:vertAlign w:val="subscript"/></w:rPr><w:t>123</w:t></w:r></w:p></w:body></w:document>',
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  const render = async (mode: 'legacy' | 'font-metrics') => {
    const view = await acquireSharedExportShaping(prepared, undefined, { scriptGlyphSizing: mode });
    const layout = layoutSemanticDocument(read.part, 1, {
      measurer: view.createMeasurer(),
      producer: view.producer,
    });
    return { layout, result: exportMarkdownLayout(layout as ExportSemanticLayout) };
  };
  const nominal = await render('legacy');
  const derived = await render('font-metrics');
  const carrier = linesOf(derived.layout)[0]!.spans[0]!.style.layoutGlyphSizing;
  expect(carrier).toBeDefined();
  expect(linesOf(derived.layout)[0]!.spans[0]!.style.fontSizePt).toBe(14);
  expect(derived.result).toEqual(nominal.result);
  expect(derived.result.markdown).toBe('123');
  expect(derived.result.warnings).toEqual([]);
  const exported = JSON.stringify(derived.result);
  expect(exported).not.toContain(carrier!.fontHash);
  expect(exported).not.toContain(carrier!.proof);
  expect(exported).not.toContain('layoutGlyphSizing');
});
