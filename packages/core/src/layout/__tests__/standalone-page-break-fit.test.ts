import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart } from '@docx-editor.dev/core/store';
import {
  createFixedMeasurer,
  createLayoutSession,
  layoutSemanticDocument,
} from '../semantic-layout.ts';
import { caretAt } from '../semantic-interaction.ts';
import type { ParagraphFragmentRecord } from '../semantic-records.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const spacing = '<w:spacing w:before="0" w:after="0" w:line="280" w:lineRule="exact"/>';
const paragraph = (content: string, extra = '') =>
  `<w:p><w:pPr>${spacing}<w:widowControl w:val="0"/>${extra}</w:pPr><w:r>${content}</w:r></w:p>`;
const pageBreak = paragraph('<w:br w:type="page"/>');
const title = paragraph('<w:t>TITLE</w:t>');
const section =
  '<w:sectPr><w:pgSz w:w="4000" w:h="4000"/>' +
  '<w:pgMar w:top="400" w:bottom="400" w:left="400" w:right="400"/></w:sectPr>';
const table = (height: number) =>
  '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar>' +
  '</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr><w:trPr><w:trHeight w:val="${height * 20}" w:hRule="exact"/></w:trPr>` +
  '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>' +
  paragraph('<w:t>TABLE</w:t>') +
  '</w:tc></w:tr></w:tbl>';
const measurer = createFixedMeasurer(5, 14);
function load(body: string) {
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body>${body}${section}</w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType: 'application/xml',
    }
  );
  if (!parsed.ok) throw new Error(parsed.reason);
  return parsed.part;
}
const textPages = (layout: ReturnType<typeof layoutSemanticDocument>) =>
  layout.pages.map((page) =>
    page.fragments
      .filter((fragment): fragment is ParagraphFragmentRecord => fragment.kind === 'paragraph')
      .flatMap((fragment) => fragment.lines.flatMap((line) => line.spans.map((span) => span.text)))
      .join('')
      .replace(/[\f\n]/g, '')
  );
const paragraphs = (layout: ReturnType<typeof layoutSemanticDocument>) =>
  layout.pages.flatMap((page, pageIndex) =>
    page.fragments
      .filter((fragment): fragment is ParagraphFragmentRecord => fragment.kind === 'paragraph')
      .map((fragment) => ({ fragment, pageIndex }))
  );

describe('plain standalone page breaks', () => {
  // WPS independently exports these public, text-only fixtures as two pages.
  for (const height of [150, 160]) {
    test(`keeps the break after a ${height}pt table on the table page`, () => {
      const part = load(table(height) + pageBreak + title);
      const source = serializeOoxmlPart(part);
      const layout = layoutSemanticDocument(part, 0, { measurer });
      expect(layout.pages).toHaveLength(2);
      expect(textPages(layout)).toEqual(['', 'TITLE']);
      const held = paragraphs(layout).find(({ fragment }) =>
        fragment.lines.some((line) => line.spans.some((span) => span.text === '\f'))
      )!;
      expect(held.pageIndex).toBe(0);
      expect(held.fragment.outOfFlow).toBe(true);
      const line = held.fragment.lines[0]!;
      expect(line.box.height).toBe(14);
      expect(line.range.end - line.range.start).toBe(1);
      expect(
        caretAt(layout, { paragraphId: held.fragment.paragraphId, offset: 0 })!.pageIndex
      ).toBe(0);
      expect(serializeOoxmlPart(part)).toBe(source);
    });
  }
  test('preserves a blank page authored by two separate breaks', () => {
    const layout = layoutSemanticDocument(load(table(150) + pageBreak + pageBreak + title), 0, {
      measurer,
    });
    expect(layout.pages).toHaveLength(3);
    expect(textPages(layout)).toEqual(['', '', 'TITLE']);
    expect(
      paragraphs(layout)
        .filter(({ fragment }) =>
          fragment.lines.some((line) => line.spans.some((span) => span.text === '\f'))
        )
        .map(({ pageIndex }) => pageIndex)
    ).toEqual([0, 1]);
  });
  test('preserves repeated breaks within one paragraph', () => {
    const repeated = paragraph('<w:br w:type="page"/><w:br w:type="page"/>');
    const layout = layoutSemanticDocument(load(table(150) + repeated + title), 0, { measurer });
    expect(layout.pages).toHaveLength(3);
    expect(textPages(layout)).toEqual(['', '', 'TITLE']);
  });
  test('does not collapse an explicit break at the start of the document', () => {
    const layout = layoutSemanticDocument(load(pageBreak + title), 0, { measurer });
    expect(textPages(layout)).toEqual(['', 'TITLE']);
  });
  test('retains ordinary geometry when the break line fits', () => {
    const layout = layoutSemanticDocument(load(table(100) + pageBreak + title), 0, { measurer });
    expect(layout.pages).toHaveLength(2);
    const held = paragraphs(layout).find(({ fragment }) =>
      fragment.lines.some((line) => line.spans.some((span) => span.text === '\f'))
    )!;
    expect(held.pageIndex).toBe(0);
    expect(held.fragment.outOfFlow).toBeUndefined();
    expect(held.fragment.lines[0]!.box.y).toBe(100);
  });
  for (const [name, extra] of [
    ['shading', '<w:shd w:val="clear" w:fill="FFFF00"/>'],
    ['border', '<w:pBdr><w:bottom w:val="single" w:sz="8"/></w:pBdr>'],
  ]) {
    test(`keeps the ordinary fit for a standalone break with ${name}`, () => {
      const layout = layoutSemanticDocument(
        load(table(150) + paragraph('<w:br w:type="page"/>', extra) + title),
        0,
        { measurer }
      );
      expect(layout.pages).toHaveLength(3);
    });
  }
  test('does not keep text before a break on an overflowing page', () => {
    const layout = layoutSemanticDocument(
      load(table(150) + paragraph('<w:t>LEAD</w:t><w:br w:type="page"/>') + title),
      0,
      { measurer }
    );
    expect(textPages(layout)).toEqual(['', 'LEAD', 'TITLE']);
  });
  test('retains ordinary empty paragraph pagination', () => {
    const layout = layoutSemanticDocument(load(table(150) + paragraph('') + title), 0, {
      measurer,
    });
    expect(layout.pages).toHaveLength(2);
    expect(
      paragraphs(layout).find(({ fragment }) =>
        fragment.lines.every((line) => line.spans.length === 0)
      )!.pageIndex
    ).toBe(1);
  });
  test('matches a cold layout after edited content and an undo-sized replacement', () => {
    const session = createLayoutSession();
    for (const [revision, content] of [
      pageBreak,
      paragraph('<w:t>LEAD</w:t><w:br w:type="page"/>'),
      pageBreak,
    ].entries()) {
      const part = load(table(150) + content + title);
      const warm = layoutSemanticDocument(part, revision, { measurer, session });
      expect(warm.pages).toEqual(layoutSemanticDocument(part, revision, { measurer }).pages);
      expect(warm.pages).toHaveLength(revision === 1 ? 3 : 2);
    }
  });
});
