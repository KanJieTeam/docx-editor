import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart } from '../../store/package/ooxml-tree.ts';
import { layoutSemanticDocument } from '../semantic-layout.ts';
import { caretAt } from '../semantic-interaction.ts';
import { cellAddressAt, tableAnchorAt, tableContextAt } from '../semantic-cell-selection.ts';
import { planTableCommand } from '../../editor/table-command-plan.ts';
import type {
  ParagraphFragmentRecord,
  SemanticLayout,
  TableFragmentRecord,
} from '../semantic-records.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const geometry = { width: 612, height: 720, margin: { top: 40, bottom: 40, left: 40, right: 40 } };
const measurer = {
  measure: (text: string) => [...text].length * 5,
  lineMetrics: () => ({ height: 12, baseline: 10 }),
};
const p = (runs: string) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>${runs}</w:p>`;
const text = (value: string) =>
  `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${value}</w:t></w:r>`;
const field = (value: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> = 1 \\* ROMAN </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  text(value) +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

function fixture({
  exact = true,
  fields = false,
  center = false,
  count = 42,
  fitting = false,
  unknown = false,
  borders = false,
} = {}) {
  const merged = Array.from({ length: count }, (_, index) =>
    p(
      text(`MERGED-LINE-${String(index + 1).padStart(2, '0')} `) +
        (fields ? field('I') : text('text'))
    )
  ).join('');
  const prefix = fitting
    ? ''
    : '<w:p><w:pPr><w:spacing w:line="4400" w:lineRule="exact"/></w:pPr>' +
      text('Introduction') +
      '</w:p>';
  const tableProperties =
    '<w:tblPr><w:tblW w:w="10000" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar><w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
      .map((side) => `<w:${side} w:val="${borders ? 'single' : 'nil'}" w:sz="4"/>`)
      .join('') +
    '</w:tblBorders></w:tblPr>';
  const header =
    '<w:tr><w:trPr><w:trHeight w:val="360" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>' +
    p(text('Header')) +
    '</w:tc></w:tr>';
  const body = Array.from(
    { length: 30 },
    (_, index) =>
      '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="360" w:hRule="' +
      (exact ? 'exact' : 'atLeast') +
      '"/></w:trPr>' +
      '<w:tc><w:tcPr><w:vMerge' +
      (index === 0 ? ' w:val="restart"' : '') +
      '/><w:vAlign w:val="' +
      (center ? 'center' : 'top') +
      '"/><w:shd w:fill="D9E2F3"/></w:tcPr>' +
      (index === 0 ? (unknown ? '<w:p><w:r><w:br/></w:r></w:p>' : '') + merged : p('')) +
      '</w:tc>' +
      '<w:tc>' +
      p(text(`ROW-${String(index + 1).padStart(2, '0')}`)) +
      '</w:tc></w:tr>'
  ).join('');
  const xml = `<w:document xmlns:w="${W}"><w:body>${prefix}<w:tbl>${tableProperties}<w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid>${header}${body}</w:tbl></w:body></w:document>`;
  const loaded = readOoxmlPart(xml, { name: '/word/document.xml', contentType: 'application/xml' });
  if (!loaded.ok) throw new Error(loaded.reason);
  return loaded.part;
}

function tables(layout: SemanticLayout) {
  return layout.pages.map((page) =>
    page.fragments.filter((fragment): fragment is TableFragmentRecord => fragment.kind === 'table')
  );
}

function paragraphs(
  layout: SemanticLayout
): { page: number; rowId: string; cellId: string; block: ParagraphFragmentRecord }[] {
  const result: ReturnType<typeof paragraphs> = [];
  for (const [pageIndex, fragments] of tables(layout).entries())
    for (const table of fragments)
      for (const row of table.rows)
        for (const cell of row.cells)
          for (const block of cell.blocks) {
            if (block.kind === 'paragraph')
              result.push({ page: pageIndex + 1, rowId: row.id, cellId: cell.id, block });
          }
  return result;
}

function merged(layout: SemanticLayout) {
  return paragraphs(layout).filter(({ block }) =>
    block.lines.some((line) =>
      line.spans
        .map((span) => span.text)
        .join('')
        .startsWith('MERGED-LINE-')
    )
  );
}

function rowsByPage(layout: SemanticLayout) {
  return layout.pages.map(
    (_, index) =>
      paragraphs(layout)
        .filter((item) => item.page === index + 1)
        .flatMap(({ block }) => block.lines.flatMap((line) => line.spans.map((span) => span.text)))
        .filter((value) => value.startsWith('ROW-')).length
  );
}

describe('merged text across protected physical rows', () => {
  for (const exact of [true, false])
    for (const fields of [true, false]) {
      test(`${exact ? 'exact' : 'atLeast'} rows carry every ${fields ? 'field' : 'paragraph'} without changing the tree`, () => {
        const part = fixture({ exact, fields });
        const before = serializeOoxmlPart(part);
        const layout = layoutSemanticDocument(part, 0, { geometry, measurer });
        const content = merged(layout);
        expect(rowsByPage(layout)).toEqual([22, 8]);
        expect(content).toHaveLength(42);
        expect(new Set(content.map((item) => item.block.paragraphId)).size).toBe(42);
        expect(content.filter((item) => item.page === 1)).toHaveLength(33);
        expect(content.filter((item) => item.page === 2)).toHaveLength(9);
        expect(new Set(content.map((item) => item.cellId)).size).toBe(1);
        expect(new Set(content.map((item) => item.rowId)).size).toBe(1);
        for (const { page, block } of content) {
          const cursor = caretAt(layout, { paragraphId: block.paragraphId, offset: 0 });
          expect(cursor?.pageIndex).toBe(page - 1);
          expect(block.box.y).toBeGreaterThanOrEqual(0);
          expect(block.box.y + block.box.height).toBeLessThanOrEqual(640.001);
          const atom = block.lines.flatMap((line) => line.spans).find((span) => span.fieldAtom);
          if (fields) {
            expect(atom?.range.end! - atom?.range.start!).toBe(1);
            expect(atom?.projected).toBe(true);
            expect(atom?.fieldAtom?.formField).toBe(false);
            const beforeAtom = caretAt(
              layout,
              { paragraphId: block.paragraphId, offset: atom!.range.start },
              measurer
            );
            const afterAtom = caretAt(
              layout,
              { paragraphId: block.paragraphId, offset: atom!.range.end },
              measurer
            );
            expect(beforeAtom?.pageIndex).toBe(page - 1);
            expect(afterAtom?.pageIndex).toBe(page - 1);
            expect(afterAtom!.x - beforeAtom!.x).toBeCloseTo(atom!.box.width, 4);
          }
        }
        expect(serializeOoxmlPart(part)).toBe(before);
        const reopened = readOoxmlPart(before, {
          name: '/word/document.xml',
          contentType: 'application/xml',
        });
        if (!reopened.ok) throw new Error(reopened.reason);
        expect(layoutSemanticDocument(reopened.part, 0, { geometry, measurer })).toEqual(layout);
      });
    }

  test('centers short content within the current fragment and keeps shading on both pages', () => {
    const layout = layoutSemanticDocument(fixture({ center: true, count: 1 }), 0, {
      geometry,
      measurer,
    });
    const content = merged(layout);
    expect(content).toHaveLength(1);
    const fragment = tables(layout)[0]![0]!;
    const row = fragment.rows.find((row) =>
      row.cells.some((cell) => cell.id === content[0]!.cellId)
    )!;
    const cell = row.cells.find((cell) => cell.id === content[0]!.cellId)!;
    expect(content[0]!.block.box.y - cell.box.y).toBeCloseTo((cell.box.height - 12) / 2, 4);
    for (const fragments of tables(layout)) {
      const head = fragments[0]!.rows
        .flatMap((row) => row.cells)
        .find((cell) => cell.id === content[0]!.cellId)!;
      expect(head.shading).toBe('D9E2F3');
      expect(head.paintInert).toBe(false);
      expect(head.box.height).toBeGreaterThan(0);
    }
  });

  test('keeps the ordinary whole-span layout when every row fits this page', () => {
    const layout = layoutSemanticDocument(
      fixture({ fitting: true, center: true, fields: true }),
      0,
      { geometry, measurer }
    );
    expect(rowsByPage(layout)).toEqual([30]);
    expect(merged(layout)).toHaveLength(42);
    expect(
      tables(layout)[0]![0]!.rows.some((row) => row.isContinuation && row.box.height === 0)
    ).toBe(false);
  });

  test('leaves manual-break content on the existing conservative path', () => {
    const layout = layoutSemanticDocument(fixture({ unknown: true }), 0, { geometry, measurer });
    expect(
      tables(layout)
        .flat()
        .flatMap((table) => table.rows)
        .some((row) => row.isContinuation && row.box.height === 0)
    ).toBe(false);
  });

  test('paints the carried head rules without an extra zero-height bottom rule', () => {
    const layout = layoutSemanticDocument(fixture({ borders: true, center: true }), 0, {
      geometry,
      measurer,
    });
    expect(merged(layout)).toHaveLength(42);
    const carried = tables(layout)[1]![0]!.rows[0]!;
    expect(carried.isContinuation).toBe(true);
    expect(carried.box.height).toBe(0);
    const head = carried.cells.find((cell) => cell.blocks.length > 0)!;
    expect(head.id).toBe(merged(layout)[0]!.cellId);
    expect(head.box.height).toBeGreaterThan(0);
    expect(head.borders?.edgeSegments?.some((edge) => edge.side === 'left')).toBe(true);
    for (const cell of carried.cells.filter((cell) => cell.box.height === 0)) {
      expect(cell.borders?.edgeSegments?.some((edge) => edge.side === 'bottom') ?? false).toBe(
        false
      );
    }
  });

  test('commands and selections in later text retain the authored head address', () => {
    const part = fixture({ fields: true });
    const layout = layoutSemanticDocument(part, 0, { geometry, measurer });
    const head = merged(layout)[0]!;
    const tail = merged(layout).at(-1)!;
    expect(tail.page).toBe(2);
    expect(tableAnchorAt(layout, tail.block.paragraphId)).toMatchObject({
      rowId: head.rowId,
      cellId: head.cellId,
    });
    expect(tableContextAt(layout, tail.block.paragraphId)?.rowIndex).toBe(1);
    expect(cellAddressAt(layout, tail.block.paragraphId)).toMatchObject({
      rowId: head.rowId,
      cellId: head.cellId,
      rowIndex: 1,
    });
    const at = { paragraphId: tail.block.paragraphId, offset: 0 };
    const plan = planTableCommand({
      command: { type: 'insertRow', where: 'above' },
      part,
      layout,
      storeRevision: layout.revision,
      selection: { anchor: at, head: at },
      cellSelection: null,
      themeColors: [],
      editable: true,
      viewing: false,
    });
    expect(plan.ok).toBe(true);
    if (plan.ok) expect(plan.ops[0]).toMatchObject({ op: 'insertTableRow', rowId: head.rowId });
  });
});
