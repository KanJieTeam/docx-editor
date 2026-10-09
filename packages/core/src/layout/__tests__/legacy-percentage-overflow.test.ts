import { describe, expect, test } from 'bun:test';
import {
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlElement,
} from '../../store/package/ooxml-tree.ts';
import { readTableStructure, tableOriginX } from '../semantic-table.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function fixture(width = '5500', grid = '2438', alignment = 'center', extra = '') {
  return (
    `<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr>` +
    `<w:tblW w:type="pct" w:w="${width}"/><w:jc w:val="${alignment}"/>` +
    '<w:tblLayout w:type="autofit"/><w:tblCellMar>' +
    '<w:left w:type="dxa" w:w="108"/><w:right w:type="dxa" w:w="108"/>' +
    `</w:tblCellMar>${extra}</w:tblPr><w:tblGrid><w:gridCol w:w="${grid}"/></w:tblGrid>` +
    '<w:tr><w:tc><w:tcPr><w:tcW w:type="pct" w:w="5000"/></w:tcPr>' +
    '<w:p><w:r><w:t>314159.26</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>'
  );
}
function open(xml: string, mode: number | undefined = 14, depth = 0) {
  const parsed = readOoxmlPart(xml, { name: '/word/document.xml', contentType: 'application/xml' });
  if (!parsed.ok) throw new Error(parsed.reason);
  const body = parsed.part.root.children.find(
    (n) => n.kind !== 'textValue' && n.localName === 'body'
  ) as OoxmlElement;
  const table = body.children.find((n) => n.kind === 'table')!;
  const before = serializeOoxmlPart(parsed.part);
  const result = readTableStructure(table, 100, depth, undefined, 'all-markup', undefined, mode)!;
  expect(serializeOoxmlPart(parsed.part)).toBe(before);
  return result;
}
describe('grid-confirmed legacy percentage overflow', () => {
  for (const mode of [11, 12, 14]) {
    test(`mode ${mode} retains a centered grid and the original width preference`, () => {
      const table = open(fixture(), mode);
      expect(table.legacyContentAlignment).toBe(true);
      expect(table.columnWidthsPt).toEqual([121.9]);
      expect(tableOriginX(table, 100)).toBeCloseTo(-10.95, 8);
      expect(table.tableWidth).toEqual({ type: 'pct', value: 100 });
    });
  }
  test('the grid can confirm the nearest twip, not a different authored width', () => {
    expect(open(fixture('5209', '2309')).legacyContentAlignment).toBe(true);
    expect(open(fixture('5209', '2308')).legacyContentAlignment).toBeUndefined();
  });
  test('the percentage range keeps its upper bound', () => {
    expect(open(fixture('10000', '4432')).legacyContentAlignment).toBe(true);
    expect(open(fixture('10001', '4432')).legacyContentAlignment).toBeUndefined();
  });
  test('implicit placement and ambiguous source properties do not enable overflow', () => {
    for (const xml of [
      fixture().replace('<w:jc w:val="center"/>', ''),
      fixture().replace('<w:tblLayout w:type="autofit"/>', ''),
      fixture('5500', '2438', 'center', '<w:tblW w:type="pct" w:w="5500"/>'),
    ])
      expect(open(xml).legacyContentAlignment).toBeUndefined();
  });
  for (const [name, xml] of [
    ['left', fixture('5500', '2438', 'left')],
    ['right', fixture('5500', '2438', 'right')],
    ['unbounded', fixture('999999')],
    ['percent suffix', fixture('110%')],
    ['different grid', fixture('5500', '2400')],
    ['nonzero indent', fixture('5500', '2438', 'center', '<w:tblInd w:type="dxa" w:w="20"/>')],
    ['float', fixture('5500', '2438', 'center', '<w:tblpPr w:horzAnchor="text"/>')],
    ['fixed', fixture().replace('autofit', 'fixed')],
    ['spacing', fixture('5500', '2438', 'center', '<w:tblCellSpacing w:type="dxa" w:w="20"/>')],
  ])
    test(`${name} retains the ordinary width path`, () => {
      expect(open(xml!).legacyContentAlignment).toBeUndefined();
    });
  test('modern and nested tables retain the ordinary width path', () => {
    for (const mode of [15, 13, 99])
      expect(open(fixture(), mode).legacyContentAlignment).toBeUndefined();
    expect(open(fixture(), 14, 1).legacyContentAlignment).toBeUndefined();
  });
});
