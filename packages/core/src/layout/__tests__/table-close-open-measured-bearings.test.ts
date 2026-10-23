import { expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart, type OoxmlElement } from '../../store/index.ts';
import { layoutSemanticDocument } from '../semantic-layout.ts';
import { buildStyleCascadeTable } from '../style-cascade.ts';
import { NO_THEME_FONTS } from '../run-style.ts';
import { paragraphContentWidthsPt } from '../table-autofit-widths.ts';
import type { TextMeasurer, TableFragmentRecord } from '../semantic-records.ts';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const text = '甲》（乙丙。';
// Synthetic proportional metrics isolate intrinsic and flow policy without installed fonts.
const measurer: TextMeasurer = {
  measure: (t, s) =>
    t.length * ((s.fontSizePt * s.horizontalScalePercent) / 100 + s.characterSpacingPt),
  lineMetrics: () => ({ height: 24, baseline: 19.2 }),
  hasResolvedFont: () => true,
  inkBounds: (t, s) =>
    t === '》'
      ? { left: s.fontSizePt * 0.1, right: s.fontSizePt * 0.6 }
      : t === '（'
        ? { left: s.fontSizePt * 0.4, right: s.fontSizePt * 0.95 }
        : { left: 0, right: s.fontSizePt },
};
function part(xml: string, name: string) {
  const r = readOoxmlPart(xml, { name, contentType: 'application/xml' });
  if (!r.ok) throw Error(r.reason);
  return r.part;
}
interface Options {
  alignment: string;
  opening: number;
  closing: number;
  noPunc: boolean;
  ea?: boolean;
  inherited?: boolean;
}
function inspect(o: Options) {
  const jc = `<w:jc w:val="${o.alignment}"/>`;
  const pp = o.inherited ? '<w:pStyle w:val="InheritedAlignment"/>' : jc;
  const run = (t: string, size: number) =>
    `<w:r><w:rPr><w:rFonts w:eastAsia="Synthetic EA" w:hint="eastAsia"/><w:sz w:val="${size * 2}"/><w:kern w:val="2"/></w:rPr><w:t>${t}</w:t></w:r>`;
  const paragraph = `<w:p><w:pPr>${pp}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/><w:overflowPunct w:val="0"/></w:pPr>${run('甲', o.opening)}${run('》', o.closing)}${run('（乙丙。', o.opening)}</w:p>`;
  const wrap = (body: string) => `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`;
  const source = part(wrap(paragraph), '/word/document.xml');
  const styles = part(
    `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>${o.ea === false ? '' : '<w:rFonts w:eastAsia="Synthetic EA"/>'}</w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="InheritedAlignment"><w:pPr>${jc}</w:pPr></w:style></w:styles>`,
    '/word/styles.xml'
  );
  const settings = part(
    `<w:settings xmlns:w="${W}"><w:characterSpacingControl w:val="compressPunctuation"/>${o.noPunc ? '<w:noPunctuationKerning/>' : ''}</w:settings>`,
    '/word/settings.xml'
  );
  const styleCascade = buildStyleCascadeTable(
    styles.root,
    { ...NO_THEME_FONTS, eastAsianFontTableFaces: ['synthetic ea'] },
    settings.root
  );
  const body =
    source.root.children.find(
      (node) => node.kind === 'generic' && 'localName' in node && node.localName === 'body'
    ) ?? source.root.children[0]!;
  if (!('children' in body)) throw Error('expected body');
  const node = body.children.find((child) => child.kind === 'paragraph') as OoxmlElement;
  const before = serializeOoxmlPart(source);
  const widths = paragraphContentWidthsPt(node, {
    context: { measurer },
    view: {
      styleCascade,
      displayMode: 'all-markup',
      authorFilter: undefined,
      readNested: () => null,
    },
    tableCellStyle: undefined,
  });
  expect(serializeOoxmlPart(source)).toBe(before);
  const natural = o.opening * 5 + o.closing;
  const width = natural - o.opening * 0.75;
  const tableSource = part(
    wrap(
      '<w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/><w:tblCellMar><w:top w:w="0"/><w:left w:w="0"/><w:bottom w:w="0"/><w:right w:w="0"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="1"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="auto" w:w="0"/></w:tcPr>' +
        paragraph +
        '</w:tc></w:tr></w:tbl>'
    ),
    '/word/document.xml'
  );
  const beforeTable = serializeOoxmlPart(tableSource);
  const model = layoutSemanticDocument(tableSource, 0, {
    measurer,
    styleCascade,
    geometry: {
      width: width + 20,
      height: 500,
      margin: { left: 10, right: 10, top: 10, bottom: 10 },
    },
  });
  expect(model.pages).toHaveLength(1);
  const table = model.pages[0]!.fragments.find(
    (block): block is TableFragmentRecord => block.kind === 'table'
  );
  expect(table).toBeDefined();
  const block = table!.rows[0]!.cells[0]!.blocks[0]!;
  if (block.kind !== 'paragraph') throw Error('expected cell paragraph');
  expect(
    block.lines
      .flatMap((line) => line.spans)
      .map((span) => span.text)
      .join('')
  ).toBe(text);
  expect(serializeOoxmlPart(tableSource)).toBe(beforeTable);
  return { widths, natural, table: table!, block };
}
for (const opening of [10.5, 12])
  for (const closing of [opening, 20])
    for (const alignment of ['both', 'left', 'center', 'right'])
      for (const noPunc of [false, true]) {
        test(`AutoFit shares pair policy at ${closing}/${opening} pt, ${alignment}, noPunc=${noPunc}`, () => {
          const a = inspect({ opening, closing, alignment, noPunc });
          // Legal CJK chunks retain their own full advances at the minimum width.
          expect(a.widths.min).toBeCloseTo(Math.max(opening + closing, opening * 2), 7);
          // Maximum width retains the complete, naturally unwrapped source line.
          // A demand fit must not lower the table's preferred width.
          expect(a.widths.max).toBeCloseTo(a.natural, 7);
          expect(a.block.lines).toHaveLength(alignment === 'both' ? 1 : 2);
          for (const line of a.block.lines) expect(line.box.height).toBeCloseTo(12, 7);
          if (a.block.lines.length === 2)
            expect(a.block.lines[1]!.box.y - a.block.lines[0]!.box.y).toBeCloseTo(12, 7);
        });
      }
test('AutoFit and cell flow inherit the same justified paragraph style', () => {
  const a = inspect({
    opening: 12,
    closing: 20,
    alignment: 'both',
    noPunc: false,
    inherited: true,
  });
  expect(a.widths.min).toBeCloseTo(32, 7);
  expect(a.widths.max).toBeCloseTo(80, 7);
  expect(a.block.lines).toHaveLength(1);
});
test('missing authored EA defaults preserve natural AutoFit maximum and the two-line cell', () => {
  const a = inspect({ opening: 12, closing: 12, alignment: 'both', noPunc: false, ea: false });
  expect(a.widths.max).toBeCloseTo(72, 7);
  expect(a.block.lines).toHaveLength(2);
});
