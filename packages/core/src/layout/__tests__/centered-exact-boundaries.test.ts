import { expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { applyLineSpacing as publicSpacing } from '@docx-editor.dev/core/layout';
import { applyLineSpacing } from '../paragraph-style.ts';
import { layoutSemanticDocument } from '../semantic-layout.ts';
import { linesOf, paragraphFragmentsOf } from '../semantic-records.ts';
import { buildStyleCascadeTable } from '../style-cascade.ts';
import { layoutHeaderFooterStory } from '../hf-layout.ts';
import {
  indexInlineDrawingProjectionsInPart,
  projectDrawing,
  DEFAULT_DRAWING_PROJECTION_LIMITS,
} from '../../store/package/drawing-projection.ts';
import type { TextMeasurer } from '@docx-editor.dev/core/layout';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function part(xml: string, name = '/word/document.xml') {
  const parsed = readOoxmlPart(xml, { name, contentType: 'application/xml' });
  if (!parsed.ok) throw new Error(parsed.reason);
  return parsed.part;
}
const run = (text: string, family = 'Tall') =>
  `<w:r><w:rPr><w:rFonts w:ascii="${family}" w:eastAsia="${family}"/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (runs: string, exact = 24) =>
  `<w:p><w:pPr><w:spacing w:line="${exact * 20}" w:lineRule="exact"/></w:pPr>${runs}</w:p>`;
function documentPart(body: string) {
  return part(
    `<w:document xmlns:w="${W}" xmlns:m="${M}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`
  );
}
const uniform: TextMeasurer = {
  measure: (text) => text.length * 5,
  lineMetrics: () => ({ height: 14, baseline: 11 }),
};
const measuredMetrics: TextMeasurer = {
  measure: (text) => text.length * 5,
  lineMetrics: () => ({ height: 18.2, baseline: 14.13125 }),
};
function cascade(compat = '') {
  return buildStyleCascadeTable(
    part(
      `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="28"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
      '/word/styles.xml'
    ).root,
    undefined,
    part(
      `<w:settings xmlns:w="${W}"><w:compat>${compat}</w:compat></w:settings>`,
      '/word/settings.xml'
    ).root
  );
}
function laid(body: string, measurer = uniform, compat = '') {
  return layoutSemanticDocument(documentPart(body), 0, { measurer, styleCascade: cascade(compat) });
}

test('center boundary 01: public and direct source share one implementation and natural line metrics', () => {
  expect(publicSpacing).toBe(applyLineSpacing);
  expect(applyLineSpacing({ rule: 'exact', value: 31.2 }, 18.2, 14.13125)).toEqual({
    height: 31.2,
    baseline: 20.63125,
  });
  expect(
    applyLineSpacing({ rule: 'exact', value: 31.2, preserveExactBaseline: true }, 18.2, 14.13125)
  ).toEqual({ height: 31.2, baseline: 14.13125 });
});

test('center boundary 02: mixed-face metrics aggregation keeps both ascents and descents', () => {
  const measurer: TextMeasurer = {
    measure: uniform.measure,
    lineMetrics: (style) =>
      style.fontFamily === 'Deep' ? { height: 12, baseline: 7 } : { height: 12, baseline: 10 },
  };
  for (const families of [
    ['Tall', 'Deep'],
    ['Deep', 'Tall'],
  ]) {
    const content = run('A ', families[0]) + run('g', families[1]);
    const natural = linesOf(laid(`<w:p>${content}</w:p>`, measurer))[0]!;
    expect(natural.box.height).toBe(15);
    expect(natural.baseline).toBe(10);
    const exact = linesOf(laid(paragraph(content), measurer))[0]!;
    expect(exact.spans.map((span) => span.text).join('')).toBe('A g');
    expect(exact.box.height).toBe(24);
    expect(exact.baseline).toBe(14.5);
  }
});

test('center boundary 03: first middle and last exact lines keep authored line advance', () => {
  const content =
    run('First') + '<w:r><w:br/></w:r>' + run('Middle') + '<w:r><w:br/></w:r>' + run('Last');
  const lines = linesOf(laid(paragraph(content, 31.2), measuredMetrics));
  expect(lines).toHaveLength(3);
  lines.forEach((line, index) => {
    expect(line.box.y).toBeCloseTo(index * 31.2, 8);
    expect(line.box.height).toBe(31.2);
    expect(line.baseline).toBeCloseTo(20.63125, 8);
    expect(line.trailingSpacing).toBe(0);
  });
});

test('center boundary 04: empty exact paragraphs keep a real band and noExtra baseline', () => {
  const content = paragraph('', 20);
  const normal = linesOf(laid(content))[0]!;
  const preserve = linesOf(laid(content, uniform, '<w:noExtraLineSpacing/>'))[0]!;
  expect(normal.spans).toHaveLength(0);
  expect(normal.box.height).toBe(20);
  expect(normal.baseline).toBe(14);
  expect(preserve.box.height).toBe(20);
  expect(preserve.baseline).toBe(11);
});

test('center boundary 05: noLeading does not silently enable a distinct exact or page-end flag', () => {
  const content = paragraph(run('Text'), 20);
  const ordinary = linesOf(laid(content))[0]!;
  const noLeading = linesOf(laid(content, uniform, '<w:noLeading/>'))[0]!;
  const explicitOff = linesOf(laid(content, uniform, '<w:noExtraLineSpacing w:val="0"/>'))[0]!;
  for (const line of [ordinary, noLeading, explicitOff]) {
    expect(line.box.height).toBe(20);
    expect(line.baseline).toBe(14);
    expect(line.trailingSpacing).toBe(0);
  }
});

test('center boundary 06: header and table use the same centred band without changing authored height', () => {
  const p = paragraph(run('Text'), 20);
  const styleCascade = cascade();
  const header = part(`<w:hdr xmlns:w="${W}">${p}</w:hdr>`, '/word/header1.xml');
  const story = layoutHeaderFooterStory(
    header,
    200,
    uniform,
    'exact-center',
    undefined,
    styleCascade
  );
  const headerLines = story.fragments.flatMap((block) =>
    block.kind === 'paragraph' ? block.lines : []
  );
  expect(headerLines).toHaveLength(1);
  expect(headerLines[0]!.box.height).toBe(20);
  expect(headerLines[0]!.baseline).toBe(14);
  const layout = laid(
    `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${p}</w:tc></w:tr></w:tbl>`
  );
  const table = layout.pages[0]!.fragments.find((block) => block.kind === 'table')!;
  const cell = table.rows[0]!.cells[0]!.blocks.find((block) => block.kind === 'paragraph')!;
  expect(cell.lines[0]!.box.height).toBe(20);
  expect(cell.lines[0]!.baseline).toBe(14);
});

test('center boundary 07: undersized exact math retains its atom and following-line position', () => {
  const math =
    '<m:oMath><m:f><m:num><m:r><m:t>aa</m:t></m:r></m:num><m:den><m:r><m:t>bb</m:t></m:r></m:den></m:f></m:oMath>';
  const source = documentPart(paragraph(math, 8) + paragraph(run('After'), 8));
  const canonical = serializeOoxmlPart(source);
  const layout = layoutSemanticDocument(source, 0, { measurer: uniform, styleCascade: cascade() });
  const lines = linesOf(layout);
  expect(lines).toHaveLength(2);
  const equation = lines[0]!.spans.find((span) => span.equation)!;
  expect(equation.equation!.geometry.box.height).toBeGreaterThan(8);
  expect(equation.range.end - equation.range.start).toBe(1);
  expect(lines[0]!.box.height).toBe(8);
  expect(lines[0]!.baseline).toBeGreaterThanOrEqual(0);
  expect(lines[0]!.baseline).toBeLessThanOrEqual(8);
  expect(lines[1]!.box.y).toBe(8);
  expect(serializeOoxmlPart(source)).toBe(canonical);
});

test('center boundary 08: undersized exact inline image keeps authored extent and line advance', () => {
  const image =
    '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="image"/><a:graphic><a:graphicData uri="' +
    PIC +
    '"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const source = documentPart(paragraph(image, 8) + paragraph(run('After'), 8));
  const canonical = serializeOoxmlPart(source);
  const atoms = indexInlineDrawingProjectionsInPart(source);
  const layout = layoutSemanticDocument(source, 0, {
    measurer: uniform,
    styleCascade: cascade(),
    inlineDrawingLayout: {
      ownerPartName: '/word/document.xml',
      projectionForAtom: (id) => atoms.get(id) ?? null,
      project: (node) =>
        atoms.get(node.id) ??
        projectDrawing(node, {
          ownerPartName: '/word/document.xml',
          limits: DEFAULT_DRAWING_PROJECTION_LIMITS,
        }),
      resourceOf: () => ({
        kind: 'ready',
        partName: '/word/media/image.png',
        contentId: 'image',
        resourceKey: 'image',
        mime: 'image/png',
        pixelWidth: 96,
        pixelHeight: 96,
        dpiX: 96,
        dpiY: 96,
      }),
    },
  });
  const lines = paragraphFragmentsOf(layout.pages[0]!).flatMap((p) => p.lines);
  expect(lines).toHaveLength(2);
  expect(lines[0]!.drawings).toHaveLength(1);
  expect(lines[0]!.drawings![0]!.height).toBe(72);
  expect(lines[0]!.box.height).toBe(8);
  expect(lines[1]!.box.y).toBe(8);
  expect(serializeOoxmlPart(source)).toBe(canonical);
});

test('center boundary 09: auto atLeast and section-grid placement keep their original contracts', () => {
  expect(applyLineSpacing({ rule: 'auto', value: 480 }, 14, 11)).toEqual({
    height: 28,
    baseline: 11,
  });
  expect(applyLineSpacing({ rule: 'atLeast', value: 20 }, 14, 11)).toEqual({
    height: 20,
    baseline: 17,
    trailing: 0,
  });
  expect(applyLineSpacing({ rule: 'atLeast', value: 10 }, 14, 11)).toEqual({
    height: 14,
    baseline: 11,
    trailing: 0,
  });
  expect(applyLineSpacing({ rule: 'auto', value: 240, gridPitch: 18 }, 14, 11.2)).toEqual({
    height: 18,
    baseline: 13.2,
    trailing: 2,
  });
  expect(applyLineSpacing({ rule: 'exact', value: 8 }, 24, 18)).toEqual({ height: 8, baseline: 8 });
});

test('center boundary 10: absent bottom suppression keeps the overflowing exact line on the next page', () => {
  const source = documentPart(paragraph(run('Lead'), 699.4) + paragraph(run('Fourth'), 31.2));
  const layout = layoutSemanticDocument(source, 0, {
    measurer: measuredMetrics,
    styleCascade: cascade(),
    geometry: {
      width: 595.3,
      height: 725.251,
      margin: { top: 0, bottom: 0, left: 0, right: 0 },
    },
  });
  expect(layout.pages).toHaveLength(2);
  const first = paragraphFragmentsOf(layout.pages[0]!)[0]!.lines[0]!;
  const fourth = paragraphFragmentsOf(layout.pages[1]!)[0]!.lines[0]!;
  expect(first.box.height).toBe(699.4);
  expect(fourth.box.height).toBe(31.2);
  expect(fourth.box.y).toBe(0);
  expect(fourth.baseline).toBeCloseTo(20.63125, 8);
  expect(fourth.trailingSpacing).toBe(0);
});
