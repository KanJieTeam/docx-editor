import { describe, expect, test } from 'bun:test';
import { readOoxmlPart } from '@docx-editor.dev/core/store';
import { buildNumberingIndex } from '../numbering-index.ts';
import { paragraphFragmentsOf } from '../semantic-records.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../semantic-layout.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer = createFixedMeasurer(6, 14);

function fixture(
  indent: string,
  levelText = '',
  format = 'none',
  {
    levelIndent = 'w:left="0" w:hanging="0"',
    text = 'Heading text',
    bidi = false,
    numbered = true,
  } = {}
) {
  const doc = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:pPr>
      ${numbered ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : ''}
      ${bidi ? '<w:bidi/>' : ''}
      <w:ind ${indent}/></w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr>
      <w:t>${text}</w:t></w:r></w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  const num = readOoxmlPart(
    `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${format}"/>
      <w:lvlText w:val="${levelText}"/><w:suff w:val="nothing"/>
      <w:pPr><w:ind ${levelIndent}/></w:pPr>
      <w:rPr><w:sz w:val="22"/></w:rPr></w:lvl></w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`,
    { name: '/word/numbering.xml', contentType: 'app/xml' }
  );
  if (!doc.ok || !num.ok) throw new Error('Invalid synthetic fixture');
  const layout = layoutSemanticDocument(doc.part, 1, {
    measurer,
    numberingIndex: buildNumberingIndex(num.part.root),
    geometry: { width: 140, height: 400, margin: { top: 20, right: 20, bottom: 20, left: 20 } },
  });
  return paragraphFragmentsOf(layout.pages[0]!)[0]!;
}

describe('paragraph indents without a painted list marker', () => {
  test('keeps a positive first-line indent on a none level', () => {
    const paragraph = fixture('w:firstLine="720"');
    expect(paragraph.indent.firstLine).toBe(36);
    expect(paragraph.marker).toBeUndefined();
    expect(paragraph.lines[0]!.spans[0]!.box.x).toBe(36);
  });

  test('keeps a hanging first line when the marker is empty', () => {
    const paragraph = fixture('w:left="720" w:hanging="360"');
    expect(paragraph.indent.hanging).toBe(18);
    expect(paragraph.marker).toBeUndefined();
    expect(paragraph.lines[0]!.spans[0]!.box.x).toBe(18);
  });

  test('inherits a hanging indent from an empty numbering level', () => {
    const paragraph = fixture('', '', 'none', { levelIndent: 'w:left="720" w:hanging="360"' });
    expect(paragraph.indent).toMatchObject({ left: 36, hanging: 18 });
    expect(paragraph.lines[0]!.spans[0]!.box.x).toBe(18);
  });

  test('keeps the continuation line at the content indent', () => {
    const paragraph = fixture('w:firstLine="720"', '', 'none', { text: 'Alpha beta gamma delta' });
    expect(paragraph.lines.length).toBeGreaterThan(1);
    expect(paragraph.lines[0]!.contentX).toBe(36);
    expect(paragraph.lines[1]!.contentX).toBe(0);
  });

  test('places a right-to-left empty marker like an ordinary paragraph', () => {
    const indent = 'w:left="720" w:firstLine="360"';
    const list = fixture(indent, '', 'none', { bidi: true });
    const plain = fixture(indent, '', 'none', { bidi: true, numbered: false });
    expect(list.marker).toBeUndefined();
    expect(list.lines[0]!.contentX).toBe(plain.lines[0]!.contentX);
    expect(list.lines[0]!.spans[0]!.box).toEqual(plain.lines[0]!.spans[0]!.box);
  });

  test('still places text after a visible marker', () => {
    const paragraph = fixture('w:firstLine="720"', '%1.', 'decimal');
    expect(paragraph.marker?.text).toBe('1.');
    expect(paragraph.lines[0]!.spans[0]!.box.x).toBe(48);
  });
});
