import { strToU8, zipSync } from 'fflate';

export const STYLES = [
  ['Normal', 'Normal', 'Testo normale'],
  ['Title', 'Title', 'Titolo'],
  ['Subtitle', 'Subtitle', 'Sottotitolo'],
  ['Heading1', 'heading 1', 'Titolo 1'],
  ['Heading2', 'heading 2', 'Titolo 2'],
  ['Heading3', 'heading 3', 'Titolo 3'],
  ['CustomH1', 'Heading 1', 'Heading 1'],
  ['heading1', 'Custom case', 'Custom case'],
  ['ListParagraph', 'List Paragraph', 'List Paragraph'],
  ['constructor', 'Prototype label', 'Prototype label'],
] as const;
export const ITALIAN = {
  _lang: 'it',
  styles: {
    normalText: 'Testo normale',
    title: 'Titolo',
    subtitle: 'Sottotitolo',
    heading1: 'Titolo 1',
    heading2: 'Titolo 2',
    heading3: 'Titolo 3',
  },
};
export const FRENCH = {
  _lang: 'fr',
  styles: {
    normalText: 'Texte normal',
    title: 'Titre',
    subtitle: 'Sous-titre',
    heading1: 'Titre 1',
    heading2: 'Titre 2',
    heading3: 'Titre 3',
  },
};
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
export const SOURCE = zipSync({
  '[Content_Types].xml': strToU8(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
  ),
  '_rels/.rels': strToU8(
    `<Relationships xmlns="${REL}"><Relationship Id="main" Type="${OD}officeDocument" Target="word/document.xml"/></Relationships>`
  ),
  'word/_rels/document.xml.rels': strToU8(
    `<Relationships xmlns="${REL}"><Relationship Id="styles" Type="${OD}styles" Target="styles.xml"/></Relationships>`
  ),
  'word/styles.xml': strToU8(
    `<w:styles xmlns:w="${W}">` +
      STYLES.map(
        ([id, name]) =>
          `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>`
      ).join('') +
      '</w:styles>'
  ),
  'word/document.xml': strToU8(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Public style text</w:t></w:r></w:p></w:body></w:document>`
  ),
});
