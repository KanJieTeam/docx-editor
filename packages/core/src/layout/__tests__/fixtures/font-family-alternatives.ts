import { strToU8, zipSync } from 'fflate';

/** Public synthetic content. This fixture contains no private report or font bytes. */
export function fontFamilyAlternativesDocx(): Uint8Array {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    ),
    'word/_rels/document.xml.rels': strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr>` +
        '<w:rFonts w:ascii="宋体;SimSun" w:hAnsi="宋体;SimSun" w:eastAsia="宋体;SimSun"/>' +
        '</w:rPr><w:t>PUBLIC FONT TEST 示例字体 ABC 123</w:t></w:r></w:p>' +
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="黑体;SimHei" w:eastAsia="黑体;SimHei"/>' +
        '<w:b/><w:i/></w:rPr><w:t>PUBLIC BOLD ITALIC 示例粗斜体</w:t></w:r></w:p>' +
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="SimSun"/></w:rPr><w:t>PLAIN FONT CONTROL</w:t></w:r></w:p>' +
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
        '</w:body></w:document>'
    ),
    'word/styles.xml': strToU8(
      `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>` +
        '<w:rFonts w:ascii="宋体;SimSun" w:eastAsia="宋体;SimSun"/><w:sz w:val="28"/>' +
        '</w:rPr></w:rPrDefault></w:docDefaults></w:styles>'
    ),
  });
}
