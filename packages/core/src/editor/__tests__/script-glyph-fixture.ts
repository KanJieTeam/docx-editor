import { readFileSync } from 'node:fs';
import type { DocxEditorInstance } from '../docx-editor.ts';
import { linesOf } from '../../layout/semantic-records.ts';
import { glyphSizeFactorOf } from '../../layout/run-style.ts';
import { sha256FontBytes } from '../../layout/font-resource.ts';
import { docx } from './paginated-surface-fixtures.ts';

const bytes = new Uint8Array(
  readFileSync(new URL('../../layout/__tests__/fixtures/fonts/DejaVuSans.ttf', import.meta.url))
);
export const scriptGlyphFonts = {
  sources: [
    {
      request: { family: 'DejaVu Sans', weight: 400, style: 'normal' as const },
      id: 'script-size-fixture',
      bytes,
      hash: sha256FontBytes(bytes),
      faceIndex: 0,
    },
  ],
};
export const scriptGlyphDocument = docx(
  '<w:p><w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="28"/><w:vertAlign w:val="subscript"/></w:rPr><w:t>123</w:t></w:r></w:p>'
);

export async function settledScriptGlyphs(editor: DocxEditorInstance): Promise<void> {
  for (let attempt = 0; attempt < 200; attempt++) {
    if (
      editor.surface &&
      !editor.fontMeasurement().resolving &&
      editor.fontMeasurement().measurer === 'shaped'
    )
      return;
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
  throw new Error('Script fixture did not finish font-backed layout');
}

export function scriptGlyphGeometry(editor: DocxEditorInstance) {
  const span = linesOf(editor.surface!.layout())
    .flatMap((line) => line.spans)
    .find((span) => span.text === '123');
  if (!span) throw new Error('Script fixture has no source span');
  return {
    nominal: span.style.fontSizePt,
    factor: glyphSizeFactorOf(span.style),
    width: span.box.width,
  };
}
