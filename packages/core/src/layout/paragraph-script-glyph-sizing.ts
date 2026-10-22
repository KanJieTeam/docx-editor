import type { FieldAwarePiece } from './field-pieces.ts';
import type { TextMeasurer } from './text-measurer.ts';
import { displayText } from './run-style.ts';
import { styleForFontSlot } from './script-itemization.ts';

/** Qualify each complete ordinary source piece before continuation and line splitting. */
export function projectScriptGlyphSizing<T extends FieldAwarePiece>(
  pieces: readonly T[],
  paragraphId: string,
  measurer: TextMeasurer
): readonly T[] {
  if (!measurer.scriptGlyphSizing) return pieces;
  let result: T[] | undefined;
  for (let index = 0; index < pieces.length; index++) {
    const piece = pieces[index]!;
    if (
      piece.style.verticalAlign === 'baseline' ||
      piece.style.caps ||
      piece.style.smallCaps ||
      !piece.text ||
      piece.projected ||
      !Number.isSafeInteger(piece.start) ||
      !Number.isSafeInteger(piece.end) ||
      piece.start < 0 ||
      piece.end - piece.start !== piece.text.length ||
      piece.measureText !== undefined ||
      piece.noteNav ||
      piece.noteSeparator ||
      piece.fieldAtom ||
      piece.inlineDrawing ||
      piece.anchoredAtom ||
      piece.equation ||
      piece.unsupported ||
      piece.positionalTab ||
      /[\t\n\r\f\u00ad]/u.test(piece.text)
    )
      continue;
    const faceStyle = styleForFontSlot(piece.style, piece.fontSlot);
    const value = measurer.scriptGlyphSizing(
      displayText(piece.text, faceStyle),
      faceStyle,
      JSON.stringify([paragraphId, piece.start, piece.end])
    );
    if (!value) continue;
    result ??= pieces.slice();
    result[index] = {
      ...piece,
      style: Object.freeze({ ...piece.style, layoutGlyphSizing: value }),
    };
  }
  return result ?? pieces;
}
