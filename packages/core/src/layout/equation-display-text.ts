// OMML stores ordinary letters and a separate mathematical style. Transform only
// the display string, so canonical text, editing offsets, and saves stay unchanged.
import type { EquationExpression } from '@docx-editor.dev/core/store';

type TextExpression = Extract<EquationExpression, { readonly kind: 'text' }>;
type MathStyle = 'p' | 'b' | 'i' | 'bi';

const GREEK_SYMBOLS = new Map([
  [0x3f4, 17],
  [0x2207, 25],
  [0x2202, 51],
  [0x3f5, 52],
  [0x3d1, 53],
  [0x3f0, 54],
  [0x3d5, 55],
  [0x3f1, 56],
  [0x3d6, 57],
]);

function styledCodePoint(code: number, style: Exclude<MathStyle, 'p'>): number {
  const latin = style === 'b' ? 0x1d400 : style === 'i' ? 0x1d434 : 0x1d468;
  if (code >= 0x41 && code <= 0x5a) return latin + code - 0x41;
  if (code >= 0x61 && code <= 0x7a) {
    // Mathematical italic small h occupies the older letterlike-symbol slot.
    if (style === 'i' && code === 0x68) return 0x210e;
    return latin + 26 + code - 0x61;
  }
  if (style === 'i' && code === 0x131) return 0x1d6a4;
  if (style === 'i' && code === 0x237) return 0x1d6a5;
  if (style !== 'i' && code >= 0x30 && code <= 0x39) return 0x1d7ce + code - 0x30;
  const greek = style === 'b' ? 0x1d6a8 : style === 'i' ? 0x1d6e2 : 0x1d71c;
  if (code >= 0x391 && code <= 0x3a9 && code !== 0x3a2) return greek + code - 0x391;
  if (code >= 0x3b1 && code <= 0x3c9) return greek + 26 + code - 0x3b1;
  const symbol = GREEK_SYMBOLS.get(code);
  return symbol === undefined ? code : greek + symbol;
}

/** Internal display derivation; it does not change the source expression. */
export function equationDisplayText(
  expression: TextExpression,
  fontFamily: string | null
): {
  readonly text: string;
  readonly mathStyle: MathStyle | null;
  readonly mathematicalGlyphs: boolean;
} {
  const format = expression.format;
  if (format?.normalText || fontFamily !== 'Cambria Math')
    return { text: expression.value, mathStyle: null, mathematicalGlyphs: false };
  // m:sty defaults to i. Word b/i properties remain separate and apply to
  // normal-text runs; they do not replace this mathematical style default.
  const mathStyle = format?.mathStyle ?? 'i';
  // Explicit non-math faces keep their existing Word formatting. This bounded
  // variant mapping applies only to the engine's default math face.
  const mathematicalGlyphs = true;
  const text =
    mathStyle === 'p' || !mathematicalGlyphs
      ? expression.value
      : Array.from(expression.value, (character) =>
          String.fromCodePoint(styledCodePoint(character.codePointAt(0)!, mathStyle))
        ).join('');
  return { text, mathStyle, mathematicalGlyphs };
}
