// The document setting affects punctuation pairs. Letter pairs retain run kerning.
import type { FieldAwarePiece } from './field-pieces.ts';
import { isRunKerningEnabled } from './run-kerning.ts';
import type { ResolvedRunStyle } from './run-style.ts';
import { segmentGraphemes } from './grapheme.ts';

const markedStyles = new WeakMap<ResolvedRunStyle, ResolvedRunStyle>();

/** Apply document policy to layout styles without changing canonical run properties. */
export function withPunctuationKerningPolicy(
  pieces: readonly FieldAwarePiece[],
  policy: boolean | { readonly noPunctuationKerning?: boolean } | undefined
): readonly FieldAwarePiece[] {
  const disabled = typeof policy === 'object' ? policy.noPunctuationKerning : policy;
  if (!disabled) return pieces;
  return pieces.map((piece) => {
    if (!isRunKerningEnabled(piece.style)) return piece;
    let style = markedStyles.get(piece.style);
    if (!style) {
      style = {
        ...piece.style,
        shaping: {
          ...(piece.style.shaping ?? { script: 'Zyyy', direction: 'ltr', level: 0, baseLevel: 0 }),
          noPunctuationKerning: true,
        },
      };
      markedStyles.set(piece.style, style);
    }
    return { ...piece, style };
  });
}

/** HarfBuzz clusters use UTF-16 offsets, including the leading shaping context. */
export function punctuationKerningRanges(
  text: string,
  contextBeforeLength = 0
): readonly { readonly start: number; readonly end: number }[] {
  return segmentGraphemes(text)
    .filter((cluster) => /^\p{P}/u.test(cluster.text))
    .map((cluster) => ({
      start: contextBeforeLength + cluster.utf16From,
      end: contextBeforeLength + cluster.utf16To,
    }));
}

/** Native Canvas and DOM consumers retain kerning within each non-punctuation segment. */
export function punctuationKerningSegments(text: string): readonly {
  readonly text: string;
  readonly punctuation: boolean;
}[] {
  const segments: { text: string; punctuation: boolean }[] = [];
  for (const cluster of segmentGraphemes(text)) {
    const punctuation = /^\p{P}/u.test(cluster.text);
    const last = segments.at(-1);
    if (last?.punctuation === punctuation) last.text += cluster.text;
    else segments.push({ text: cluster.text, punctuation });
  }
  return segments;
}
