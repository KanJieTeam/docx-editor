import { setLineHangingPunctuationAdvance } from './line-hanging-punctuation.ts';
import { hangingCjkPunctuationAdvance } from './cjk-spacing.ts';
import type { FieldAwarePiece } from './field-projection.ts';
import type { PendingLine } from './pending-line.ts';
import type { StyleSpanRecord, TextMeasurer } from './semantic-records.ts';

/** Retain the actual overflow decision for alignment after placement and cache reuse. */
export function rememberHangingPunctuation(
  line: PendingLine,
  span: StyleSpanRecord,
  piece: FieldAwarePiece,
  measurer: TextMeasurer,
  hangs: boolean,
  available: number
): void {
  const advance =
    hangs && line.width + span.box.width > available + 0.001
      ? hangingCjkPunctuationAdvance(span.text, piece, measurer)
      : undefined;
  if (advance !== undefined || !span.lineEndWhitespace)
    setLineHangingPunctuationAdvance(line, advance);
}
