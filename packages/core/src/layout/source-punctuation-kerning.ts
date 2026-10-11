import type { FieldAwarePiece } from './field-pieces.ts';
import type { CjkParagraphTypography } from './cjk-typography.ts';
import type { TextMeasurer } from './semantic-records.ts';
import { isRunKerningEnabled } from './run-kerning.ts';
import { styleForFontSlot } from './script-itemization.ts';
import { formatRevisionOf } from './revision-projection.ts';
import { withLayoutDerivedSpacing } from './layout-derived-spacing.ts';
import { resolvedPunctuationFace } from './resolved-punctuation-face.ts';
import { compressCjkPieces } from './cjk-spacing.ts';

export function resolveCjkPieces(
  pieces: readonly FieldAwarePiece[],
  policy: CjkParagraphTypography,
  measurer: TextMeasurer,
  preserveColonAdvances: boolean
): readonly FieldAwarePiece[] {
  return kernSourcePunctuationPieces(
    compressCjkPieces(pieces, policy, measurer, preserveColonAdvances),
    policy,
    measurer
  );
}

const PAIRS = new Set(['》（', '）（']);
const safe = (p: FieldAwarePiece): boolean =>
  p.end - p.start === p.text.length &&
  !p.projected &&
  p.measureText === undefined &&
  !p.fieldAtom &&
  !p.equation &&
  !p.inlineDrawing &&
  !p.anchoredAtom &&
  !p.positionalTab &&
  !p.breakKind &&
  !p.link &&
  !p.revisions?.length &&
  !p.changeSites?.length &&
  !formatRevisionOf(p.props) &&
  !p.style.textOutline &&
  !p.style.underline &&
  !p.style.strike &&
  !p.style.doubleStrike &&
  !p.style.smallCaps &&
  !p.style.caps &&
  p.style.shaping?.direction !== 'rtl' &&
  !/[\t\r\n\f]/u.test(p.text);

// This is separate from characterSpacingControl. Compression runs first. An
// already shortened seam must not receive a second reduction from source kern.
export function kernSourcePunctuationPieces(
  pieces: readonly FieldAwarePiece[],
  policy: CjkParagraphTypography,
  measurer: TextMeasurer
): readonly FieldAwarePiece[] {
  if (
    policy.settings?.punctuationKerning === false ||
    !pieces.length ||
    pieces.some((p) => !safe(p))
  )
    return pieces;
  const text = pieces.map((p) => p.text).join('');
  if (text.length > 100000) return pieces;
  const starts: number[] = [];
  let cursor = 0;
  for (const p of pieces) {
    starts.push(cursor);
    cursor += p.text.length;
  }
  const reductions = new Map<number, number>();
  let owner = 0;
  for (let i = 0; i + 1 < text.length; i++) {
    if (!PAIRS.has(text.slice(i, i + 2))) continue;
    while (owner + 1 < pieces.length && starts[owner + 1]! <= i) owner++;
    const left = pieces[owner]!;
    const rightOwner = i + 1 < starts[owner]! + left.text.length ? owner : owner + 1;
    const right = pieces[rightOwner];
    if (!right || (rightOwner !== owner && left.end !== right.start)) continue;
    const a = styleForFontSlot(left.style, left.fontSlot);
    const b = styleForFontSlot(right.style, right.fontSlot);
    if (
      !isRunKerningEnabled(a) ||
      !isRunKerningEnabled(b) ||
      a.characterSpacingPt !== 0 ||
      b.characterSpacingPt !== 0 ||
      a.horizontalScalePercent !== 100 ||
      b.horizontalScalePercent !== 100 ||
      a.fontSizePt !== b.fontSizePt ||
      a.bold ||
      b.bold ||
      a.italic ||
      b.italic
    )
      continue;
    const fa = resolvedPunctuationFace(measurer, a, text[i]!);
    const fb = resolvedPunctuationFace(measurer, b, text[i + 1]!);
    if (
      !fa ||
      !fb ||
      fa.hash !== fb.hash ||
      fa.faceIndex !== fb.faceIndex ||
      !['SimSun', '宋体'].includes(fa.family) ||
      !['SimSun', '宋体'].includes(fb.family) ||
      fa.weight !== 400 ||
      fb.weight !== 400 ||
      fa.style !== 'normal' ||
      fb.style !== 'normal'
    )
      continue;
    const leftWidth = measurer.measure(text[i]!, a),
      rightWidth = measurer.measure(text[i + 1]!, b);
    const pairWidth = measurer.measure(text.slice(i, i + 2), a);
    if (
      ![leftWidth, rightWidth, pairWidth].every(Number.isFinite) ||
      Math.abs(leftWidth - a.fontSizePt) > 0.001 ||
      Math.abs(rightWidth - b.fontSizePt) > 0.001 ||
      Math.abs(pairWidth - leftWidth - rightWidth) > 0.001
    )
      continue;
    reductions.set(i, leftWidth / 2);
  }
  if (!reductions.size) return pieces;
  const out: FieldAwarePiece[] = [];
  for (let index = 0; index < pieces.length; index++) {
    const p = pieces[index]!;
    let from = 0;
    for (let local = 0; local < p.text.length; local++) {
      const reduction = reductions.get(starts[index]! + local);
      if (!reduction) continue;
      if (local > from)
        out.push({
          ...p,
          text: p.text.slice(from, local),
          start: p.start + from,
          end: p.start + local,
        });
      out.push({
        ...p,
        text: p.text[local]!,
        start: p.start + local,
        end: p.start + local + 1,
        style: withLayoutDerivedSpacing(p.style, -reduction),
      });
      from = local + 1;
    }
    if (!from) out.push(p);
    else if (from < p.text.length)
      out.push({ ...p, text: p.text.slice(from), start: p.start + from });
  }
  return out;
}
