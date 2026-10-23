import type { StyleSpanRecord, TextMeasurer } from './semantic-records.ts';
import { measureDisplayText } from './run-style.ts';
import { styleForFontSlot } from './script-itemization.ts';
import { layoutDerivedSpacingStyleForEditing } from './layout-derived-spacing.ts';

type PairSide = Pick<StyleSpanRecord, 'text' | 'style' | 'fontSlot' | 'glyphOffsetPt'>;

/** Measured blank sides use the opening face size. These are total authored budgets. */
export function closeOpenBearingBudgets(
  closing: PairSide,
  opening: PairSide,
  measurer: TextMeasurer
): { closing: number; opening: number } | undefined {
  if (!measurer.inkBounds) return undefined;
  const ea = (side: PairSide) =>
    side.fontSlot === 'eastAsia' ||
    (side.fontSlot === undefined &&
      !!side.style.fontFamilyEastAsia &&
      side.style.fontFamilyEastAsia === side.style.fontFamily);
  if (!ea(closing) || !ea(opening)) return undefined;
  const a = layoutDerivedSpacingStyleForEditing(styleForFontSlot(closing.style, closing.fontSlot));
  const b = layoutDerivedSpacingStyleForEditing(styleForFontSlot(opening.style, opening.fontSlot));
  if (
    a.characterSpacingPt !== 0 ||
    b.characterSpacingPt !== 0 ||
    a.horizontalScalePercent !== 100 ||
    b.horizontalScalePercent !== 100 ||
    a.shaping !== undefined ||
    b.shaping !== undefined ||
    a.fontFamily !== b.fontFamily ||
    a.bold !== b.bold ||
    a.italic !== b.italic ||
    a.verticalAlign !== 'baseline' ||
    b.verticalAlign !== 'baseline' ||
    a.baselineShiftPt !== 0 ||
    b.baselineShiftPt !== 0 ||
    !Number.isFinite(a.fontSizePt) ||
    a.fontSizePt <= 0 ||
    !Number.isFinite(b.fontSizePt) ||
    b.fontSizePt <= 0 ||
    measurer.hasResolvedFont?.(a) !== true ||
    measurer.hasResolvedFont?.(b) !== true
  )
    return undefined;
  // A larger closing run must not select a larger budget than the opening run.
  const atOpeningSize = { ...a, fontSizePt: b.fontSizePt };
  const closeAdvance = measureDisplayText(closing.text, atOpeningSize, measurer);
  const openAdvance = measureDisplayText(opening.text, b, measurer);
  const closeInk = measurer.inkBounds(closing.text, atOpeningSize);
  const openInk = measurer.inkBounds(opening.text, b);
  if (
    !closeInk ||
    !openInk ||
    ![closeAdvance, openAdvance, closeInk.left, closeInk.right, openInk.left, openInk.right].every(
      Number.isFinite
    ) ||
    closeAdvance <= 0 ||
    openAdvance <= 0 ||
    closeInk.left < 0 ||
    openInk.left < 0 ||
    closeInk.right < closeInk.left ||
    openInk.right < openInk.left ||
    closeInk.right > closeAdvance ||
    openInk.right > openAdvance
  )
    return undefined;
  return {
    closing: Math.max(0, Math.min(openAdvance / 2, closeAdvance - closeInk.right)),
    opening: Math.max(0, Math.min(openAdvance / 2, openInk.left)),
  };
}
