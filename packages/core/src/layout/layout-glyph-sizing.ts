import type { ResolvedFont } from './font-resource.ts';
import type { ResolvedRunStyle, VerticalAlign } from './run-style.ts';
import type { ShapedRun } from './shaped-run.ts';

/** A qualified glyph size on a layout projection, never on authored run properties. @public */
export interface LayoutGlyphSizing {
  readonly version: 1;
  readonly factor: number;
  readonly fontHash: string;
  readonly faceIndex: number;
  readonly fontRequest: string;
  readonly fontFamily: string | null;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly nominalSizePt: number;
  readonly verticalAlign: VerticalAlign;
  /** The effective shaping inputs at the whole-piece qualification boundary. */
  readonly styleKey: string;
  /** Content and shaping-context identity. Prefixes retain the whole-piece proof. */
  readonly proof: string;
  /** Whole-piece proof boundary. Splits retain this value; independent pieces do not merge. */
  readonly scope: string;
}

/** Read the pinned result without consulting mutable font configuration or global epochs. */
export function layoutGlyphSizingOf(style: ResolvedRunStyle): LayoutGlyphSizing | undefined {
  const value = style.layoutGlyphSizing;
  if (!value) return undefined;
  const valid =
    value.version === 1 &&
    Number.isFinite(value.factor) &&
    value.factor > 0 &&
    value.factor <= 1 &&
    value.nominalSizePt === style.fontSizePt &&
    value.verticalAlign === style.verticalAlign &&
    value.fontFamily === style.fontFamily &&
    value.bold === style.bold &&
    value.italic === style.italic &&
    !style.smallCaps &&
    !style.caps &&
    value.styleKey === scriptGlyphStyleKey(style);
  if (!valid) throw new ScriptGlyphSizingError();
  return value;
}

export function scriptGlyphStyleKey(style: ResolvedRunStyle): string {
  return JSON.stringify([
    style.fontFamily,
    style.fontFamilyEastAsia,
    style.fontSizePt,
    style.verticalAlign,
    style.bold,
    style.italic,
    style.caps,
    style.smallCaps,
    style.complexLane,
    style.shaping?.script,
    style.shaping?.direction,
    style.shaping?.context,
    style.kerningMinPt,
    style.kerningEnabled,
    style.ligatures,
  ]);
}

export function layoutGlyphSizingEqual(a: ResolvedRunStyle, b: ResolvedRunStyle): boolean {
  const left = a.layoutGlyphSizing;
  const right = b.layoutGlyphSizing;
  if (left === right) return true;
  return (
    left !== undefined &&
    right !== undefined &&
    left.version === right.version &&
    left.factor === right.factor &&
    left.fontHash === right.fontHash &&
    left.faceIndex === right.faceIndex &&
    left.fontRequest === right.fontRequest &&
    left.fontFamily === right.fontFamily &&
    left.bold === right.bold &&
    left.italic === right.italic &&
    left.nominalSizePt === right.nominalSizePt &&
    left.verticalAlign === right.verticalAlign &&
    left.styleKey === right.styleKey &&
    left.proof === right.proof &&
    left.scope === right.scope
  );
}

/** Remove layout-only metadata before formatting reads or writes. */
export function scriptGlyphStyleForEditing(style: ResolvedRunStyle): ResolvedRunStyle {
  if (!style.layoutGlyphSizing) return style;
  const { layoutGlyphSizing: _layoutGlyphSizing, ...authored } = style;
  return authored;
}

/** Check the actual glyph run, including complete, contiguous primary-face coverage. */
export function isSingleScriptFace(run: ShapedRun, font: ResolvedFont): boolean {
  if (run.glyphs.length === 0 || run.glyphs.some((glyph) => glyph.id === 0)) return false;
  let end = 0;
  for (const span of run.fontSpans) {
    if (
      span.glyphStart !== end ||
      span.glyphEnd <= end ||
      span.glyphEnd > run.glyphs.length ||
      span.fallbackIndex !== null ||
      span.font.hash !== font.hash ||
      span.font.faceIndex !== font.faceIndex ||
      scriptFontRequestKey(span.font) !== scriptFontRequestKey(font)
    )
      return false;
    end = span.glyphEnd;
  }
  return end === run.glyphs.length;
}

export function scriptFaceMatches(
  style: ResolvedRunStyle,
  font: ResolvedFont,
  run: ShapedRun
): boolean {
  const value = layoutGlyphSizingOf(style);
  return (
    !value ||
    (value.fontHash === font.hash &&
      value.faceIndex === font.faceIndex &&
      value.fontRequest === scriptFontRequestKey(font) &&
      isSingleScriptFace(run, font))
  );
}

export function scriptFontRequestKey(font: ResolvedFont): string {
  return JSON.stringify([font.request.family, font.request.weight, font.request.style]);
}

/** Refuse inconsistent geometry instead of applying another factor to a published span. */
export class ScriptGlyphSizingError extends Error {
  constructor() {
    super('Qualified script text no longer uses its admitted font face');
    this.name = 'ScriptGlyphSizingError';
  }
}
