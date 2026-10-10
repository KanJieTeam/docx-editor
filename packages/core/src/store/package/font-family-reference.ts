// File-derived family references must not become CSS syntax.
const FAMILY_NAME = /^[\p{L}\p{N}\p{M} \-.+_]{1,64}$/u;
const NO_FAMILIES: readonly string[] = Object.freeze([]);

/** Parse bounded alternatives. Reject the whole reference if any name is unsafe. */
export function fontFamilyAlternatives(raw: string | null | undefined): readonly string[] {
  if (!raw || raw.length > 128 || [...raw].length > 64) return NO_FAMILIES;
  const names = raw.split(';');
  if (names.length > 8) return NO_FAMILIES;
  if (names.some((name) => !FAMILY_NAME.test(name))) return NO_FAMILIES;
  const families = names.map((name) => name.trim());
  if (families.some((name) => !FAMILY_NAME.test(name))) return NO_FAMILIES;
  return families;
}

/** Keep the authored reference. Parsing does not rewrite canonical XML. */
export function validFontFamilyReference(raw: string | undefined): string | null {
  return fontFamilyAlternatives(raw).length > 0 ? raw! : null;
}
