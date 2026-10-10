import { fontFamilyAlternatives } from '../store/package/font-family-reference.ts';

/** Keep supplied faces before platform names, including glyph fallback alternatives. */
export function fontFamilyStack(
  reference: string | null | undefined,
  fontAlias?: (family: string) => string | undefined
): string | null {
  const names = fontFamilyAlternatives(reference);
  if (names.length === 0) return null;
  const aliases = new Set<string>();
  for (const family of names.length > 1 ? [reference!, ...names] : [reference!]) {
    const alias = fontAlias?.(family);
    if (!alias || alias.includes(';')) continue;
    const parsed = fontFamilyAlternatives(alias);
    if (parsed.length === 1) aliases.add(parsed[0]!);
  }
  return [...aliases, ...names].map((name) => `"${name}"`).join(', ');
}
