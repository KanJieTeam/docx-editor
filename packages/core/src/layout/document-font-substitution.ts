import type { FontConfiguration, FontSourceSubstitution } from '../contracts/editor.ts';
import type { HeadlessDocumentView } from '../store/headless-document-view.ts';
import { fontTableAlternates } from '../store/package/font-table-alternates.ts';
import { fontFamilyName } from '../store/package/font-family-name.ts';
import type { OoxmlElement, OoxmlNode } from '../store/package/ooxml-tree.ts';
import { WML_NAMESPACE_URI } from '../store/package/ooxml-shared.ts';
import { eastAsianDefaultFamily } from '../store/package/theme-font-scheme.ts';
import { composeFontConfiguration, MAX_RESOLVER_FAMILIES } from './font-composition.ts';
import {
  fontRequestKey,
  HARD_MAX_FONT_SOURCES,
  FontResolutionError,
  type FontResourceSnapshot,
} from './font-resource.ts';

export interface DocumentFontSubstitutionPlan {
  readonly families: readonly string[];
  readonly alternates: ReadonlyMap<string, string>;
  readonly fallbacks: ReadonlyMap<string, string>;
  readonly ambiguousFamilies: readonly string[];
}

/** Report admitted document substitutions without reporting generic host metric twins. */
export function admittedDocumentSubstitutions(
  configuration: FontConfiguration,
  plan: DocumentFontSubstitutionPlan,
  snapshot: FontResourceSnapshot
): ReadonlySet<string> {
  const families = new Set<string>();
  for (const substitution of configuration.substitutions ?? []) {
    const key = substitution.from.family.trim().toLowerCase();
    const target = substitution.to.family.trim().toLowerCase();
    const documentTarget = [plan.alternates.get(key), plan.fallbacks.get(key)].some(
      (candidate) => candidate?.trim().toLowerCase() === target
    );
    if (
      target !== key &&
      documentTarget &&
      !(snapshot.resolve(substitution.from) instanceof FontResolutionError)
    ) {
      families.add(substitution.from.family.toLowerCase());
    }
  }
  return families;
}

function wordAttribute(node: OoxmlElement, name: string): string | undefined {
  return node.attributes.find(
    (attribute) => attribute.localName === name && attribute.namespaceUri === WML_NAMESPACE_URI
  )?.value;
}

function child(node: OoxmlElement | null | undefined, name: string): OoxmlElement | undefined {
  for (const candidate of (node?.children ?? []) as readonly OoxmlNode[]) {
    if (
      candidate.kind !== 'textValue' &&
      candidate.namespaceUri === WML_NAMESPACE_URI &&
      candidate.localName === name
    )
      return candidate;
  }
  return undefined;
}

/** Avoid assigning one global substitute when a whole name needs different script defaults. */
function scriptDefaults(
  view: Pick<HeadlessDocumentView, 'currentPackage' | 'stylesRoot' | 'documentThemeFonts'>,
  defaultFamily: string
): {
  fallbacks: ReadonlyMap<string, string>;
  ambiguousFamilies: readonly string[];
  defaults: readonly string[];
} {
  const defaultsProperties = child(
    child(child(view.stylesRoot(), 'docDefaults'), 'rPrDefault'),
    'rPr'
  );
  const defaultFonts = child(defaultsProperties, 'rFonts');
  const language = child(defaultsProperties, 'lang');
  const theme = view.documentThemeFonts();
  const latin =
    fontFamilyName(
      defaultFonts && (wordAttribute(defaultFonts, 'ascii') ?? wordAttribute(defaultFonts, 'hAnsi'))
    ) ??
    theme.minor ??
    defaultFamily;
  const eastAsia =
    fontFamilyName(defaultFonts && wordAttribute(defaultFonts, 'eastAsia')) ??
    theme.minorEastAsia ??
    eastAsianDefaultFamily(language && wordAttribute(language, 'eastAsia'));
  const complex =
    fontFamilyName(defaultFonts && wordAttribute(defaultFonts, 'cs')) ??
    theme.minorBidi ??
    defaultFamily;
  const needs = new Map<string, Set<string>>();
  let inspected = 0;
  for (const part of view.currentPackage().parts.values()) {
    const stack: OoxmlNode[] = [part.root];
    while (stack.length > 0 && inspected++ < 100_000) {
      const node = stack.pop()!;
      if (node.kind === 'textValue') continue;
      if (node.namespaceUri === WML_NAMESPACE_URI && node.localName === 'rFonts') {
        for (const [slot, target] of [
          ['ascii', latin],
          ['hAnsi', latin],
          ['eastAsia', eastAsia],
          ['cs', complex],
        ] as const) {
          const family = fontFamilyName(wordAttribute(node, slot));
          if (!family || !target) continue;
          const key = family.trim().toLowerCase();
          const targets = needs.get(key) ?? new Set<string>();
          targets.add(target);
          needs.set(key, targets);
        }
      }
      for (const next of node.children) stack.push(next);
    }
    // An incomplete scan cannot prove that a family has only one script target.
    if (stack.length > 0) return { fallbacks: new Map(), ambiguousFamilies: [], defaults: [] };
  }
  const fallbacks = new Map<string, string>();
  const ambiguousFamilies: string[] = [];
  for (const [family, targets] of needs) {
    if (targets.size === 1) fallbacks.set(family, [...targets][0]!);
    else ambiguousFamilies.push(family);
  }
  return {
    fallbacks,
    ambiguousFamilies,
    defaults: [
      ...new Set([latin, eastAsia, complex].filter((family): family is string => family !== null)),
    ],
  };
}

/** Request complete table aliases without interpreting punctuation as a list. */
export function documentFontSubstitutionPlan(
  view: Pick<HeadlessDocumentView, 'currentPackage' | 'stylesRoot' | 'documentThemeFonts'>,
  families: readonly string[],
  defaults: readonly string[] = [],
  defaultFamily = 'Calibri'
): DocumentFontSubstitutionPlan {
  const alternates = fontTableAlternates(view.currentPackage());
  const scripts = scriptDefaults(view, defaultFamily);
  const targets = families.flatMap((family) => {
    const alternate = alternates.get(family.trim().toLowerCase());
    return alternate ? [alternate] : [];
  });
  const requested = new Map<string, string>();
  // Origins already receive the Latin default through request.defaultFamily.
  const scriptRequests = scripts.defaults.filter(
    (family) => family.trim().toLowerCase() !== defaultFamily.trim().toLowerCase()
  );
  for (const raw of [...families, ...targets, ...defaults, ...scriptRequests]) {
    const family = fontFamilyName(raw);
    if (family && !requested.has(family.trim().toLowerCase()))
      requested.set(family.trim().toLowerCase(), family);
    if (requested.size === MAX_RESOLVER_FAMILIES) break;
  }
  return {
    families: [...requested.values()],
    alternates,
    fallbacks: scripts.fallbacks,
    ambiguousFamilies: scripts.ambiguousFamilies,
  };
}

/** A direct face wins. A whole table alias precedes a host's generic substitute. */
export function applyDocumentFontSubstitutions(
  configuration: FontConfiguration,
  plan: DocumentFontSubstitutionPlan,
  fallbackFamilies: ReadonlyMap<string, string> = plan.fallbacks
): FontConfiguration {
  const sources = new Map(
    configuration.sources.map((source) => [fontRequestKey(source.request), source.request])
  );
  const substitutions: FontSourceSubstitution[] = [];
  const existing = new Set(
    (configuration.substitutions ?? []).map((substitution) => fontRequestKey(substitution.from))
  );
  let remaining = Math.max(0, HARD_MAX_FONT_SOURCES - existing.size);
  for (const family of plan.families) {
    const key = family.trim().toLowerCase();
    const alternate = plan.alternates.get(key);
    const fallback = fallbackFamilies.get(key);
    for (const weight of [400, 700]) {
      for (const style of ['normal', 'italic'] as const) {
        const from = { family, weight, style };
        if (sources.has(fontRequestKey(from))) continue;
        for (const candidate of [alternate, fallback]) {
          if (!candidate || candidate.trim().toLowerCase() === key) continue;
          const to = sources.get(fontRequestKey({ family: candidate, weight, style }));
          if (!to) continue;
          if (!existing.has(fontRequestKey(from))) {
            if (remaining === 0) break;
            remaining--;
          }
          substitutions.push({ from, to });
          break;
        }
      }
    }
  }
  return composeFontConfiguration({
    ...configuration,
    substitutions: [...substitutions, ...(configuration.substitutions ?? [])],
  });
}
