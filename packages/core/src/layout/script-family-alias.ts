import { trustedFontBytes, type ResolvedFont } from './font-resource.ts';
import type { ResolvedRunStyle } from './run-style.ts';
import { sfntTables } from './sfnt-script-tables.ts';

const familyKey = (value: string) => value.trim().toLowerCase();
const MAX_NAME_BYTES = 1024 * 1024;
const MAX_NAME_RECORDS = 256;
const MAX_FAMILY_BYTES = 512;

function unicodeFamily(view: DataView, offset: number, length: number): string | undefined {
  if (length === 0 || length > MAX_FAMILY_BYTES || length % 2 !== 0) return undefined;
  let value = '';
  for (let index = 0; index < length; index += 2) {
    const code = view.getUint16(offset + index);
    if (code >= 0xd800 && code <= 0xdbff) {
      if (index + 2 >= length) return undefined;
      const low = view.getUint16(offset + index + 2);
      if (low < 0xdc00 || low > 0xdfff) return undefined;
      value += String.fromCharCode(code, low);
      index += 2;
    } else {
      if ((code >= 0xdc00 && code <= 0xdfff) || code < 0x20 || code === 0x7f) return undefined;
      value += String.fromCharCode(code);
    }
  }
  const key = familyKey(value);
  return key.length > 0 ? key : undefined;
}

// A translated family name must occur on the actual admitted face, not in a host alias list.
function sameFaceFamilyNames(font: ResolvedFont): ReadonlySet<string> | undefined {
  const tables = sfntTables(trustedFontBytes(font), font.faceIndex);
  const name = tables?.tableBounds(0x6e616d65, 6);
  if (!tables || !name || name.length > MAX_NAME_BYTES) return undefined;
  const view = tables.view;
  const format = view.getUint16(name.offset);
  const count = view.getUint16(name.offset + 2);
  const storage = view.getUint16(name.offset + 4);
  // Format 1 language tags and non-Unicode family records remain outside this subset.
  if (
    format !== 0 ||
    count === 0 ||
    count > MAX_NAME_RECORDS ||
    storage < 6 + count * 12 ||
    storage > name.length
  )
    return undefined;
  const names = new Set<string>();
  const records = new Map<string, string>();
  for (let index = 0; index < count; index++) {
    const record = name.offset + 6 + index * 12;
    const platform = view.getUint16(record),
      encoding = view.getUint16(record + 2);
    const language = view.getUint16(record + 4),
      nameId = view.getUint16(record + 6);
    const length = view.getUint16(record + 8),
      relative = view.getUint16(record + 10);
    if (storage + relative + length > name.length) return undefined;
    if (nameId !== 1 && nameId !== 16) continue;
    if (
      !(
        (platform === 0 && encoding <= 6) ||
        (platform === 3 && (encoding === 1 || encoding === 10))
      )
    )
      continue;
    const value = unicodeFamily(view, name.offset + storage + relative, length);
    if (!value) return undefined;
    const key = `${platform}:${encoding}:${language}:${nameId}`;
    const previous = records.get(key);
    if (previous !== undefined && previous !== value) return undefined;
    records.set(key, value);
    names.add(value);
  }
  return names;
}

/** Private per-measurer qualification. This does not alter font registration or substitution. */
export function createScriptFamilyAliasQualifier(): (
  font: ResolvedFont,
  style: ResolvedRunStyle
) => boolean {
  const qualified = new WeakMap<ResolvedFont, boolean>();
  return (font, style) => {
    const substitution = font.substitution;
    if (!substitution || substitution.lineMetrics !== undefined || !style.fontFamily) return false;
    const { requested, resolved } = substitution;
    if (
      requested.weight !== resolved.weight ||
      requested.style !== resolved.style ||
      resolved.weight !== font.request.weight ||
      resolved.style !== font.request.style ||
      familyKey(resolved.family) !== familyKey(font.request.family) ||
      familyKey(requested.family) !== familyKey(style.fontFamily) ||
      requested.weight !== (style.bold ? 700 : 400) ||
      requested.style !== (style.italic ? 'italic' : 'normal')
    )
      return false;
    const cached = qualified.get(font);
    if (cached !== undefined) return cached;
    const names = sameFaceFamilyNames(font);
    const result =
      names?.has(familyKey(requested.family)) === true && names.has(familyKey(resolved.family));
    qualified.set(font, result);
    return result;
  };
}
