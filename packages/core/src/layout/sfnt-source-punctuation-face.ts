import { trustedFontBytes, type ResolvedFont } from './font-resource.ts';

const families = new Set(['SimSun', '宋体']);
const known = new WeakMap<ResolvedFont, boolean>();

// A resource request can name any font. Check the selected binary name table,
// not the authored family or a host alias. No font bytes leave this module.
export function hasSourcePunctuationFace(font: ResolvedFont): boolean {
  const cached = known.get(font);
  if (cached !== undefined) return cached;
  let result = false;
  try {
    const bytes = trustedFontBytes(font);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const within = (at: number, count: number): boolean =>
      Number.isSafeInteger(at) && at >= 0 && count >= 0 && at <= bytes.length - count;
    if (!within(0, 12)) return false;
    let offset = 0;
    if (view.getUint32(0) === 0x74746366) {
      const count = view.getUint32(8);
      if (font.faceIndex < 0 || font.faceIndex >= count || !within(12 + font.faceIndex * 4, 4))
        return false;
      offset = view.getUint32(12 + font.faceIndex * 4);
    } else if (font.faceIndex !== 0) return false;
    if (!within(offset, 12)) return false;
    const tables = view.getUint16(offset + 4);
    if (!within(offset + 12, tables * 16)) return false;
    let nameOffset = 0,
      nameLength = 0;
    for (let i = 0; i < tables; i++) {
      const at = offset + 12 + i * 16;
      if (view.getUint32(at) !== 0x6e616d65) continue;
      if (nameLength) return false;
      nameOffset = view.getUint32(at + 8);
      nameLength = view.getUint32(at + 12);
    }
    if (!nameLength || !within(nameOffset, nameLength) || nameLength < 6) return false;
    const count = view.getUint16(nameOffset + 2),
      storage = view.getUint16(nameOffset + 4);
    if (
      count > 16384 ||
      6 + count * 12 > nameLength ||
      storage < 6 + count * 12 ||
      storage > nameLength
    )
      return false;
    let matched = false;
    for (let i = 0; i < count; i++) {
      const at = nameOffset + 6 + i * 12,
        platform = view.getUint16(at);
      const nameId = view.getUint16(at + 6),
        length = view.getUint16(at + 8),
        start = view.getUint16(at + 10);
      if (![1, 16].includes(nameId) || ![0, 3].includes(platform)) continue;
      if (!length || length % 2 || length > 256 || storage + start + length > nameLength)
        return false;
      let name = '';
      for (let j = 0; j < length; j += 2)
        name += String.fromCharCode(view.getUint16(nameOffset + storage + start + j));
      if (families.has(name)) matched = true;
    }
    result = matched;
  } catch {
    result = false;
  } finally {
    known.set(font, result);
  }
  return result;
}
