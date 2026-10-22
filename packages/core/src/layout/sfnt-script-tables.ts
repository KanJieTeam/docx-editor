// Bounded table metadata lookup. This does not admit fonts or change strikeout parsing.
const TTC_TAG = 0x74746366;
const MAX_TABLE_RECORDS = 4096;
const SFNT_SIGNATURES = new Set([0x00010000, 0x4f54544f, 0x74727565]);

export interface SfntTables {
  readonly view: DataView;
  table(tag: number, minimumLength: number): number | null;
}

export function sfntTables(bytes: Uint8Array, faceIndex: number): SfntTables | null {
  if (!Number.isSafeInteger(faceIndex) || faceIndex < 0 || bytes.byteLength < 12) return null;
  const length = bytes.byteLength;
  const view = new DataView(bytes.buffer, bytes.byteOffset, length);
  let base = 0;
  if (view.getUint32(0) === TTC_TAG) {
    const count = view.getUint32(8);
    if (count === 0 || faceIndex >= count || 12 + count * 4 > length) return null;
    base = view.getUint32(12 + faceIndex * 4);
  } else if (faceIndex !== 0) {
    return null;
  }
  if (base + 12 > length || !SFNT_SIGNATURES.has(view.getUint32(base))) return null;
  const count = view.getUint16(base + 4);
  if (count > MAX_TABLE_RECORDS || base + 12 + count * 16 > length) return null;
  return Object.freeze({
    view,
    table(tag: number, minimumLength: number): number | null {
      let found: number | null = null;
      for (let index = 0; index < count; index++) {
        const record = base + 12 + index * 16;
        if (view.getUint32(record) !== tag) continue;
        // Ambiguous table records cannot define one trusted metric.
        if (found !== null) return null;
        const offset = view.getUint32(record + 8);
        const tableLength = view.getUint32(record + 12);
        if (tableLength < minimumLength || offset + tableLength > length) return null;
        found = offset;
      }
      return found;
    },
  });
}
