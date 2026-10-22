import { sfntTables } from './sfnt-script-tables.ts';

export interface ScriptSizeEm {
  readonly subscript: number | null;
  readonly superscript: number | null;
}

/** Uniform glyph-size recommendations only. No placement or X/Y distortion. */
export function sfntScriptSizeEm(bytes: Uint8Array, faceIndex: number): ScriptSizeEm | null {
  const tables = sfntTables(bytes, faceIndex);
  if (!tables) return null;
  const head = tables.table(0x68656164, 54);
  const os2 = tables.table(0x4f532f32, 26);
  if (head === null || os2 === null) return null;
  const version = tables.view.getUint16(os2);
  const em = tables.view.getUint16(head + 18);
  // OpenType head range. Unknown OS/2 versions keep the existing glyph policy.
  if (em < 16 || em > 16384 || version > 5) return null;
  const size = (offset: number): number | null => {
    const ratio = tables.view.getInt16(os2 + offset) / em;
    // Opt-in supports shrinking scripts only. Refuse rather than clamp.
    return ratio > 0 && ratio <= 1 ? ratio : null;
  };
  return Object.freeze({ subscript: size(12), superscript: size(20) });
}
