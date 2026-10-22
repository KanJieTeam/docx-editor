export function syntheticScriptFont(options: {
  readonly em: number;
  readonly subscriptY: number;
  readonly superscriptY: number;
  readonly omitOs2?: boolean;
}): Uint8Array {
  const count = options.omitOs2 ? 1 : 2;
  const head = 12 + count * 16;
  const os2 = head + 54;
  const bytes = new Uint8Array(os2 + 26);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint16(4, count);
  view.setUint32(12, 0x68656164);
  view.setUint32(20, head);
  view.setUint32(24, 54);
  view.setUint16(head + 18, options.em);
  if (!options.omitOs2) {
    view.setUint32(28, 0x4f532f32);
    view.setUint32(36, os2);
    view.setUint32(40, 26);
    view.setUint16(os2, 3);
    view.setInt16(os2 + 12, options.subscriptY);
    view.setInt16(os2 + 20, options.superscriptY);
  }
  return bytes;
}

export function syntheticScriptCollection(faces: readonly Uint8Array[]): Uint8Array {
  const header = 12 + faces.length * 4;
  const bytes = new Uint8Array(header + faces.reduce((sum, face) => sum + face.length, 0));
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x74746366);
  view.setUint32(4, 0x00010000);
  view.setUint32(8, faces.length);
  let start = header;
  faces.forEach((face, index) => {
    view.setUint32(12 + index * 4, start);
    bytes.set(face, start);
    const count = view.getUint16(start + 4);
    for (let record = 0; record < count; record++) {
      const field = start + 12 + record * 16 + 8;
      view.setUint32(field, view.getUint32(field) + start);
    }
    start += face.length;
  });
  return bytes;
}
