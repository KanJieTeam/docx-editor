import { readFileSync } from 'node:fs';

/** Public DejaVu fixture in a structural TTC envelope. The signature is not authenticated. */
export function renderOnlyCollectionFontFixture(): Uint8Array {
  const face = new Uint8Array(
    readFileSync(new URL('../../layout/__tests__/fixtures/fonts/DejaVuSans.ttf', import.meta.url))
  );
  const stride = (face.length + 3) & ~3;
  const signatureOffset = 32 + stride * 2;
  const bytes = new Uint8Array(signatureOffset + 8);
  const view = new DataView(bytes.buffer);
  bytes.set(new TextEncoder().encode('ttcf'));
  view.setUint32(4, 0x00020000);
  view.setUint32(8, 2);
  view.setUint32(20, 0x44534947);
  view.setUint32(24, 8);
  view.setUint32(28, signatureOffset);
  view.setUint32(signatureOffset, 1);
  for (let index = 0; index < 2; index++) {
    const base = 32 + index * stride;
    view.setUint32(12 + index * 4, base);
    bytes.set(face, base);
    for (let table = 0; table < view.getUint16(base + 4); table++) {
      const entry = base + 12 + table * 16;
      view.setUint32(entry + 8, base + view.getUint32(entry + 8));
    }
  }
  return bytes;
}
