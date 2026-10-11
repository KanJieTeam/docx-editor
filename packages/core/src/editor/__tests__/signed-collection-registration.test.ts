import { test } from 'bun:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import { registerEmbeddedFontFaces, type FontFaceLike } from '../embedded-font-faces.ts';
import { materializeCollectionFace } from '../font-collection-face.ts';
import { boundedStructuralFontValidator } from '../../layout/font-resource.ts';

const fixtures = new URL('../../layout/__tests__/fixtures/fonts/', import.meta.url);
const fonts = ['DejaVuSans.ttf', 'DejaVuSans-Bold.ttf'].map(
  (file) => new Uint8Array(fs.readFileSync(new URL(file, fixtures)))
);
const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
function collection(signature: boolean): Uint8Array {
  const count = fonts.length;
  const headerSize = 12 + count * 4 + 12;
  const offsets = [headerSize, headerSize + fonts[0]!.length];
  const end = headerSize + fonts.reduce((size, font) => size + font.length, 0);
  const result = new Uint8Array(end + (signature ? 8 : 0));
  const view = new DataView(result.buffer);
  result.set(new TextEncoder().encode('ttcf'));
  view.setUint32(4, 0x00020000);
  view.setUint32(8, count);
  fonts.forEach((font, index) => {
    const offset = offsets[index]!;
    view.setUint32(12 + index * 4, offset);
    result.set(font, offset);
    const original = new DataView(font.buffer, font.byteOffset, font.length);
    for (let table = 0; table < original.getUint16(4); table++) {
      const entry = 12 + table * 16;
      view.setUint32(offset + entry + 8, original.getUint32(entry + 8) + offset);
    }
  });
  if (signature) {
    view.setUint32(12 + count * 4, 0x44534947);
    view.setUint32(16 + count * 4, 8);
    view.setUint32(20 + count * 4, end);
    // A public DSIG envelope tests preservation. It has no cryptographic signature.
    view.setUint32(end, 1);
  }
  return result;
}
async function register(bytes: Uint8Array, faceIndex: number, failsLoading = false) {
  const untouched = hash(bytes);
  assert.deepEqual(boundedStructuralFontValidator(bytes, faceIndex), { valid: true });
  const created: { family: string; bytes: Uint8Array; face: FontFaceLike }[] = [];
  const added: FontFaceLike[] = [],
    deleted: FontFaceLike[] = [];
  const registration = await registerEmbeddedFontFaces(
    [
      {
        request: { family: 'Public Collection', weight: 400, style: 'normal' },
        id: 'public-collection',
        bytes,
        hash: `sha256:${untouched}`,
        faceIndex,
      },
    ],
    {
      fontSet: { add: (face) => added.push(face), delete: (face) => deleted.push(face) },
      createFontFace: (family, source) => {
        const face = {
          load: async () => {
            if (failsLoading) throw new Error('Browser refusal');
          },
        };
        created.push({ family, bytes: source, face });
        return face;
      },
    }
  );
  assert.equal(hash(bytes), untouched);
  return { registration, created, added, deleted };
}

test('first signed collection face preserves the complete admitted bytes and private alias', async () => {
  const bytes = collection(true);
  assert.throws(() => materializeCollectionFace(bytes), /signature-aware/);
  const result = await register(bytes, 0);
  assert.equal(result.registration.installed, 1);
  assert.equal(result.created[0]!.bytes, bytes);
  assert.equal(result.created[0]!.family, result.registration.alias('Public Collection'));
  assert.match(result.created[0]!.family, /^docx-embedded-/);
  result.registration.dispose();
  result.registration.dispose();
  assert.deepEqual(result.deleted, result.added);
});

test('nonzero signed face retains the materialization refusal and never earns an alias', async () => {
  const result = await register(collection(true), 1);
  assert.equal(result.registration.installed, 0);
  assert.equal(result.registration.alias('Public Collection'), undefined);
  assert.deepEqual(result.created, []);
});

test('unsigned first face keeps all bytes while its next face uses the selected SFNT', async () => {
  const bytes = collection(false);
  const first = await register(bytes, 0),
    next = await register(bytes, 1);
  assert.equal(first.registration.installed, 1);
  assert.equal(next.registration.installed, 1);
  assert.equal(first.created[0]!.bytes, bytes);
  assert.deepEqual(next.created[0]!.bytes, materializeCollectionFace(bytes, 1));
  assert.notEqual(
    first.registration.alias('Public Collection'),
    next.registration.alias('Public Collection')
  );
  first.registration.dispose();
  next.registration.dispose();
});

test('a browser that refuses a signed collection does not advertise or leak its face', async () => {
  const result = await register(collection(true), 0, true);
  assert.equal(result.registration.installed, 0);
  assert.equal(result.created.length, 1);
  assert.equal(result.registration.alias('Public Collection'), undefined);
  assert.deepEqual(result.added, []);
  result.registration.dispose();
  assert.deepEqual(result.deleted, []);
});

test('an invalid first collection directory never reaches the browser loader', async () => {
  const bytes = collection(true);
  new DataView(bytes.buffer).setUint32(12, bytes.length);
  let calls = 0;
  const result = await registerEmbeddedFontFaces(
    [
      {
        request: { family: 'Invalid Collection', weight: 400, style: 'normal' },
        id: 'invalid',
        bytes,
        hash: `sha256:${hash(bytes)}`,
        faceIndex: 0,
      },
    ],
    {
      fontSet: { add: () => assert.fail('Invalid face added'), delete: () => {} },
      createFontFace: () => {
        calls++;
        return { load: async () => {} };
      },
    }
  );
  assert.equal(calls, 0);
  assert.equal(result.installed, 0);
});
