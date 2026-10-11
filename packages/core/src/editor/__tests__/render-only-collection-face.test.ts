import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { materializeCollectionFace, sfntChecksum } from '../font-collection-face.ts';
import { materializeRenderOnlyCollectionFace } from '../render-only-collection-face.ts';
import { registerEmbeddedFontFaces } from '../embedded-font-faces.ts';

// These tiny public structural fixtures contain no real font or private document.
// They test table transport and host policy, not actual glyph rendering or signatures.
const policy = { allowUnsignedDerivedFaces: true, requireSignedFonts: false };
const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
function collection(individualSignature = false, collectionSignature = true): Uint8Array {
  const tags = individualSignature ? ['DSIG', 'GPOS', 'MATH', 'head'] : ['GPOS', 'MATH', 'head'];
  const faceBytes =
    12 + tags.length * 16 + tags.reduce((n, tag) => n + (tag === 'head' ? 56 : 12), 0);
  const header = 32,
    signatureOffset = header + faceBytes * 2;
  const bytes = new Uint8Array(signatureOffset + 8),
    view = new DataView(bytes.buffer);
  bytes.set(new TextEncoder().encode('ttcf'));
  view.setUint32(4, 0x00020000);
  view.setUint32(8, 2);
  view.setUint32(20, 0x44534947);
  view.setUint32(24, 8);
  view.setUint32(28, signatureOffset);
  view.setUint32(signatureOffset, 1);
  for (let face = 0; face < 2; face++) {
    const base = header + face * faceBytes;
    view.setUint32(12 + face * 4, base);
    view.setUint32(base, 0x00010000);
    view.setUint16(base + 4, tags.length);
    let offset = base + 12 + tags.length * 16;
    tags.forEach((tag, index) => {
      const entry = base + 12 + index * 16,
        length = tag === 'head' ? 54 : 12;
      bytes.set(new TextEncoder().encode(tag), entry);
      view.setUint32(entry + 8, offset);
      view.setUint32(entry + 12, length);
      for (let i = 0; i < length; i++) bytes[offset + i] = (face + 1) * 10 + i;
      if (tag === 'head') view.setUint32(offset + 12, 0x5f0f3cf5);
      offset += (length + 3) & ~3;
    });
  }
  if (!collectionSignature) bytes.subarray(20, 32).fill(0);
  return bytes;
}
function tables(bytes: Uint8Array): Map<string, Uint8Array> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength),
    result = new Map<string, Uint8Array>();
  for (let i = 0; i < view.getUint16(4); i++) {
    const entry = 12 + i * 16,
      tag = String.fromCharCode(...bytes.subarray(entry, entry + 4));
    result.set(
      tag,
      bytes.subarray(
        view.getUint32(entry + 8),
        view.getUint32(entry + 8) + view.getUint32(entry + 12)
      )
    );
  }
  return result;
}

test('explicit unsigned render projection keeps source signature bytes and chosen MATH/GPOS tables', () => {
  const bytes = collection(),
    before = hash(bytes);
  assert.throws(() => materializeCollectionFace(bytes, 1), /signature-aware/);
  const projected = materializeRenderOnlyCollectionFace(bytes, 1, policy);
  assert.equal(hash(bytes), before);
  assert.equal(projected.signature, 'unsigned');
  assert.equal(projected.sourceCollectionSignaturePresent, true);
  assert.equal(projected.sourceSignatureValidated, false);
  assert.equal(projected.faceIndex, 1);
  assert(Object.isFrozen(projected));
  assert.equal(sfntChecksum(projected.bytes), 0xb1b0afba);
  const view = new DataView(bytes.buffer),
    base = view.getUint32(16),
    output = tables(projected.bytes);
  for (const [i, tag] of ['GPOS', 'MATH'].entries()) {
    const entry = base + 12 + i * 16,
      start = view.getUint32(entry + 8),
      length = view.getUint32(entry + 12);
    assert.deepEqual(output.get(tag), bytes.subarray(start, start + length));
  }
  assert.equal(output.has('DSIG'), false);
  assert.throws(() => materializeCollectionFace(bytes, 1), /signature-aware/);
});

test('host absence, unknown policy, disabled policy, and signed-required policy refuse unsigned output', () => {
  const bytes = collection(),
    before = hash(bytes);
  for (const denied of [
    undefined,
    {},
    { allowUnsignedDerivedFaces: false, requireSignedFonts: false },
    { allowUnsignedDerivedFaces: true },
    { allowUnsignedDerivedFaces: true, requireSignedFonts: true },
  ]) {
    assert.throws(
      () => materializeRenderOnlyCollectionFace(bytes, 1, denied as typeof policy),
      /host/
    );
    assert.equal(hash(bytes), before);
  }
});

test('collection signature range, final placement, tag, alignment, and reserved flags stay bounded', () => {
  const changes: ((view: DataView) => void)[] = [
    (v) => v.setUint32(20, 0x42414421),
    (v) => v.setUint32(24, 7),
    (v) => v.setUint32(28, 0xfffffff0),
    (v) => v.setUint32(28, 33),
    (v) => v.setUint32(28, 32),
    (v) => v.setUint32(v.getUint32(28), 2),
    (v) => v.setUint16(v.getUint32(28) + 6, 2),
    (v) => v.setUint16(v.getUint32(28) + 4, 65535),
  ];
  for (const change of changes) {
    const bytes = collection();
    change(new DataView(bytes.buffer));
    const before = hash(bytes);
    assert.throws(() => materializeRenderOnlyCollectionFace(bytes, 1, policy));
    assert.equal(hash(bytes), before);
  }
});

test('signature payload cannot overlap a selected directory or selected MATH table', () => {
  for (const directory of [false, true]) {
    const bytes = collection(),
      view = new DataView(bytes.buffer),
      signature = view.getUint32(28);
    if (directory) view.setUint32(16, signature);
    else view.setUint32(view.getUint32(16) + 12 + 16 + 8, signature);
    const before = hash(bytes);
    assert.throws(() => materializeRenderOnlyCollectionFace(bytes, 1, policy));
    assert.equal(hash(bytes), before);
  }
});

test('an individually signed selected face retains the old refusal even with host opt-in', () => {
  const bytes = collection(true),
    before = hash(bytes);
  assert.throws(() => materializeRenderOnlyCollectionFace(bytes, 1, policy), /Signed faces/);
  assert.equal(hash(bytes), before);
});

test('non-collection bytes and invalid face indexes cannot take the opt-in path', () => {
  const bytes = collection(),
    before = hash(bytes);
  for (const index of [-1, 0.5, 2, NaN, Infinity, '1', undefined])
    assert.throws(() => materializeRenderOnlyCollectionFace(bytes, index as number, policy));
  const nonCollection = bytes.slice();
  new DataView(nonCollection.buffer).setUint32(0, 0x00010000);
  assert.throws(
    () => materializeRenderOnlyCollectionFace(nonCollection, 1, policy),
    /requires a collection/
  );
  assert.equal(hash(bytes), before);
});

test('a subarray with a nonzero backing-buffer offset keeps its entire backing buffer intact', () => {
  const source = collection(),
    backing = new Uint8Array(source.length + 37).fill(0x9a);
  backing.set(source, 13);
  const subarray = backing.subarray(13, 13 + source.length),
    before = hash(backing);
  const projected = materializeRenderOnlyCollectionFace(subarray, 1, policy);
  assert.equal(hash(backing), before);
  assert.deepEqual(projected.bytes, materializeRenderOnlyCollectionFace(source, 1, policy).bytes);
  assert.notEqual(projected.bytes.buffer, backing.buffer);
  assert.equal(sfntChecksum(projected.bytes), 0xb1b0afba);
});

test('explicit copied host settings remain settings, while malformed truthy values remain refused', () => {
  const bytes = collection(),
    before = hash(bytes);
  assert.equal(
    materializeRenderOnlyCollectionFace(bytes, 0, Object.freeze({ ...policy })).faceIndex,
    0
  );
  assert.notDeepEqual(
    materializeRenderOnlyCollectionFace(bytes, 0, policy).bytes,
    materializeRenderOnlyCollectionFace(bytes, 1, policy).bytes
  );
  for (const malformed of [
    null,
    true,
    'allow',
    { allowUnsignedDerivedFaces: 1, requireSignedFonts: false },
    { allowUnsignedDerivedFaces: true, requireSignedFonts: 0 },
  ]) {
    assert.throws(
      () => materializeRenderOnlyCollectionFace(bytes, 1, malformed as unknown as typeof policy),
      /host/
    );
  }
  assert.equal(hash(bytes), before);
});

for (const signed of [false, true]) {
  for (const [policyIndex, hostPolicy] of [
    undefined,
    { allowUnsignedDerivedFaces: false, requireSignedFonts: false },
    policy,
    { allowUnsignedDerivedFaces: true, requireSignedFonts: true },
  ].entries()) {
    test(`registration preserves existing faces (signed=${signed}, hostPolicy=${policyIndex})`, async () => {
      const bytes = collection(false, signed),
        before = hash(bytes),
        request = { family: 'Public Math Collection', weight: 400, style: 'normal' as const };
      const source = {
        request,
        id: 'public-math-face1',
        bytes,
        hash: 'sha256:' + before,
        faceIndex: 1,
      };
      const created: Uint8Array[] = [],
        added: object[] = [],
        removed: object[] = [];
      const result = await registerEmbeddedFontFaces([source], {
        renderOnlyCollectionPolicy: hostPolicy,
        fontSet: { add: (face) => added.push(face), delete: (face) => removed.push(face) },
        createFontFace: (_family, rendered, descriptors) => {
          assert.deepEqual(descriptors, { weight: '400', style: 'normal' });
          created.push(rendered);
          return { load: async () => {} };
        },
      });
      const allowed = !signed || hostPolicy === policy;
      assert.equal(result.installed, allowed ? 1 : 0);
      assert.equal(created.length, allowed ? 1 : 0);
      assert.equal(typeof result.alias(request.family), allowed ? 'string' : 'undefined');
      if (allowed) {
        const expected = signed
          ? materializeRenderOnlyCollectionFace(bytes, 1, policy).bytes
          : materializeCollectionFace(bytes, 1);
        assert.deepEqual(created[0], expected);
      }
      result.dispose();
      result.dispose();
      assert.deepEqual(removed, added);
      assert.equal(hash(bytes), before);
      assert.equal(source.faceIndex, 1);
      assert.equal(source.hash, 'sha256:' + before);
      assert.equal(source.request, request);
      assert.equal(source.bytes, bytes);
    });
  }
}

for (const signed of [false, true]) {
  for (const [policyIndex, hostPolicy] of [
    undefined,
    { allowUnsignedDerivedFaces: false, requireSignedFonts: false },
    policy,
    { allowUnsignedDerivedFaces: true, requireSignedFonts: true },
  ].entries()) {
    test(`a separate host policy retains the default browser environment (signed=${signed}, hostPolicy=${policyIndex})`, async () => {
      const bytes = collection(false, signed),
        before = hash(bytes),
        source = {
          request: { family: 'Public Math Collection', weight: 400, style: 'normal' as const },
          id: 'public-default-environment-face1',
          bytes,
          hash: 'sha256:' + before,
          faceIndex: 1,
        };
      const added = new Set<object>(),
        rendered: Uint8Array[] = [];
      const oldDocument = Object.getOwnPropertyDescriptor(globalThis, 'document'),
        oldFace = Object.getOwnPropertyDescriptor(globalThis, 'FontFace');
      class Face {
        readonly family: string;
        constructor(
          family: string,
          buffer: ArrayBuffer,
          descriptors: { readonly weight: string; readonly style: 'normal' | 'italic' }
        ) {
          this.family = family;
          assert.match(family, /^docx-embedded-/);
          assert.deepEqual(descriptors, { weight: '400', style: 'normal' });
          rendered.push(new Uint8Array(buffer));
        }
        async load(): Promise<this> {
          return this;
        }
      }
      Object.defineProperty(globalThis, 'document', {
        configurable: true,
        value: {
          fonts: {
            add: (face: object) => added.add(face),
            delete: (face: object) => added.delete(face),
          },
        },
      });
      Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: Face });
      try {
        const registration = await registerEmbeddedFontFaces([source], undefined, [], hostPolicy);
        const allowed = !signed || hostPolicy === policy;
        assert.equal(registration.installed, allowed ? 1 : 0);
        assert.equal(added.size, allowed ? 1 : 0);
        assert.equal(rendered.length, allowed ? 1 : 0);
        assert.equal(
          typeof registration.alias(source.request.family),
          allowed ? 'string' : 'undefined'
        );
        if (allowed) {
          const expected = signed
            ? materializeRenderOnlyCollectionFace(bytes, 1, policy).bytes
            : materializeCollectionFace(bytes, 1);
          assert.deepEqual(rendered[0], expected);
        }
        registration.dispose();
        registration.dispose();
        assert.equal(added.size, 0);
        // An explicitly empty environment retains its previous conservative behavior.
        assert.equal((await registerEmbeddedFontFaces([source], {}, [], hostPolicy)).installed, 0);
        assert.equal(rendered.length, allowed ? 1 : 0);
        assert.equal(hash(bytes), before);
        assert.equal(source.bytes, bytes);
      } finally {
        if (oldDocument) Object.defineProperty(globalThis, 'document', oldDocument);
        else Reflect.deleteProperty(globalThis, 'document');
        if (oldFace) Object.defineProperty(globalThis, 'FontFace', oldFace);
        else Reflect.deleteProperty(globalThis, 'FontFace');
      }
    });
  }
}
