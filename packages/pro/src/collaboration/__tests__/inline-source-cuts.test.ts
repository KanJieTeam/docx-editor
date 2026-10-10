/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { afterEach, expect, test } from 'bun:test';
import { createPeerHarness, zipDocument } from './document-peer-support.ts';
import { DocumentRegistry } from '../document/registry.ts';
import { NODE_INLINE_SOURCE_CUT_FIELD } from '../document/inline-source-cuts.ts';
const TEXT = 'North Bay 4 River Avenue';
const harness = createPeerHarness('inline-source-cuts', { offlineEditing: true });
const registries: DocumentRegistry[] = [];
afterEach(() => {
  for (const registry of registries.splice(0)) registry.destroy();
  harness.cleanup();
});
function registryOf(peer: Awaited<ReturnType<typeof harness.pair>>['alice']) {
  const registry = new DocumentRegistry(peer.ydoc);
  registry.rebuildDerivedIndexes();
  registries.push(registry);
  return registry;
}
for (const op of ['insertHardBreak', 'insertTab'] as const)
  for (const offset of [0, 9, TEXT.length]) {
    test(`${op} records its source boundary in the creation journal at ${offset}`, async () => {
      const { alice, bob } = await harness.pair(
        zipDocument('<w:p><w:r><w:t xml:space="preserve">' + TEXT + '</w:t></w:r></w:p>')
      );
      const paragraphId = harness.paragraphIdAt(alice, 0);
      harness.apply(alice, [{ op, paragraphId, offset }]);
      const left = registryOf(alice),
        right = registryOf(bob);
      const atoms = [...left.schema.nodes].filter(([, node]) =>
        node.has(NODE_INLINE_SOURCE_CUT_FIELD)
      );
      expect(atoms).toHaveLength(1);
      const atomId = atoms[0]![0],
        cut = left.inlineSourceCut(atomId)!;
      expect(cut.index).toBe(offset);
      expect(cut.text.toString()).toBe(TEXT);
      expect(right.inlineSourceCut(atomId)!.index).toBe(offset);
      expect(left.inlineAtomsForSource(cut.sourceId)).toEqual([atomId]);
      alice.room.session.undo();
      expect(left.inlineSourceCut(atomId)).toBeNull();
      expect(right.inlineSourceCut(atomId)).toBeNull();
      alice.room.session.redo();
      expect(left.inlineSourceCut(atomId)!.index).toBe(offset);
      expect(right.inlineSourceCut(atomId)!.index).toBe(offset);
    });
  }
