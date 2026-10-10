/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { afterEach, expect, spyOn, test } from 'bun:test';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support.ts';
import type { OoxmlNode, TreeDocOp } from '@docx-editor.dev/core/store';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import type { DocumentRegistry } from '../document/registry.ts';
import { commitSessionTreeOps } from '../../../../core/src/binding/tree-session-apply.ts';

const source = 'North Bay 4 River Avenue';
const harness = createPeerHarness('inline-source-placement', { offlineEditing: true });
afterEach(() => harness.cleanup());
let operation = 0;
function apply(peer: Peer, ops: readonly TreeDocOp[]): void {
  const errors: string[] = [];
  const stop = peer.room.session.subscribeStatus((status, reason, detail) => {
    if (status === 'error') errors.push(String(reason) + ':' + String(detail));
  });
  const scope = { kind: 'body' } as const;
  const refused = peer.room.session.gateOperations(ops, scope);
  expect(refused).toBeNull();
  const result = commitSessionTreeOps(peer.store, ops, undefined, undefined, scope, {
    historyGroup: Symbol('one-edit'),
    recordsHistory: false,
    actorId: peer.room.session.identity.actorId,
    operationId: `inline-source-${++operation}`,
  });
  expect(result.committed).toBe(true);
  expect(result.rejected).toBe(false);
  peer.port.flushPendingJournals();
  stop();
  expect(errors).toEqual([]);
}
test('editing an indexed source does not scan unrelated shared nodes', async () => {
  const unrelated = Array.from(
    { length: 100 },
    (_, index) => `<w:p><w:r><w:t>Other paragraph ${index}</w:t></w:r></w:p>`
  ).join('');
  const { alice, bob } = await harness.pair(
    zipDocument(`<w:p><w:r><w:t>${source}</w:t></w:r></w:p>` + unrelated)
  );
  apply(alice, [{ op: 'insertTab', paragraphId: harness.paragraphIdAt(alice, 0), offset: 10 }]);
  const registry = (alice.room.session as unknown as { registry: DocumentRegistry }).registry;
  const scan = spyOn(registry.schema.nodes, 'forEach').mockImplementation(() => {
    throw new Error('whole node scan');
  });
  try {
    apply(alice, [
      { op: 'insertText', paragraphId: harness.paragraphIdAt(alice, 0), offset: 2, text: 'X' },
    ]);
    expect(scan).not.toHaveBeenCalled();
  } finally {
    scan.mockRestore();
  }
  harness.expectConverged(alice, bob);
});
function textOf(node: OoxmlNode): string {
  if (node.kind === 'textValue') return node.value;
  if (node.kind === 'hardBreak') return '\n';
  if (node.kind === 'tab') return '\t';
  return node.children.map(textOf).join('');
}
function text(peer: Peer): string {
  const document = harness.packageOf(peer);
  return textOf(document.parts.get(document.mainDocumentPart)!.root);
}
for (const reverse of [false, true])
  for (const author of ['alice', 'bob'] as const) {
    test(`independent inline cuts survive merging (${author}, reverse=${reverse})`, async () => {
      const pair = await harness.pair(zipDocument(`<w:p><w:r><w:t>${source}</w:t></w:r></w:p>`));
      pair.alice.ydoc.clientID = reverse ? 74002 : 74001;
      pair.bob.ydoc.clientID = reverse ? 74001 : 74002;
      const marker = pair[author],
        other = pair[author === 'alice' ? 'bob' : 'alice'];
      pair.pause();
      apply(marker, [
        { op: 'insertHardBreak', paragraphId: harness.paragraphIdAt(marker, 0), offset: 10 },
      ]);
      apply(other, [{ op: 'insertTab', paragraphId: harness.paragraphIdAt(other, 0), offset: 16 }]);
      pair.resume();
      const expected = source.slice(0, 10) + '\n' + source.slice(10, 16) + '\t' + source.slice(16);
      expect(text(marker)).toBe(expected);
      expect(text(other)).toBe(expected);
      harness.expectConverged(marker, other);
      const joined = await harness.join(marker, 'cold');
      expect(text(joined)).toBe(expected);
    });
    for (const [start, end] of [
      [0, 4],
      [4, 13],
      [10, 17],
    ] as const) {
      test(`two concurrent generations preserve inline cuts (${author}, reverse=${reverse}, ${start}:${end})`, async () => {
        const pair = await harness.pair(zipDocument(`<w:p><w:r><w:t>${source}</w:t></w:r></w:p>`));
        pair.alice.ydoc.clientID = reverse ? 74002 : 74001;
        pair.bob.ydoc.clientID = reverse ? 74001 : 74002;
        const marker = pair[author],
          other = pair[author === 'alice' ? 'bob' : 'alice'];
        pair.pause();
        apply(marker, [
          { op: 'insertHardBreak', paragraphId: harness.paragraphIdAt(marker, 0), offset: 10 },
        ]);
        apply(other, [
          {
            op: 'setRunProperties',
            paragraphId: harness.paragraphIdAt(other, 0),
            start,
            end,
            properties: [{ localName: 'b' }],
          },
        ]);
        pair.resume();
        const first = source.slice(0, 10) + '\n' + source.slice(10);
        expect(text(marker)).toBe(first);
        expect(text(other)).toBe(first);
        pair.pause();
        apply(other, [
          { op: 'insertTab', paragraphId: harness.paragraphIdAt(other, 0), offset: 17 },
        ]);
        apply(marker, [
          {
            op: 'setRunProperties',
            paragraphId: harness.paragraphIdAt(marker, 0),
            start: 11,
            end: 20,
            properties: [{ localName: 'i' }],
          },
        ]);
        pair.resume();
        const expected = first.slice(0, 17) + '\t' + first.slice(17);
        expect(text(marker)).toBe(expected);
        expect(text(other)).toBe(expected);
        harness.expectConverged(marker, other);
        const joined = await harness.join(marker, 'cold');
        expect(text(joined)).toBe(expected);
        other.room.session.undo();
        expect(text(marker)).toBe(first);
        expect(text(other)).toBe(first);
        expect(text(joined)).toBe(first);
        other.room.session.redo();
        expect(text(marker)).toBe(expected);
        expect(text(other)).toBe(expected);
        expect(text(joined)).toBe(expected);
        marker.room.session.undo();
        expect(text(marker)).toBe(expected);
        expect(text(other)).toBe(expected);
        marker.room.session.redo();
        expect(text(marker)).toBe(expected);
        expect(text(other)).toBe(expected);
        const afterRedoJoin = await harness.join(marker, 'cold-after-redo');
        expect(text(afterRedoJoin)).toBe(expected);
        harness.expectConverged(marker, afterRedoJoin);
        for (const peer of [marker, other]) {
          const pkg = harness.packageOf(peer);
          expect(
            paragraphTextOf(pkg.parts.get(pkg.mainDocumentPart)!, harness.paragraphIdAt(peer, 0))
          ).toBe(expected);
        }
        pair.pause();
        apply(other, [
          { op: 'deleteText', paragraphId: harness.paragraphIdAt(other, 0), start: 10, end: 11 },
        ]);
        apply(marker, [
          {
            op: 'insertText',
            paragraphId: harness.paragraphIdAt(marker, 0),
            offset: 18,
            text: 'X',
          },
        ]);
        pair.resume();
        const withoutBreak = expected.slice(0, 10) + expected.slice(11);
        const afterDelete = withoutBreak.slice(0, 17) + 'X' + withoutBreak.slice(17);
        expect(text(marker)).toBe(afterDelete);
        expect(text(other)).toBe(afterDelete);
        expect(text(joined)).toBe(afterDelete);
        harness.expectConverged(marker, other);
        pair.pause();
        apply(other, [
          { op: 'insertText', paragraphId: harness.paragraphIdAt(other, 0), offset: 0, text: 'Y' },
        ]);
        apply(marker, [
          {
            op: 'setRunProperties',
            paragraphId: harness.paragraphIdAt(marker, 0),
            start: 14,
            end: 20,
            properties: [{ localName: 'b' }],
          },
        ]);
        pair.resume();
        expect(text(marker)).toBe('Y' + afterDelete);
        expect(text(other)).toBe('Y' + afterDelete);
        harness.expectConverged(marker, other);
      });
    }
  }
