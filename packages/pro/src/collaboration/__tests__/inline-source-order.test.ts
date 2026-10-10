/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { orderInlineSourceMembers } from '../document/inline-source-order.ts';
import type { InlineSourceCut } from '../document/inline-source-cuts.ts';
const limits = { maxChildren: 8, maxTreeDepth: 4 };
function fixture() {
  const doc = new Y.Doc(),
    text = doc.getText('source');
  text.insert(0, 'North Bay');
  const member = (id: string, index = 4, afterAtomId?: string) => ({
    id,
    cut: {
      sourceId: 'source',
      text,
      index,
      ...(afterAtomId === undefined ? {} : { afterAtomId }),
    } satisfies InlineSourceCut,
  });
  return { doc, member };
}
test('literal chains follow their predecessors before replica tie breaking', () => {
  const { doc, member } = fixture();
  try {
    const members = [
      member('a', 4, 'b'),
      member('b', 4, 'c'),
      member('c'),
      member('d', 2),
      member('z'),
    ];
    expect(orderInlineSourceMembers(members, () => null, limits).map((item) => item.id)).toEqual([
      'd',
      'c',
      'b',
      'a',
      'z',
    ]);
    expect(
      orderInlineSourceMembers([...members].reverse(), () => null, limits).map((item) => item.id)
    ).toEqual(['d', 'c', 'b', 'a', 'z']);
  } finally {
    doc.destroy();
  }
});
test('deleted predecessors retain their boundary without returning to the visible sequence', () => {
  const { doc, member } = fixture();
  try {
    expect(
      orderInlineSourceMembers(
        [member('literal', 4, 'deleted')],
        () => member('deleted').cut,
        limits
      ).map((item) => item.id)
    ).toEqual(['literal']);
  } finally {
    doc.destroy();
  }
});
test('unknown, cross-source, and cross-position predecessors are refused', () => {
  const { doc, member } = fixture();
  try {
    const members = [member('literal', 4, 'previous')];
    for (const previous of [
      null,
      member('previous', 3).cut,
      { ...member('previous').cut, sourceId: 'other' },
    ])
      expect(() => orderInlineSourceMembers(members, () => previous, limits)).toThrow(
        'inline source predecessor'
      );
  } finally {
    doc.destroy();
  }
});
test('cyclic predecessors are refused independently of delivery order', () => {
  const { doc, member } = fixture();
  try {
    const members = [member('a', 4, 'b'), member('b', 4, 'a')];
    expect(() => orderInlineSourceMembers(members, () => null, limits)).toThrow(
      'inline source cycle'
    );
    expect(() => orderInlineSourceMembers(members.reverse(), () => null, limits)).toThrow(
      'inline source cycle'
    );
  } finally {
    doc.destroy();
  }
});
test('member counts and transitive predecessor depths have finite limits', () => {
  const { doc, member } = fixture();
  try {
    const members = [
      member('a', 4, 'b'),
      member('b', 4, 'c'),
      member('c', 4, 'd'),
      member('d', 4, 'e'),
      member('e'),
    ];
    expect(() => orderInlineSourceMembers(members, () => null, limits)).toThrow('tree-too-deep');
    expect(() =>
      orderInlineSourceMembers(members, () => null, { ...limits, maxChildren: 4 })
    ).toThrow('too-many-children');
  } finally {
    doc.destroy();
  }
});
