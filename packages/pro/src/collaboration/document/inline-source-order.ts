/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { CollaborationSchemaError } from '../schema.ts';
import type { InlineSourceCut } from './inline-source-cuts.ts';
import type { DocumentLimits } from './limits.ts';

interface Member {
  readonly id: string;
  readonly cut: InlineSourceCut;
}

/** Equal-position literals follow their recorded predecessor. Replica IDs break other ties. */
export function orderInlineSourceMembers(
  members: readonly Member[],
  read: (id: string) => InlineSourceCut | null,
  limits: Pick<DocumentLimits, 'maxChildren' | 'maxTreeDepth'>
): Member[] {
  if (members.length > limits.maxChildren) throw new CollaborationSchemaError('too-many-children');
  const sorted = [...members].sort(
    (a, b) => a.cut.index - b.cut.index || (a.id < b.id ? -1 : a.id > b.id ? 1 : 0)
  );
  const byId = new Map(sorted.map((member) => [member.id, member]));
  const state = new Map<string, number>(),
    depths = new Map<string, number>(),
    result: Member[] = [];
  for (const first of sorted) {
    if (state.get(first.id) === 2) continue;
    const stack = [{ member: first, expanded: false }];
    while (stack.length > 0) {
      const frame = stack[stack.length - 1]!,
        member = frame.member;
      if (frame.expanded) {
        const depth = 1 + (depths.get(member.cut.afterAtomId ?? '') ?? 0);
        if (depth > limits.maxTreeDepth) throw new CollaborationSchemaError('tree-too-deep');
        depths.set(member.id, depth);
        state.set(member.id, 2);
        result.push(member);
        stack.pop();
        continue;
      }
      if (state.get(member.id) === 2) {
        stack.pop();
        continue;
      }
      if (state.get(member.id) === 1)
        throw new CollaborationSchemaError('invalid-bound', 'inline source cycle');
      state.set(member.id, 1);
      frame.expanded = true;
      const after = member.cut.afterAtomId;
      if (after === undefined) continue;
      const previous = byId.get(after),
        cut = previous?.cut ?? read(after);
      if (!cut || cut.sourceId !== member.cut.sourceId || cut.index !== member.cut.index)
        throw new CollaborationSchemaError('invalid-bound', 'inline source predecessor');
      if (previous) {
        if (state.get(after) === 1)
          throw new CollaborationSchemaError('invalid-bound', 'inline source cycle');
        if (stack.length >= limits.maxTreeDepth)
          throw new CollaborationSchemaError('tree-too-deep');
        if (state.get(after) !== 2) stack.push({ member: previous, expanded: false });
      }
    }
  }
  return result;
}
