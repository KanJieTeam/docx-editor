/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import type { DocumentRegistry } from './registry.ts';
import type { SplitTextRange, SplitTextSources } from './split-text-sources.ts';
import type { InlineSourceCuts } from './inline-source-cuts.ts';
import { childArrayOf, NODE_TEXT_FIELD } from './schema.ts';

export function authoredInlineSourceRange(
  registry: DocumentRegistry,
  sources: SplitTextSources,
  id: string
): SplitTextRange | null {
  const range = sources.range(id);
  if (range) return range;
  const text = registry.schema.nodes.get(id)?.get(NODE_TEXT_FIELD);
  return text instanceof Y.Text
    ? {
        sourceId: id,
        text,
        start: 0,
        end: text.length,
        startAssoc: -1,
        endAssoc: 0,
        startAnchorDeleted: false,
        endAnchorDeleted: false,
      }
    : null;
}

/** Inspect only one run, its text wrappers, and their leaves. Never scan unrelated nodes. */
export function noteInlineSourceOwner(
  registry: DocumentRegistry,
  cuts: InlineSourceCuts,
  id: string,
  depth = 0
): void {
  if (depth > 2) return;
  const rec = registry.schema.nodes.get(id);
  if (!rec) return;
  const kind = registry.kindOf(id);
  if (kind === 'textValue') {
    const range = registry.authoredSourceRange(id);
    if (range) cuts.noteSource(range.sourceId);
  } else if (kind === 'text' || kind === 'run') {
    const children = childArrayOf(rec);
    if (children && children.length <= registry.limits.maxChildren)
      for (const child of children.toArray())
        noteInlineSourceOwner(registry, cuts, child, depth + 1);
  }
}
