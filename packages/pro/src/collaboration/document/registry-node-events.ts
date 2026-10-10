/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import type { SplitDedupIndex } from './split-dedup.ts';
import type { SplitTextSources } from './split-text-sources.ts';
import type { InlineSourceCuts } from './inline-source-cuts.ts';
import { rejectDangerousKey } from './limits.ts';
import { NODE_DELETED_FIELD, NODE_REPLACED_BY_FIELD } from './schema.ts';
interface NodeEventIndex {
  readonly splitDedup: SplitDedupIndex;
  readonly splitTextSources: SplitTextSources;
  readonly inlineSourceCuts: InlineSourceCuts;
  noteInlineSourceOwner(id: string): void;
  noteNodeCount(action: 'add' | 'update' | 'delete'): void;
  syncChildListings(id: string): readonly string[];
  isTombstoned(id: string): boolean;
  syncAdoptee(id: string): void;
}
/** Update the node indexes before resolving contested parents and source placement. */
export function applyRegistryNodeEvents(
  events: readonly Y.YEvent<Y.AbstractType<unknown>>[],
  index: NodeEventIndex
): ReadonlySet<string> {
  const changed = new Set<string>();
  for (const event of events) {
    // Text and boundary metadata do not change which split branches are reachable.
    if (
      !(event.target instanceof Y.Text) &&
      (event.path.length === 0 ||
        event.target instanceof Y.Array ||
        (event.target instanceof Y.Map &&
          [...event.changes.keys.keys()].some((key) =>
            ['deleted', 'replacedBy', 'splitFrom', 'splitLineage', 'children'].includes(String(key))
          )))
    )
      index.splitDedup.invalidate();
    if (event.path.length > 0) index.splitTextSources.noteChanged(String(event.path[0]));
    if (event.path.length > 0) index.inlineSourceCuts.noteChanged(String(event.path[0]));
    if (event.path.length > 0) index.noteInlineSourceOwner(String(event.path[0]));
    // A remote applyUpdate delivers a new element record with its children already filled.
    // Yjs does not emit a child-array event for that initial fill. Skipping it left
    // `parentOf` null, so an attribute-only journal could not dirty the part root and the
    // receiving replica kept the cached `commentsExtended.xml`.
    if (event.path.length === 0 && event.target instanceof Y.Map) {
      for (const [key, change] of event.changes.keys) {
        index.noteNodeCount(change.action);
        if (change.action === 'delete' || rejectDangerousKey(String(key))) continue;
        // A run a peer split off carries its origin; index it so the loser-dedup sees the
        // concurrent split the moment the remote record arrives, not only after a rebuild.
        index.splitDedup.indexExisting(String(key));
        index.splitTextSources.indexExisting(String(key));
        index.inlineSourceCuts.indexExisting(String(key));
        index.noteInlineSourceOwner(String(key));
        for (const childId of index.syncChildListings(String(key))) changed.add(childId);
      }
      continue;
    }
    // Undo retains record containers. Redo restores scalar provenance on those existing
    // maps, so a late joiner must index field changes as well as new map entries.
    if (
      event.target instanceof Y.Map &&
      event.path.length === 1 &&
      (event.changes.keys.has('splitFrom') || event.changes.keys.has('splitLineage'))
    ) {
      index.splitDedup.indexExisting(String(event.path[0]));
    }
    if (event.target instanceof Y.Array && event.path.length > 0) {
      const parentId = String(event.path[0]);
      for (const childId of index.syncChildListings(parentId)) changed.add(childId);
      if (index.isTombstoned(parentId)) index.syncAdoptee(parentId);
      continue;
    }
    if (
      event.target instanceof Y.Map &&
      event.path.length === 1 &&
      (event.changes.keys.has(NODE_DELETED_FIELD) || event.changes.keys.has(NODE_REPLACED_BY_FIELD))
    ) {
      index.syncAdoptee(String(event.path[0]));
    }
  }
  return changed;
}
