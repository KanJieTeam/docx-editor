/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type { DocumentRegistry } from './registry.ts';
import { childArrayOf } from './schema.ts';
import { CollaborationSchemaError } from '../schema.ts';
import { orderInlineSourceMembers } from './inline-source-order.ts';

interface Slice {
  readonly textId: string;
  readonly leafId: string;
  readonly runId: string;
  readonly start: number;
  readonly end: number;
}

/** Current source ownership, separate from the character ranges authored by run splits. */
export class InlineSourcePlacement {
  private running = false;
  private readonly ends = new Map<string, number>();
  private readonly values = new Map<string, string>();
  private readonly forcedBySource = new Map<string, ReadonlySet<string>>();
  private readonly leavesBySource = new Map<string, ReadonlySet<string>>();
  private readonly changedLeaves = new Set<string>();
  private readonly childrenByRun = new Map<string, readonly string[]>();
  private readonly ownerByChild = new Map<string, string>();
  private readonly runsBySource = new Map<string, ReadonlySet<string>>();
  private readonly ownersBySource = new Map<string, ReadonlySet<string>>();
  private readonly overlayValues = new Map<string, string>();
  private overlayBase: ReadonlyMap<string, string> | null = null;
  private loserBase: ReadonlySet<string> | null = null;
  private loserResult: ReadonlySet<string> | null = null;

  constructor(private readonly registry: DocumentRegistry) {}

  visibleLosers(raw: ReadonlySet<string>): ReadonlySet<string> {
    if (this.loserBase === raw && this.loserResult) return this.loserResult;
    if (this.forcedBySource.size === 0) return raw;
    const result = new Set(raw);
    for (const ids of this.forcedBySource.values()) for (const id of ids) result.delete(id);
    this.loserBase = raw;
    this.loserResult = result;
    return result;
  }
  effectiveEnd(id: string, fallback: number): number {
    return this.ends.get(id) ?? fallback;
  }
  textValue(id: string, fallback: string | null): string | null {
    return this.values.get(id) ?? fallback;
  }
  children(id: string, fallback: readonly string[]): readonly string[] {
    // A journal can move a borrowed member into a new run before Yjs delivers its events.
    // The previous placement must not remove that member from an intermediate split source.
    if (this.registry.hasUnobservedWrites() && !this.childrenByRun.has(id)) return fallback;
    return (this.childrenByRun.get(id) ?? fallback).filter((child) => {
      const owner = this.ownerByChild.get(child);
      return owner === undefined || owner === id;
    });
  }
  hasChildren(id: string): boolean {
    return this.childrenByRun.has(id);
  }
  parentOf(id: string): string | null {
    return this.ownerByChild.get(id) ?? null;
  }
  overlays(base: ReadonlyMap<string, string>, changed: ReadonlySet<string>) {
    const changedIds = new Set(changed);
    for (const id of this.changedLeaves) changedIds.add(id);
    if (this.overlayBase !== base) {
      this.overlayValues.clear();
      for (const [id, value] of base) this.overlayValues.set(id, value);
      this.overlayBase = base;
    }
    for (const id of changedIds) {
      const value = this.values.get(id) ?? base.get(id);
      if (value === undefined) this.overlayValues.delete(id);
      else this.overlayValues.set(id, value);
    }
    this.changedLeaves.clear();
    return { values: this.overlayValues, changedIds };
  }
  reset(): void {
    this.loserBase = null;
    this.loserResult = null;
    for (const id of this.ends.keys()) this.changedLeaves.add(id);
    this.ends.clear();
    this.values.clear();
    this.forcedBySource.clear();
    this.leavesBySource.clear();
    this.childrenByRun.clear();
    this.ownerByChild.clear();
    this.runsBySource.clear();
    this.ownersBySource.clear();
  }

  reconcile(sources: readonly string[]): void {
    if (this.running || sources.length === 0) return;
    this.running = true;
    try {
      for (const source of sources) this.reconcileSource(source);
    } finally {
      this.running = false;
    }
  }

  private rawChildren(id: string): readonly string[] {
    const record = this.registry.schema.nodes.get(id);
    const children = record ? childArrayOf(record) : null;
    if (children && children.length > this.registry.limits.maxChildren)
      throw new CollaborationSchemaError('too-many-children');
    return children?.toArray() ?? [];
  }

  private slice(textId: string, runId: string, sourceId: string): Slice | null {
    if (this.registry.kindOf(textId) !== 'text') return null;
    const children = this.rawChildren(textId);
    if (children.length !== 1) return null;
    const leafId = children[0]!;
    const range = this.registry.authoredSourceRange(leafId);
    if (!range || range.sourceId !== sourceId) return null;
    return { textId, leafId, runId, start: range.start, end: range.end };
  }

  private reconcileSource(sourceId: string): void {
    this.loserResult = null;
    const registry = this.registry;
    const cuts = orderInlineSourceMembers(
      registry
        .inlineAtomsForSource(sourceId)
        .filter((id) => !registry.isTombstoned(id) && !this.inDeletedRevision(id))
        .map((id) => ({ id, cut: registry.inlineSourceCut(id)! }))
        .filter((item) => item.cut !== null),
      (id) => registry.inlineSourceCut(id),
      registry.limits
    );
    const rawLosers = registry.rawReplacementLosers();
    const slices: Slice[] = [];
    for (const leafId of registry.sourceAliases(sourceId)) {
      for (const textId of registry.listedParents(leafId)) {
        if (registry.isTombstoned(textId) || rawLosers.has(textId)) continue;
        for (const runId of registry.listedParents(textId)) {
          if (
            registry.kindOf(runId) !== 'run' ||
            registry.isTombstoned(runId) ||
            rawLosers.has(runId) ||
            registry.parentOf(runId) === null
          )
            continue;
          const slice = this.slice(textId, runId, sourceId);
          if (slice) slices.push(slice);
        }
      }
    }
    // Earlier, wider ranges own an overlapping borrowed continuation. Run-format winners remain authoritative.
    slices.sort(
      (a, b) =>
        a.start - b.start ||
        b.end - a.end ||
        (a.textId < b.textId ? -1 : a.textId > b.textId ? 1 : 0)
    );
    const starts: Slice[] = [];
    for (const slice of slices)
      if (starts[starts.length - 1]?.start !== slice.start) starts.push(slice);
    const carriers = starts.map((slice, index) => ({
      ...slice,
      end: Math.min(slice.end, starts[index + 1]?.start ?? slice.end),
    }));
    const assignments = new Map<string, typeof cuts>();
    for (const item of cuts) {
      const owner =
        carriers.find((slice) => slice.start <= item.cut.index && item.cut.index < slice.end) ??
        [...carriers].reverse().find((slice) => slice.end === item.cut.index);
      if (!owner) continue;
      const owned = assignments.get(owner.textId) ?? [];
      owned.push(item);
      assignments.set(owner.textId, owned);
    }
    const insertions = new Map<string, readonly string[]>(),
      nextEnds = new Map<string, number>(),
      forced = new Set<string>();
    for (const carrier of carriers)
      if (carrier.end !== registry.authoredSourceRange(carrier.leafId)?.end)
        nextEnds.set(carrier.leafId, carrier.end);
    const dirtyRuns = new Set<string>(),
      moving = new Set<string>();
    for (const carrier of carriers) {
      const owned = assignments.get(carrier.textId) ?? [];
      if (owned.length === 0) continue;
      const sequence: string[] = [];
      let current = carrier;
      for (const item of owned) {
        forced.add(item.id);
        moving.add(item.id);
        if (item.cut.index === current.start) {
          sequence.push(item.id);
          continue;
        }
        sequence.push(current.textId, item.id);
        nextEnds.set(current.leafId, item.cut.index);
        if (item.cut.index === carrier.end) {
          current = { ...current, start: carrier.end };
          continue;
        }
        const continuationId = item.cut.continuationId;
        const continuation =
          continuationId && !registry.isTombstoned(continuationId)
            ? this.slice(continuationId, carrier.runId, sourceId)
            : null;
        const existing = carriers.find(
          (slice) => slice.start === item.cut.index && slice.runId === carrier.runId
        );
        const next = continuation ?? existing;
        if (!next)
          throw new CollaborationSchemaError('invalid-bound', 'inline source continuation');
        if (next.start !== item.cut.index || next.end < carrier.end)
          throw new CollaborationSchemaError('invalid-bound', 'inline source interval');
        current = next;
        forced.add(current.textId);
        moving.add(current.textId);
      }
      if (current.start < carrier.end) {
        sequence.push(current.textId);
        nextEnds.set(current.leafId, carrier.end);
      }
      insertions.set(carrier.textId, sequence);
      dirtyRuns.add(carrier.runId);
    }
    const desired = new Map<string, readonly string[]>();
    for (const runId of dirtyRuns) {
      const children = this.rawChildren(runId),
        next: string[] = [];
      for (const id of children) {
        const insertion = insertions.get(id);
        if (insertion) {
          for (const child of insertion) next.push(child);
        } else if (!moving.has(id)) next.push(id);
      }
      desired.set(runId, next);
    }
    // The plan is complete before any child list changes. It retains the original shared node identities.
    for (const runId of this.runsBySource.get(sourceId) ?? []) {
      if (this.childrenByRun.delete(runId)) this.changedLeaves.add(runId);
    }
    for (const id of this.ownersBySource.get(sourceId) ?? []) this.ownerByChild.delete(id);
    const owners = new Set<string>();
    for (const [runId, next] of desired) {
      const previous = this.childrenByRun.get(runId);
      if (
        !previous ||
        previous.length !== next.length ||
        previous.some((id, index) => id !== next[index])
      )
        this.changedLeaves.add(runId);
      this.childrenByRun.set(runId, next);
      for (const id of next)
        if (moving.has(id)) {
          this.ownerByChild.set(id, runId);
          owners.add(id);
        }
    }
    this.runsBySource.set(sourceId, new Set(desired.keys()));
    this.ownersBySource.set(sourceId, owners);
    for (const id of this.leavesBySource.get(sourceId) ?? []) {
      if (!nextEnds.has(id)) {
        if (this.ends.delete(id)) this.changedLeaves.add(id);
        this.values.delete(id);
      }
    }
    const leaves = new Set<string>();
    const sourceText = cuts[0]?.cut.text.toString();
    for (const [id, end] of nextEnds) {
      leaves.add(id);
      const range = registry.authoredSourceRange(id);
      if (!range) continue;
      const value = (sourceText ?? range.text.toString()).slice(range.start, end);
      if (this.ends.get(id) !== end || this.values.get(id) !== value) this.changedLeaves.add(id);
      this.ends.set(id, end);
      this.values.set(id, value);
    }
    this.leavesBySource.set(sourceId, leaves);
    this.forcedBySource.set(sourceId, forced);
  }

  private inDeletedRevision(id: string): boolean {
    const seen = new Set<string>();
    let current = this.registry.parentOf(id);
    for (let depth = 0; current !== null && depth < this.registry.limits.maxTreeDepth; depth += 1) {
      if (seen.has(current))
        throw new CollaborationSchemaError('invalid-bound', 'inline source ancestry');
      seen.add(current);
      if (this.registry.kindOf(current) === 'revisionDelete') return true;
      current = this.registry.parentOf(current);
    }
    if (current !== null) throw new CollaborationSchemaError('tree-too-deep');
    return false;
  }
}
