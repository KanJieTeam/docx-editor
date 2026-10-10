/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import { CollaborationSchemaError } from '../schema.ts';
import { rejectDangerousKey, type DocumentLimits } from './limits.ts';
import { isNodeMap, NODE_TEXT_FIELD, readNodeShell } from './schema.ts';

/** Source character boundary owned by an inserted hard break or tab. Never serialized to DOCX. */
export const NODE_INLINE_SOURCE_CUT_FIELD = 'inlineSourceCut';
const RESTORED_CUT_ORIGIN = Object.freeze({ kind: 'docx-package-restored-inline-cut' });
interface Cut {
  readonly sourceId: string;
  readonly anchor: Y.RelativePosition;
  readonly continuationId?: string;
  readonly afterAtomId?: string;
}
export interface InlineSourceCut {
  readonly sourceId: string;
  readonly text: Y.Text;
  readonly index: number;
  readonly continuationId?: string;
  readonly afterAtomId?: string;
}
function invalid(code: 'invalid-string' | 'invalid-bound'): never {
  throw new CollaborationSchemaError(code, NODE_INLINE_SOURCE_CUT_FIELD);
}
function object(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}
function identity(value: unknown): boolean {
  return (
    object(value) &&
    Object.keys(value).every((key) => key === 'client' || key === 'clock') &&
    Number.isSafeInteger(value.client) &&
    Number(value.client) >= 0 &&
    Number.isSafeInteger(value.clock) &&
    Number(value.clock) >= 0
  );
}
function position(value: unknown): Y.RelativePosition {
  if (
    !object(value) ||
    Object.keys(value).some((key) => !['type', 'item', 'assoc'].includes(key)) ||
    !identity(value.type) ||
    (value.item !== undefined && !identity(value.item)) ||
    (value.assoc !== 0 && value.assoc !== -1)
  )
    return invalid('invalid-string');
  return Y.createRelativePositionFromJSON(value);
}

/** An event-maintained source index. Only a cold load visits the shared node map. */
export class InlineSourceCuts {
  private readonly sourceByAtom = new Map<string, string>();
  private readonly atomsBySource = new Map<string, Set<string>>();
  private readonly pendingSources = new Set<string>();
  constructor(
    private readonly nodes: Y.Map<Y.Map<unknown>>,
    private readonly doc: Y.Doc,
    private readonly limits: Pick<DocumentLimits, 'maxTextLength' | 'maxStringLength'>
  ) {}

  register(
    atomId: string,
    sourceId: string,
    text: Y.Text,
    index: number,
    continuationId?: string,
    afterAtomId?: string
  ): void {
    const atom = this.nodes.get(atomId);
    if (
      !isNodeMap(atom) ||
      text !== this.sourceText(sourceId) ||
      !Number.isSafeInteger(index) ||
      index < 0 ||
      index > text.length
    )
      return invalid('invalid-bound');
    if (
      continuationId !== undefined &&
      (continuationId.length === 0 ||
        continuationId.length > this.limits.maxStringLength ||
        rejectDangerousKey(continuationId))
    )
      return invalid('invalid-string');
    if (
      afterAtomId !== undefined &&
      (afterAtomId === atomId ||
        afterAtomId.length === 0 ||
        afterAtomId.length > this.limits.maxStringLength ||
        rejectDangerousKey(afterAtomId))
    )
      return invalid('invalid-string');
    const encoded = JSON.stringify({
      sourceId,
      ...(continuationId === undefined ? {} : { continuationId }),
      ...(afterAtomId === undefined ? {} : { afterAtomId }),
      cut: Y.relativePositionToJSON(
        Y.createRelativePositionFromTypeIndex(text, index, index === 0 ? -1 : 0)
      ),
    });
    if (encoded.length > this.limits.maxStringLength) return invalid('invalid-string');
    atom.set(NODE_INLINE_SOURCE_CUT_FIELD, encoded);
    this.indexExisting(atomId);
  }

  read(atomId: string): InlineSourceCut | null {
    const cut = this.decode(atomId);
    if (!cut) return null;
    const text = this.sourceText(cut.sourceId);
    const resolved = Y.createAbsolutePositionFromRelativePosition(cut.anchor, this.doc);
    if (!resolved || resolved.type !== text || resolved.index < 0 || resolved.index > text.length)
      return invalid('invalid-bound');
    return {
      sourceId: cut.sourceId,
      text,
      index: resolved.index,
      ...(cut.continuationId === undefined ? {} : { continuationId: cut.continuationId }),
      ...(cut.afterAtomId === undefined ? {} : { afterAtomId: cut.afterAtomId }),
    };
  }

  atomIds(sourceId: string): readonly string[] {
    return [...(this.atomsBySource.get(sourceId) ?? [])];
  }
  hasSource(sourceId: string): boolean {
    return this.atomsBySource.has(sourceId);
  }
  noteSource(sourceId: string): void {
    if (this.hasSource(sourceId)) this.pendingSources.add(sourceId);
  }
  dirtySources(): readonly string[] {
    const sources = [...this.pendingSources];
    this.pendingSources.clear();
    return sources;
  }
  noteChanged(id: string): void {
    this.indexExisting(id);
    if (this.atomsBySource.has(id)) this.pendingSources.add(id);
  }
  indexExisting(id: string): void {
    const next = this.read(id)?.sourceId;
    const previous = this.sourceByAtom.get(id);
    if (previous !== next && previous !== undefined) {
      const atoms = this.atomsBySource.get(previous);
      atoms?.delete(id);
      if (atoms?.size === 0) this.atomsBySource.delete(previous);
      this.sourceByAtom.delete(id);
      this.pendingSources.add(previous);
    }
    if (next !== undefined) {
      const atoms = this.atomsBySource.get(next) ?? new Set<string>();
      atoms.add(id);
      this.atomsBySource.set(next, atoms);
      this.sourceByAtom.set(id, next);
      this.pendingSources.add(next);
    }
  }
  reset(): void {
    this.sourceByAtom.clear();
    this.atomsBySource.clear();
    this.pendingSources.clear();
    this.nodes.forEach((_record, id) => this.indexExisting(id));
  }

  /** Publish revived character identities after undo, so cold peers use the same boundary. */
  normalizeRestoredAnchors(): void {
    const writes: { id: string; encoded: string }[] = [];
    for (const id of this.sourceByAtom.keys()) {
      const cut = this.decode(id);
      if (!cut || cut.anchor.item === null) continue;
      const original = cut.anchor.item;
      let target = original,
        item = Y.getItem(this.doc.store, target);
      while (item instanceof Y.Item && item.redone !== null) {
        target = Y.createID(item.redone.client, item.redone.clock + target.clock - item.id.clock);
        item = Y.getItem(this.doc.store, target);
      }
      if (!(item instanceof Y.Item) || item.deleted || Y.compareIDs(original, target)) continue;
      const encoded = JSON.stringify({
        sourceId: cut.sourceId,
        ...(cut.continuationId === undefined ? {} : { continuationId: cut.continuationId }),
        ...(cut.afterAtomId === undefined ? {} : { afterAtomId: cut.afterAtomId }),
        cut: {
          ...Y.relativePositionToJSON(cut.anchor),
          item: { client: target.client, clock: target.clock },
        },
      });
      if (encoded.length > this.limits.maxStringLength) return invalid('invalid-string');
      writes.push({ id, encoded });
    }
    if (writes.length === 0) return;
    this.doc.transact(() => {
      for (const write of writes) {
        this.nodes.get(write.id)!.set(NODE_INLINE_SOURCE_CUT_FIELD, write.encoded);
        this.indexExisting(write.id);
      }
    }, RESTORED_CUT_ORIGIN);
  }

  private sourceText(id: string): Y.Text {
    const node = this.nodes.get(id),
      text = isNodeMap(node) ? node.get(NODE_TEXT_FIELD) : null;
    if (
      !(text instanceof Y.Text) ||
      text.doc !== this.doc ||
      text.length > this.limits.maxTextLength
    )
      return invalid('invalid-bound');
    return text;
  }
  private decode(id: string): Cut | null {
    const node = this.nodes.get(id);
    if (!isNodeMap(node)) return null;
    const encoded = node.get(NODE_INLINE_SOURCE_CUT_FIELD);
    if (encoded === undefined) return null;
    const shell = readNodeShell(node);
    if (
      !(
        (shell.kind === 'hardBreak' && shell.localName === 'br') ||
        (shell.kind === 'tab' && shell.localName === 'tab') ||
        (shell.kind === 'text' && shell.localName === 't')
      )
    )
      return invalid('invalid-bound');
    if (typeof encoded !== 'string' || encoded.length > this.limits.maxStringLength)
      return invalid('invalid-string');
    let raw: unknown;
    try {
      raw = JSON.parse(encoded);
    } catch {
      return invalid('invalid-string');
    }
    if (
      !object(raw) ||
      Object.keys(raw).some(
        (key) =>
          key !== 'sourceId' && key !== 'cut' && key !== 'continuationId' && key !== 'afterAtomId'
      ) ||
      typeof raw.sourceId !== 'string' ||
      raw.sourceId.length === 0 ||
      raw.sourceId.length > this.limits.maxStringLength ||
      rejectDangerousKey(raw.sourceId)
    )
      return invalid('invalid-string');
    if (
      raw.continuationId !== undefined &&
      (typeof raw.continuationId !== 'string' ||
        raw.continuationId.length === 0 ||
        raw.continuationId.length > this.limits.maxStringLength ||
        rejectDangerousKey(raw.continuationId))
    )
      return invalid('invalid-string');
    if (
      raw.afterAtomId !== undefined &&
      (typeof raw.afterAtomId !== 'string' ||
        raw.afterAtomId === id ||
        raw.afterAtomId.length === 0 ||
        raw.afterAtomId.length > this.limits.maxStringLength ||
        rejectDangerousKey(raw.afterAtomId))
    )
      return invalid('invalid-string');
    return {
      sourceId: raw.sourceId,
      anchor: position(raw.cut),
      ...(raw.continuationId === undefined ? {} : { continuationId: raw.continuationId as string }),
      ...(raw.afterAtomId === undefined ? {} : { afterAtomId: raw.afterAtomId as string }),
    };
  }
}
