/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { applyRegistryNodeEvents } from './registry-node-events.ts';
import {
  authoredInlineSourceRange,
  noteInlineSourceOwner,
} from './registry-inline-source-reads.ts';
import { observeDirtyPaths } from './registry-dirty-paths.ts';
import { captureInsertionBoundary } from './split-text-boundaries.ts';
import { SplitTextSources, type SplitTextRange } from './split-text-sources.ts';
import { InlineSourceCuts, type InlineSourceCut } from './inline-source-cuts.ts';
import { InlineSourcePlacement } from './inline-source-placement.ts';
import * as Y from 'yjs';
import type { CanonicalBinaryDescriptor } from '@docx-editor.dev/core/collaboration/replication';
import { partNameKey } from '@docx-editor.dev/core/store';
import { type LogicalId, type NodeIdentityMeta, wordFacingIdsOf } from './identity.ts';
import { SplitDedupIndex, type SplitTextOverlays } from './split-dedup.ts';
import { runIsPresent } from './run-text-reads.ts';
import {
  DEFAULT_DOCUMENT_LIMITS,
  mergeLimits,
  rejectDangerousKey,
  type DocumentLimits,
} from './limits.ts';
import {
  NODE_CHILDREN_FIELD,
  NODE_DELETED_FIELD,
  NODE_REPLACED_BY_FIELD,
  NODE_SHELL_FIELD,
  NODE_TEXT_FIELD,
  childArrayOf,
  internNamespace,
  namespaceIdOf,
  isElementRecord,
  isNodeMap,
  isTextNodeMap,
  makeBinaryEntry,
  makeElementRecord,
  makePartEntry,
  makeRelationshipEntry,
  makeTextRecord,
  nodeRecordReplacedBy,
  nodeRecordSplitFrom,
  nodeRecordTombstoned,
  packNodeShell,
  packageSchemaOf,
  parseAttributeMapKey,
  parseBindingMapKey,
  type DirtyPaths,
  type ElementRecord,
  type EncodedAttribute,
  type EncodedBinding,
  type EncodedRelationship,
  type PackageSchema,
  type PartDirectoryEntry,
  type SharedRecord,
  type TextRecord,
} from './schema.ts';
import { readRelationships, relationshipKey } from './relationship-store.ts';
import { observeRegistrySchema } from './registry-observers.ts';
import {
  assignFirstReachableParents,
  resolveContestedPlacements,
} from './registry-contested-placement.ts';
import {
  applyAttributeMapEvent,
  applyBindingMapEvent,
  deleteSharedAttribute,
  deleteSharedBinding,
  upsertIndexedAttribute,
  upsertIndexedBinding,
  writeSharedAttribute,
  writeSharedBinding,
} from './registry-side-maps.ts';
import {
  elementRecordOf,
  itemKeyOf,
  asTrackedType,
  readString,
  nodeKindOf,
  nodeShapeOf,
  sameChildOrder,
  type NodeShape,
} from './registry-node-reads.ts';
import {
  readBinaries,
  readContentTypeDefaults,
  readContentTypeOverrides,
  readPartEntries,
} from './registry-package-reads.ts';
import { retainsNumberingInfrastructure } from './undo-numbering.ts';
import { nodeRecordDeleteFilter } from './undo-delete-filter.ts';

/** Child-ID arrays are the only replicated membership and order authority. */
export class DocumentRegistry {
  readonly schema: PackageSchema;
  readonly limits: DocumentLimits;
  /** First reachable preorder parent. Derived, never replicated. */
  private parentIndex = new Map<LogicalId, LogicalId>();
  /** Every child-array listing. Derived, never replicated. */
  private listings = new Map<LogicalId, Set<LogicalId>>();
  private childrenSnapshot = new Map<LogicalId, readonly LogicalId[]>();
  private adoptees = new Map<LogicalId, LogicalId[]>();
  private tombstoneSources = new Map<LogicalId, Set<LogicalId>>();
  /** Which survivor's source set lists each tombstone. Derived, never replicated. */
  private tombstoneSurvivor = new Map<LogicalId, LogicalId>();
  private attributesByNode = new Map<LogicalId, Map<string, EncodedAttribute>>();
  private bindingsByNode = new Map<LogicalId, Map<string, EncodedBinding>>();
  /** Deterministic dedup of concurrent format splits (#581). */
  private readonly splitDedup: SplitDedupIndex;
  private readonly splitTextSources: SplitTextSources;
  private readonly inlineSourceCuts: InlineSourceCuts;
  private readonly inlineSourcePlacement: InlineSourcePlacement;
  private bulkLoad = 0;
  private unobservedWrites = 0;
  /** Node total as of the last observed event batch. Negative means "not counted yet". */
  private nodeCountCache = -1;
  /** Nodes written since that batch. Yjs delivers a nested transaction's events late. */
  private pendingNodeAdds = 0;
  /** Relationship-record total as of the last decode. Negative means "not counted yet". */
  private relationshipCountCache = -1;
  /** Part-entry total as of the last read. Negative means "not counted yet". */
  private partCountCache = -1;

  private readonly stopObserving: () => void;

  constructor(
    readonly doc: Y.Doc,
    limits?: Partial<DocumentLimits>
  ) {
    this.schema = packageSchemaOf(doc);
    this.limits = mergeLimits(limits);
    this.splitDedup = new SplitDedupIndex(this.schema.nodes);
    this.splitTextSources = new SplitTextSources(this.schema.nodes, this.doc, this.limits);
    this.inlineSourceCuts = new InlineSourceCuts(this.schema.nodes, this.doc, this.limits);
    this.inlineSourcePlacement = new InlineSourcePlacement(this);
    this.stopObserving = observeRegistrySchema(this.schema, {
      onNodeEvents: (events) => {
        if (this.bulkLoad > 0) return;
        this.applyChildArrayEvents(events);
      },
      onAttributeEvent: (event) => {
        if (this.bulkLoad > 0) return;
        this.applyAttributeMapEvent(event);
      },
      onBindingEvent: (event) => {
        if (this.bulkLoad > 0) return;
        this.applyBindingMapEvent(event);
      },
      // Dropping a stale count is safe at any time, so these run even during a bulk load. The
      // deep observers catch a peer on an earlier build writing inside a nested owner map.
      onRelationshipChange: () => {
        this.relationshipCountCache = -1;
      },
      onPartChange: () => {
        this.partCountCache = -1;
      },
    });
  }

  /**
   * Detach this registry's observers from the shared document.
   *
   * The registry does not own `doc`, so teardown has to give the observers back. A registry
   * left observing outlives its consumer: every later transaction pays its handlers, and its
   * derived indexes retain the whole tree. `readCollaborationDocument` builds one registry
   * per call on a long-lived server document, so this detach is load-bearing.
   */
  destroy(): void {
    this.stopObserving();
  }

  beginBulkLoad(): void {
    this.bulkLoad += 1;
  }

  endBulkLoad(): void {
    this.bulkLoad = Math.max(0, this.bulkLoad - 1);
    if (this.bulkLoad === 0) this.rebuildDerivedIndexes();
  }

  trackedTypes(): readonly Y.AbstractType<unknown>[] {
    return [
      asTrackedType(this.schema.meta),
      asTrackedType(this.schema.nodes),
      asTrackedType(this.schema.parts),
      asTrackedType(this.schema.relationships),
      asTrackedType(this.schema.overrides),
      asTrackedType(this.schema.defaults),
      asTrackedType(this.schema.binaries),
      asTrackedType(this.schema.attributes),
      asTrackedType(this.schema.bindings),
    ];
  }

  /**
   * The `deleteFilter` an undo manager over {@link trackedTypes} has to be built with.
   *
   * Undo reverses `nodes.set(id, record)` by deleting the record, and a deleted record takes a
   * peer's concurrently typed characters with it, unreachably. See
   * {@link nodeRecordDeleteFilter}. Omitting this is silent data loss, not a missing nicety.
   */
  undoDeleteFilter(): (item: Y.Item) => boolean {
    const nodeFilter = nodeRecordDeleteFilter(this.schema.nodes);
    return (item) => !retainsNumberingInfrastructure(this.schema, item) && nodeFilter(item);
  }

  encodeSnapshot(): Uint8Array {
    return Y.encodeStateAsUpdate(this.doc);
  }

  encodeUpdate(remoteStateVector: Uint8Array): Uint8Array {
    return Y.encodeStateAsUpdate(this.doc, remoteStateVector);
  }

  partEntries(): readonly PartDirectoryEntry[] {
    return readPartEntries(this.schema.parts);
  }

  /**
   * How many XML parts the directory holds.
   *
   * {@link partEntries} decodes and sorts the whole directory. The journal's part cap read it
   * once per `putXmlPart` effect, so admitting N parts walked the directory N times. The
   * count is cached; any write to the part map, local or remote, drops it.
   */
  partCount(): number {
    if (this.partCountCache < 0) this.partCountCache = this.partEntries().length;
    return this.partCountCache;
  }

  /**
   * How many relationship records shared state holds.
   *
   * {@link relationships} decodes and sorts every record. The journal's relationship cap read
   * it once per `putRelationship` effect, so pasting N images decoded the map N times. The
   * count is cached; any write to the relationship map, local or remote, drops it.
   */
  relationshipCount(): number {
    if (this.relationshipCountCache < 0) {
      this.relationshipCountCache = this.relationships().length;
    }
    return this.relationshipCountCache;
  }

  mainDocumentPart(): string {
    return readString(this.schema.meta.get('mainDocumentPart'));
  }

  record(logicalId: LogicalId): SharedRecord | null {
    if (rejectDangerousKey(logicalId)) return null;
    const rec = this.schema.nodes.get(logicalId);
    // A peer can plant a scalar here; treat it as no such node (see #567).
    if (!isNodeMap(rec)) return null;
    if (isTextNodeMap(rec)) {
      const text = rec.get(NODE_TEXT_FIELD);
      const value = text instanceof Y.Text ? text.toString() : '';
      return { logicalId, kind: 'textValue', value } satisfies TextRecord;
    }
    return elementRecordOf(
      logicalId,
      rec,
      this.schema,
      [...(this.attributesByNode.get(logicalId)?.values() ?? [])],
      [...(this.bindingsByNode.get(logicalId)?.values() ?? [])]
    );
  }

  parentOf(logicalId: LogicalId): LogicalId | null {
    return this.parentIndex.get(logicalId) ?? null;
  }
  projectionParentOf(id: LogicalId): LogicalId | null {
    return this.inlineSourcePlacement.parentOf(id) ?? this.parentOf(id);
  }

  identityMeta(logicalId: LogicalId): NodeIdentityMeta | null {
    const rec = this.schema.nodes.get(logicalId);
    if (!rec) return null;
    const shared = this.record(logicalId);
    const attributes = shared && isElementRecord(shared) ? shared.attributes : [];
    return {
      logicalId,
      yjsItemKey: itemKeyOf(rec),
      wordFacingIds: wordFacingIdsOf(attributes),
    };
  }

  isTombstoned(logicalId: LogicalId): boolean {
    return nodeRecordTombstoned(this.schema.nodes.get(logicalId));
  }

  replacedByOf(logicalId: LogicalId): LogicalId | null {
    return nodeRecordReplacedBy(this.schema.nodes.get(logicalId));
  }

  adoptedChildren(survivorId: LogicalId): readonly LogicalId[] {
    return this.adoptees.get(survivorId) ?? [];
  }

  /**
   * Every survivor that adopts children from a tombstone, and what it adopts.
   *
   * Adoption is DERIVED from the tombstone edge and is not a child array, so tombstoning a
   * node dirties the tombstone alone while the survivor is the node that has to grow a
   * child. A consumer that caches per node has to compare this index, or the adopted
   * children silently disappear from the survivor.
   */
  adoptionIndex(): ReadonlyMap<LogicalId, readonly LogicalId[]> {
    return this.adoptees;
  }

  /**
   * True when shared state has ever interned this namespace URI.
   *
   * Every element and attribute interns its namespace on write, so a URI absent from the
   * table cannot appear anywhere in the document. That answers "does this document use
   * feature X at all" without a walk over every node.
   */
  hasNamespace(uri: string): boolean {
    return this.schema.namespaces.get(namespaceIdOf(uri)) === uri;
  }

  /**
   * Say that shared state was written through this registry.
   *
   * Every derived index here, and every dirty set an observer builds, is maintained from Yjs
   * EVENTS. Yjs delivers the events of a transaction opened during another transaction's
   * cleanup only after that cleanup finishes — which is exactly what a queued local journal
   * flushed on arrival of a remote update does. Between the write and the delivery, shared
   * state holds the edit and every index still describes the state before it. A consumer that
   * trusts the indexes in that window drops the edit.
   */
  noteWrite(): void {
    this.unobservedWrites += 1;
  }

  /** True while a write through this registry is waiting for its Yjs events. */
  hasUnobservedWrites(): boolean {
    return this.unobservedWrites > 0;
  }

  listingParents(logicalId: LogicalId): readonly LogicalId[] {
    return [...(this.listings.get(logicalId) ?? [])];
  }

  allLogicalIds(): readonly LogicalId[] {
    return [...this.schema.nodes.keys()].filter((key) => !rejectDangerousKey(key));
  }

  /**
   * How many nodes shared state holds.
   *
   * `Y.Map.size` walks every key and allocates an array to measure it, and the journal's
   * node cap reads this once per commit. That made a keystroke cost the whole document:
   * 630us on a 200-page fixture, against a whole attached commit of about 4ms. The count is
   * maintained instead, from the same events every other derived index here is built from,
   * plus the writes those events have not described yet.
   */
  nodeCount(): number {
    if (this.nodeCountCache < 0) this.nodeCountCache = this.schema.nodes.size;
    return this.nodeCountCache + this.pendingNodeAdds;
  }

  /** The quantities a bound check reads, without the attribute arrays {@link record} builds. */
  nodeShape(logicalId: LogicalId): NodeShape | null {
    if (rejectDangerousKey(logicalId)) return null;
    return nodeShapeOf(this.schema.nodes, logicalId);
  }

  /** One node's kind, without building its text. */
  kindOf(logicalId: LogicalId): string | null {
    if (rejectDangerousKey(logicalId)) return null;
    return nodeKindOf(this.schema.nodes, logicalId);
  }

  private noteNodeWrite(logicalId: LogicalId): void {
    if (!this.schema.nodes.has(logicalId)) this.pendingNodeAdds += 1;
  }

  putElement(record: Omit<ElementRecord, 'childIds'>): void {
    this.noteNodeWrite(record.logicalId);
    const namespaceId = internNamespace(
      this.schema.namespaces,
      record.namespaceUri,
      this.limits.maxStringLength
    );
    this.schema.nodes.set(
      record.logicalId,
      makeElementRecord({
        logicalId: record.logicalId,
        kind: record.kind,
        namespaceUri: record.namespaceUri,
        namespaceId,
        localName: record.localName,
        prefix: record.prefix,
      })
    );
    for (const attribute of record.attributes) {
      writeSharedAttribute(
        this.schema,
        this.limits.maxStringLength,
        record.logicalId,
        attribute,
        attribute.value
      );
    }
    for (const binding of record.bindings) {
      writeSharedBinding(
        this.schema,
        this.limits.maxStringLength,
        record.logicalId,
        binding.prefix,
        binding.namespaceUri
      );
    }
  }

  /**
   * Rename one existing element in place.
   *
   * A note conversion changes `w:footnoteReference` to `w:endnoteReference` and keeps the
   * node, its children and its attributes. Replacing the record instead would drop the
   * children, and minting a new id would leave the old element in the tree unchanged.
   */
  updateElementShell(
    logicalId: LogicalId,
    shell: {
      readonly kind: string;
      readonly namespaceUri: string;
      readonly localName: string;
      readonly prefix?: string;
    }
  ): void {
    const rec = this.require(logicalId);
    const namespaceId = internNamespace(
      this.schema.namespaces,
      shell.namespaceUri,
      this.limits.maxStringLength
    );
    rec.set(
      NODE_SHELL_FIELD,
      packNodeShell(shell.kind, namespaceId, shell.localName, shell.prefix ?? '')
    );
  }

  putText(logicalId: LogicalId, value: string): void {
    this.noteNodeWrite(logicalId);
    this.schema.nodes.set(logicalId, makeTextRecord(value));
  }

  spliceText(logicalId: LogicalId, utf16Start: number, deleteCount: number, insert: string): void {
    const range = this.splitTextSources.range(logicalId);
    const restoreBoundary = captureInsertionBoundary(
      this.schema.nodes,
      this.splitTextSources,
      logicalId,
      range,
      utf16Start,
      insert
    );
    const text = range?.text ?? this.textOf(logicalId);
    const start = (range?.start ?? 0) + utf16Start;
    if (deleteCount > 0) text.delete(start, deleteCount);
    if (insert.length > 0) text.insert(start, insert);
    restoreBoundary?.();
  }

  setAttribute(
    logicalId: LogicalId,
    attribute: {
      readonly namespaceUri: string;
      readonly localName: string;
      readonly prefix?: string;
    },
    value: string | null
  ): void {
    if (value === null) {
      deleteSharedAttribute(this.schema, logicalId, attribute.namespaceUri, attribute.localName);
      return;
    }
    writeSharedAttribute(this.schema, this.limits.maxStringLength, logicalId, attribute, value);
  }

  setNamespaceBinding(logicalId: LogicalId, prefix: string, uri: string | null): void {
    if (uri === null) {
      deleteSharedBinding(this.schema, logicalId, prefix);
      return;
    }
    writeSharedBinding(this.schema, this.limits.maxStringLength, logicalId, prefix, uri);
  }

  spliceChildren(
    parentId: LogicalId,
    index: number,
    deleteCount: number,
    insertIds: readonly LogicalId[]
  ): void {
    const children = this.childArray(parentId);
    if (deleteCount > 0) children.delete(index, deleteCount);
    if (insertIds.length > 0) children.insert(index, [...insertIds]);
  }

  moveNode(nodeId: LogicalId, destParentId: LogicalId, destIndex: number): void {
    this.unlinkFromAllParents(nodeId);
    const dest = this.childArray(destParentId);
    dest.insert(Math.max(0, Math.min(destIndex, dest.length)), [nodeId]);
  }

  tombstone(logicalId: LogicalId, replacedBy?: LogicalId): void {
    this.unlinkFromAllParents(logicalId);
    const rec = this.require(logicalId);
    rec.set(NODE_DELETED_FIELD, true);
    if (replacedBy) rec.set(NODE_REPLACED_BY_FIELD, replacedBy);
  }

  recordRunSplit(root: LogicalId, replaced: LogicalId, runs: readonly LogicalId[]): void {
    this.splitDedup.record(root, replaced, runs);
  }

  /** The run a given run split off from, or null — the caller resolves the chain (#581). */
  splitOriginOf(logicalId: LogicalId): LogicalId | null {
    return nodeRecordSplitFrom(this.schema.nodes.get(logicalId));
  }

  /** Runs a concurrent format split superseded and this replica must not materialize (#581). */
  replacementLoserRuns(): ReadonlySet<LogicalId> {
    return this.inlineSourcePlacement.visibleLosers(this.rawReplacementLosers());
  }

  rawReplacementLosers(): ReadonlySet<LogicalId> {
    if (this.hasUnobservedWrites()) this.splitDedup.invalidate();
    return this.splitDedup.loserRuns({ isPresent: (id) => runIsPresent(this, id) });
  }

  splitTextRange(id: LogicalId): SplitTextRange | null {
    const range = this.splitTextSources.range(id);
    return range ? { ...range, end: this.inlineSourcePlacement.effectiveEnd(id, range.end) } : null;
  }

  authoredSourceRange(id: LogicalId): SplitTextRange | null {
    return authoredInlineSourceRange(this, this.splitTextSources, id);
  }

  sourceAliases(source: string): readonly string[] {
    return [source, ...this.splitTextSources.sourceAliases(source)];
  }

  listedParents(id: string): readonly string[] {
    return [...(this.listings.get(id) ?? [])];
  }

  projectedTextValue(id: LogicalId): string | null {
    return this.inlineSourcePlacement.textValue(id, this.splitTextSources.value(id));
  }

  projectedInlineChildren(id: LogicalId, children: readonly LogicalId[]): readonly LogicalId[] {
    const extras = this.adoptedChildren(id);
    const relevant =
      this.inlineSourcePlacement.hasChildren(id) ||
      extras.some((child) => {
        if (this.inlineSourceCut(child)) return true;
        const shape = this.nodeShape(child);
        return (
          shape?.children.some((leaf) => {
            const source = this.authoredSourceRange(leaf)?.sourceId;
            return source !== undefined && this.inlineSourceCuts.hasSource(source);
          }) ?? false
        );
      });
    if (!relevant) return this.inlineSourcePlacement.children(id, children);
    const merged = [...children];
    for (const child of extras) if (!merged.includes(child)) merged.push(child);
    return this.inlineSourcePlacement.children(id, merged);
  }

  registerSplitText(product: LogicalId, source: LogicalId, start: number, end: number): void {
    this.splitTextSources.register(product, source, start, end);
  }

  registerInlineSourceCut(...args: Parameters<InlineSourceCuts['register']>): void {
    this.inlineSourceCuts.register(...args);
  }

  inlineSourceCut(atom: LogicalId): InlineSourceCut | null {
    return this.inlineSourceCuts.read(atom);
  }

  inlineAtomsForSource(source: string): readonly LogicalId[] {
    return this.inlineSourceCuts.atomIds(source);
  }

  normalizeRestoredSplitTextAnchors(): void {
    this.splitTextSources.normalizeRestoredAnchors();
    this.inlineSourceCuts.normalizeRestoredAnchors();
  }

  concurrentSplitTextOverlays(): SplitTextOverlays {
    const base = this.splitTextSources.overlays();
    return this.inlineSourcePlacement.overlays(base.values, base.changedIds);
  }

  /**
   * Drop from a node's child array the children this replica had already seen there.
   *
   * Adoption rescues what a CONCURRENT peer put inside a node this replica tombstoned. It must
   * not rescue what this replica's own edit superseded. A run split leaves the replaced `w:t`
   * listed under the run the edit dropped, and adopting it puts the pre-edit text back beside
   * the new text on every replica but the author's.
   *
   * Each id is deleted as many times as the caller saw it, newest occurrence first, so the
   * delete stays item-level: an id a peer inserted concurrently is a different Yjs item, keeps
   * its place in the array, and still reaches the survivor.
   */
  unlistChildren(parentId: LogicalId, childIds: readonly LogicalId[]): void {
    if (childIds.length === 0) return;
    const rec = this.schema.nodes.get(parentId);
    const children = rec ? childArrayOf(rec) : null;
    if (!children) return;
    const remaining = new Map<LogicalId, number>();
    for (const childId of childIds) remaining.set(childId, (remaining.get(childId) ?? 0) + 1);
    for (let index = children.length - 1; index >= 0; index -= 1) {
      const value = children.get(index);
      const left = remaining.get(value) ?? 0;
      if (left === 0) continue;
      remaining.set(value, left - 1);
      children.delete(index, 1);
    }
  }

  putXmlPart(entry: PartDirectoryEntry): void {
    // Dropped here as well as in the observer: a nested transaction delivers its events late,
    // and a queued journal can read the count inside that window.
    this.partCountCache = -1;
    this.schema.parts.set(
      entry.name,
      makePartEntry(entry.id, entry.rootLogicalId, entry.contentType)
    );
  }

  deleteXmlPart(name: string): void {
    this.partCountCache = -1;
    this.schema.parts.delete(name);
  }

  /**
   * Write one relationship under a key of its own.
   *
   * Keying by part alone meant two peers adding the FIRST relationship to the same part both
   * took the "create the owner map" branch, and a nested `Y.Map` at one key resolves
   * last-writer-wins: the loser's map went away with its rId inside it, leaving a broken image
   * or a dead hyperlink that no later edit could repair. Part AND id means concurrent writers
   * touch different keys and cannot collide. Two writers of the same id still resolve
   * last-writer-wins, which is the honest answer to one id with two targets.
   *
   * The value stays a map of entries, so a reader also sees owner maps written by a peer on an
   * earlier build.
   */
  putRelationship(record: EncodedRelationship): void {
    if (rejectDangerousKey(record.ownerPart) || rejectDangerousKey(record.id)) return;
    this.relationshipCountCache = -1;
    const holder = new Y.Map<Y.Map<unknown>>();
    holder.set(record.id, makeRelationshipEntry(record));
    this.schema.relationships.set(relationshipKey(record.ownerPart, record.id), holder);
  }

  deleteRelationship(ownerPart: string, relationshipId: string): void {
    this.relationshipCountCache = -1;
    this.schema.relationships.delete(relationshipKey(ownerPart, relationshipId));
    this.schema.relationships.get(ownerPart)?.delete(relationshipId);
  }

  putContentTypeOverride(partName: string, mediaType: string): void {
    this.schema.overrides.set(partNameKey(partName), mediaType);
  }

  deleteContentTypeOverride(partName: string): void {
    this.schema.overrides.delete(partNameKey(partName));
  }

  putContentTypeDefault(extension: string, mediaType: string): void {
    this.schema.defaults.set(extension, mediaType);
  }

  putBinary(descriptor: CanonicalBinaryDescriptor): void {
    this.schema.binaries.set(descriptor.storageKey, makeBinaryEntry(descriptor));
  }

  deleteBinary(storageKey: string): void {
    this.schema.binaries.delete(storageKey);
  }

  relationships(): readonly EncodedRelationship[] {
    return readRelationships(this.schema.relationships);
  }

  contentTypeOverrides(): ReadonlyMap<string, string> {
    return readContentTypeOverrides(this.schema.overrides);
  }

  contentTypeDefaults(): ReadonlyMap<string, string> {
    return readContentTypeDefaults(this.schema.defaults);
  }

  binaries(): readonly CanonicalBinaryDescriptor[] {
    return readBinaries(this.schema.binaries);
  }

  hasNode(logicalId: LogicalId): boolean {
    return this.schema.nodes.has(logicalId);
  }

  observeDirty(onDirty: (paths: DirtyPaths) => void): () => void {
    return observeDirtyPaths(this.schema, onDirty);
  }

  rebuildDerivedIndexes(): void {
    this.unobservedWrites = 0;
    this.nodeCountCache = -1;
    this.pendingNodeAdds = 0;
    this.relationshipCountCache = -1;
    this.partCountCache = -1;
    this.parentIndex = new Map();
    this.listings = new Map();
    this.childrenSnapshot = new Map();
    this.adoptees = new Map();
    this.tombstoneSources = new Map();
    this.tombstoneSurvivor = new Map();
    this.attributesByNode = new Map();
    this.bindingsByNode = new Map();
    this.splitDedup.reset();
    this.splitTextSources.reset();
    this.inlineSourceCuts.reset();
    this.inlineSourcePlacement.reset();
    this.schema.nodes.forEach((rec, parentId) => {
      if (rejectDangerousKey(parentId)) return;
      const children = childArrayOf(rec);
      if (!children) return;
      const childIds = children.toArray();
      this.childrenSnapshot.set(parentId, childIds);
      for (const childId of childIds) this.addListing(childId, parentId);
    });
    this.schema.nodes.forEach((rec, id) => {
      if (nodeRecordTombstoned(rec)) this.syncAdoptee(id);
      this.splitDedup.indexExisting(id);
    });
    this.schema.attributes.forEach((packed, key) => {
      if (typeof packed !== 'string' || rejectDangerousKey(key)) return;
      const parsed = parseAttributeMapKey(key);
      if (!parsed || rejectDangerousKey(parsed.logicalId) || rejectDangerousKey(parsed.localName)) {
        return;
      }
      upsertIndexedAttribute(
        this.schema,
        this.attributesByNode,
        parsed.logicalId,
        parsed.namespaceId,
        parsed.localName,
        packed
      );
    });
    this.schema.bindings.forEach((namespaceId, key) => {
      if (typeof namespaceId !== 'string' || rejectDangerousKey(key)) return;
      const parsed = parseBindingMapKey(key);
      if (!parsed || rejectDangerousKey(parsed.logicalId) || rejectDangerousKey(parsed.prefix)) {
        return;
      }
      upsertIndexedBinding(
        this.schema,
        this.bindingsByNode,
        parsed.logicalId,
        parsed.prefix,
        namespaceId
      );
    });
    this.assignFirstReachable(null);
    this.inlineSourcePlacement.reconcile(this.inlineSourceCuts.dirtySources());
  }

  assertNoParentFields(): void {
    this.schema.nodes.forEach((value) => {
      if (isNodeMap(value) && (value.has('parent') || value.has('parentId'))) {
        throw new Error('registry record must not replicate a parent field');
      }
    });
  }

  private applyChildArrayEvents(events: Y.YEvent<Y.AbstractType<unknown>>[]): void {
    this.unobservedWrites = 0;
    this.pendingNodeAdds = 0;
    const changed = applyRegistryNodeEvents(events, {
      splitDedup: this.splitDedup,
      splitTextSources: this.splitTextSources,
      inlineSourceCuts: this.inlineSourceCuts,
      noteInlineSourceOwner: (id) => noteInlineSourceOwner(this, this.inlineSourceCuts, id),
      noteNodeCount: (action) => {
        if (this.nodeCountCache >= 0 && action !== 'update')
          this.nodeCountCache += action === 'add' ? 1 : -1;
      },
      syncChildListings: (id) => this.syncChildListings(id),
      isTombstoned: (id) => this.isTombstoned(id),
      syncAdoptee: (id) => this.syncAdoptee(id),
    });
    if (changed.size > 0) this.resolveParents(changed);
    this.inlineSourcePlacement.reconcile(this.inlineSourceCuts.dirtySources());
  }

  private applyAttributeMapEvent(event: Y.YMapEvent<string>): void {
    this.unobservedWrites = 0;
    applyAttributeMapEvent(this.schema, this.attributesByNode, event);
  }

  private applyBindingMapEvent(event: Y.YMapEvent<string>): void {
    this.unobservedWrites = 0;
    applyBindingMapEvent(this.schema, this.bindingsByNode, event);
  }

  private syncChildListings(parentId: LogicalId): LogicalId[] {
    const rec = this.schema.nodes.get(parentId);
    const next = rec ? (childArrayOf(rec)?.toArray() ?? []) : [];
    const prev = this.childrenSnapshot.get(parentId) ?? [];
    // The top-level node map reports a whole record as one key change, so this runs for every
    // node a journal writes, whether or not that node's children moved. An unchanged listing
    // has nothing to say and no snapshot to replace.
    if (sameChildOrder(prev, next)) return [];
    // Membership by set, not by scan. The body root lists every block in the document, and a
    // journal now publishes on the commit that produces it, so this sits on the keystroke
    // path: a scan per child cost ~640,000 comparisons to append one block to a list of 800.
    const nextSet = new Set(next);
    const prevSet = new Set(prev);
    const affected: LogicalId[] = [];
    for (const childId of prev) {
      if (nextSet.has(childId)) continue;
      this.removeListing(childId, parentId);
      affected.push(childId);
    }
    for (const childId of next) {
      // A child the snapshot already listed here is already in `listings`, because the two
      // only ever move together. Re-adding it cost three map operations per sibling, so
      // appending one block to a list of 800 rewrote all 800 listings to change one.
      if (prevSet.has(childId)) continue;
      this.addListing(childId, parentId);
      affected.push(childId);
    }
    this.childrenSnapshot.set(parentId, next);
    return affected;
  }

  private addListing(childId: LogicalId, parentId: LogicalId): void {
    const listed = this.listings.get(childId) ?? new Set<LogicalId>();
    listed.add(parentId);
    this.listings.set(childId, listed);
  }

  private removeListing(childId: LogicalId, parentId: LogicalId): void {
    const listed = this.listings.get(childId);
    if (!listed) return;
    listed.delete(parentId);
    if (listed.size === 0) this.listings.delete(childId);
  }

  private resolveParents(ids: ReadonlySet<LogicalId>): void {
    const multi: LogicalId[] = [];
    for (const id of ids) {
      const listed = this.listings.get(id);
      if (!listed || listed.size === 0) {
        this.parentIndex.delete(id);
        continue;
      }
      if (listed.size === 1) {
        this.parentIndex.set(id, [...listed][0]!);
        continue;
      }
      multi.push(id);
    }
    if (multi.length === 0) return;
    const resolved = resolveContestedPlacements(
      {
        nodes: this.schema.nodes,
        parentIndex: this.parentIndex,
        listings: this.listings,
        childrenSnapshot: this.childrenSnapshot,
        partRoots: this.partEntries().map((entry) => entry.rootLogicalId),
      },
      multi
    );
    if (!resolved) this.assignFirstReachable(new Set(multi));
  }

  private assignFirstReachable(only: ReadonlySet<LogicalId> | null): void {
    assignFirstReachableParents(
      {
        nodes: this.schema.nodes,
        parentIndex: this.parentIndex,
        listings: this.listings,
        childrenSnapshot: this.childrenSnapshot,
        partRoots: this.partEntries().map((entry) => entry.rootLogicalId),
      },
      only
    );
  }

  private syncAdoptee(removedId: LogicalId): void {
    // A tombstone lists at most one survivor, and every listing is written here, so the
    // reverse index names the one source set that can hold `removedId`. Scanning every
    // survivor instead made each tombstone event cost every tombstone in the session.
    const previous = this.tombstoneSurvivor.get(removedId);
    if (previous !== undefined) {
      this.tombstoneSurvivor.delete(removedId);
      const sources = this.tombstoneSources.get(previous);
      if (sources) {
        sources.delete(removedId);
        if (sources.size === 0) this.tombstoneSources.delete(previous);
      }
    }
    const rec = this.schema.nodes.get(removedId);
    const survivor = nodeRecordReplacedBy(rec);
    if (survivor === null) return;
    if (nodeRecordTombstoned(rec)) {
      const sources = this.tombstoneSources.get(survivor) ?? new Set<LogicalId>();
      sources.add(removedId);
      this.tombstoneSources.set(survivor, sources);
      this.tombstoneSurvivor.set(removedId, survivor);
    }
    // An un-tombstoned node keeps its `replacedBy`, and its former survivor has to give the
    // adopted children back, so the survivor recomputes in both branches.
    this.recomputeAdoptees(survivor);
  }

  private isContentWitness(id: LogicalId): boolean {
    const kind = this.kindOf(id);
    return kind !== null && !kind.endsWith('Properties');
  }

  private recomputeAdoptees(survivorId: LogicalId): void {
    const extras: LogicalId[] = [];
    for (const tombstoneId of this.tombstoneSources.get(survivorId) ?? []) {
      const rec = this.schema.nodes.get(tombstoneId);
      if (!rec) continue;
      for (const childId of childArrayOf(rec)?.toArray() ?? []) {
        if (this.isTombstoned(childId) || !this.schema.nodes.has(childId)) continue;
        // A join drops the removed paragraph's `w:pPr`. Adopting that property node onto the
        // survivor produced two `paragraphProperties` children, which `known-node-invariant`
        // refused, so the receiving replica never installed the joined tree.
        if (!this.isContentWitness(childId)) continue;
        extras.push(childId);
      }
    }
    if (extras.length > 0) this.adoptees.set(survivorId, extras);
    else this.adoptees.delete(survivorId);
  }

  unlinkFromAllParents(id: LogicalId): void {
    const parents = [...(this.listings.get(id) ?? [])];
    if (parents.length === 0) {
      this.schema.nodes.forEach((rec) => {
        const children = childArrayOf(rec);
        if (children) this.removeFromArray(children, id);
      });
      return;
    }
    for (const parentId of parents) {
      const rec = this.schema.nodes.get(parentId);
      const children = rec ? childArrayOf(rec) : null;
      if (children) this.removeFromArray(children, id);
    }
  }

  private removeFromArray(array: Y.Array<string>, id: string): void {
    for (let index = array.length - 1; index >= 0; index -= 1) {
      if (array.get(index) === id) array.delete(index, 1);
    }
  }

  textOf(logicalId: LogicalId): Y.Text {
    const text = this.require(logicalId).get(NODE_TEXT_FIELD);
    if (!(text instanceof Y.Text)) throw new Error(`no text at ${logicalId}`);
    return text;
  }

  childArray(logicalId: LogicalId): Y.Array<string> {
    const children = this.require(logicalId).get(NODE_CHILDREN_FIELD);
    if (!(children instanceof Y.Array)) throw new Error(`no children at ${logicalId}`);
    return children;
  }

  private require(logicalId: LogicalId): Y.Map<unknown> {
    const rec = this.schema.nodes.get(logicalId);
    if (!rec) throw new Error(`registry node ${logicalId} missing`);
    return rec;
  }
}

export const DOCUMENT_LIMITS = DEFAULT_DOCUMENT_LIMITS;
