// Carry bounded merged text between physical row groups without changing the tree.
// The carried record keeps the authored head row and cell identity on every page.
import type { OoxmlNode } from '@docx-editor.dev/core/store';
import type {
  SemanticTableCell,
  SemanticTableRow,
  SemanticTableStructure,
} from './semantic-table.ts';
import type {
  BlockFragmentRecord,
  TableCellFragmentRecord,
  TableRowFragmentRecord,
} from './semantic-records.ts';
import {
  initialCellCursors,
  layoutRowFragmentBounded,
  measureRowHeight,
  TablePaginationError,
  type CellPlaceCursor,
  type TableFlowDeps,
} from './semantic-table-layout.ts';
import type { CellContentInsets } from './table-cell-geometry.ts';
import type { RowVMergeLayoutOptions } from './table-vmerge-heights.ts';
import { resolveVMergeSpans } from './table-vmerge.ts';
import { stripAnchorSinksForProbe } from './table-probe-deps.ts';
import { withoutZeroHeightBottomEdge } from './table-vmerge-boundary.ts';
import { isCarriedHeadRow } from './table-carried-head-row.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const EPSILON = 0.001;
// Notes, drawings, nested tables, breaks, and unknown extensions stay on the existing path.
// Field instructions remain inert and use the native projection and its atomic source range.
const TEXT_NAMES = new Set([
  'p',
  'pPr',
  'r',
  'rPr',
  't',
  'spacing',
  'jc',
  'ind',
  'rFonts',
  'sz',
  'szCs',
  'lang',
  'b',
  'bCs',
  'i',
  'iCs',
  'u',
  'color',
  'highlight',
  'shd',
  'vertAlign',
  'kern',
  'position',
  'w',
  'fitText',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'vanish',
  'noProof',
  'rtl',
  'cs',
  'pStyle',
  'rStyle',
  'contextualSpacing',
  'widowControl',
  'fldChar',
  'instrText',
  'fldSimple',
]);

function textOnly(cell: SemanticTableCell, budget: { remaining: number }): boolean {
  if (cell.textDirection !== 'horizontal') return false;
  const pending: OoxmlNode[] = [...cell.blocks];
  while (pending.length > 0) {
    if (--budget.remaining < 0) return false;
    const node = pending.pop()!;
    if (node.kind === 'textValue') continue;
    if (!('localName' in node) || node.namespaceUri !== W || !TEXT_NAMES.has(node.localName))
      return false;
    for (const child of node.children) pending.push(child);
  }
  return true;
}

interface Head {
  readonly cell: SemanticTableCell;
  cursor: CellPlaceCursor | null;
  complete: boolean;
}

interface Group {
  readonly first: number;
  readonly last: number;
  readonly source: SemanticTableRow;
  readonly heads: Head[];
  heights: number[];
  active: boolean;
  headRecord?: TableRowFragmentRecord;
}

export interface VMergeFragmentCarry {
  rowAt(index: number, row: SemanticTableRow, top: number, bottom: number): SemanticTableRow;
  optionsAt(index: number): RowVMergeLayoutOptions | undefined;
  finish(rows: TableRowFragmentRecord[]): TableRowFragmentRecord[];
  publish(
    records: TableRowFragmentRecord[],
    sources: SemanticTableRow[],
    insets: Map<TableRowFragmentRecord, ReadonlyMap<string, CellContentInsets>>,
    left: number
  ): { rows: TableRowFragmentRecord[]; sources: SemanticTableRow[] };
}

/** A protected, text-only merge whose rows can cross a page while its text uses a cursor. */
export function createVMergeFragmentCarry(
  structure: SemanticTableStructure,
  rows: readonly SemanticTableRow[],
  deps: TableFlowDeps,
  left: () => number
): VMergeFragmentCarry {
  const groups: Group[] = [];
  const atRow = new Map<number, Group>();
  const rowIndex = new Map(rows.map((row, index) => [row.id, index]));
  const spans = resolveVMergeSpans(rows);
  const grouped = new Map<string, Group>();
  let work = 0;
  for (const [index, row] of rows.entries()) {
    for (const cell of row.cells) {
      const count = spans.get(cell.id) ?? 1;
      if (count < 2 || cell.vMergeContinue) continue;
      work += count;
      if (work > 100_000 || grouped.size >= 4096) break;
      const last = index + count - 1;
      if (last >= rows.length) continue;
      const key = `${index}:${last}`;
      let group = grouped.get(key);
      if (!group) {
        group = { first: index, last, source: row, heads: [], heights: [], active: false };
        grouped.set(key, group);
      }
      group.heads.push({ cell, cursor: null, complete: false });
    }
    if (work > 100_000 || grouped.size >= 4096) break;
  }
  if (
    work <= 100_000 &&
    grouped.size < 4096 &&
    !structure.float &&
    !deps.pageExclusionZones?.().length
  ) {
    const candidates = [...grouped.values()].sort((a, b) => a.first - b.first);
    const nodeBudget = { remaining: 100_000 };
    let previousEnd = -1;
    for (const [candidateIndex, group] of candidates.entries()) {
      const overlapsBefore = previousEnd >= group.first;
      previousEnd = Math.max(previousEnd, group.last);
      if (group.last - group.first < 2) continue;
      if (overlapsBefore || (candidates[candidateIndex + 1]?.first ?? Infinity) <= group.last)
        continue;
      if (!group.heads.every((head) => textOnly(head.cell, nodeBudget))) continue;
      let valid = true;
      for (let index = group.first; index <= group.last; index++) {
        const row = rows[index]!;
        if (row.isHeader || (!row.cantSplit && row.height.rule !== 'exact')) {
          valid = false;
          break;
        }
        const cells = new Map<number, SemanticTableCell>();
        for (const cell of row.cells) {
          if (cells.has(cell.gridColumn)) {
            valid = false;
            break;
          }
          cells.set(cell.gridColumn, cell);
        }
        for (const head of group.heads) {
          const matching = cells.get(head.cell.gridColumn);
          if (
            !matching ||
            matching.gridSpan !== head.cell.gridSpan ||
            (index > group.first && !matching.vMergeContinue)
          ) {
            valid = false;
            break;
          }
        }
      }
      if (!valid) continue;
      groups.push(group);
      for (let index = group.first; index <= group.last; index++) atRow.set(index, group);
    }
  }

  function emptied(group: Group, row: SemanticTableRow): SemanticTableRow {
    const headIds = new Set(group.heads.map((head) => head.cell.id));
    return {
      ...row,
      cells: row.cells.map((cell) =>
        headIds.has(cell.id) ? { ...cell, blocks: [], vMergeContinue: true } : cell
      ),
    };
  }

  function prepare(group: Group): boolean {
    let lineId = 0;
    const probeDeps: TableFlowDeps = {
      ...stripAnchorSinksForProbe(deps),
      cache: undefined,
      borderOwnershipBudget: undefined,
      vMergeResolveBudget: undefined,
      nextLineId: () => `probe-vmerge-fragment-${lineId++}`,
    };
    const cols = structure.columnWidthsPt;
    const spacing = structure.cellSpacingPt;
    const heights: number[] = [];
    for (let index = group.first; index <= group.last; index++) {
      heights.push(
        measureRowHeight(emptied(group, rows[index]!), cols, left(), 0, probeDeps, spacing)
      );
    }
    if (heights.some((height) => !Number.isFinite(height) || height <= EPSILON)) return false;
    let total = heights.reduce((sum, height) => sum + height, 0);
    for (const head of group.heads) {
      const needed = measureRowHeight(
        { ...group.source, height: { rule: 'auto' }, cells: [head.cell] },
        cols,
        left(),
        0,
        probeDeps,
        spacing
      );
      if (!Number.isFinite(needed) || needed < 0) return false;
      if (needed > total + EPSILON) {
        let grow = group.last;
        while (grow >= group.first && rows[grow]!.height.rule === 'exact') grow--;
        // Exact spans that cannot hold the whole head keep the authored clipping behavior.
        if (grow < group.first) return false;
        heights[grow - group.first]! += needed - total;
        total = needed;
      }
    }
    group.heights = heights;
    return true;
  }

  function placeHead(
    group: Group,
    head: Head,
    first: TableRowFragmentRecord,
    bottom: number,
    x: number
  ): TableCellFragmentRecord {
    const sourceCell = { ...head.cell, vAlign: 'top' as const };
    const row: SemanticTableRow = {
      ...group.source,
      height: { rule: 'auto' },
      cantSplit: false,
      cells: [sourceCell],
    };
    const cursors = initialCellCursors(row);
    if (head.cursor) cursors[0] = head.cursor;
    let blocks: readonly BlockFragmentRecord[] = [];
    if (!head.complete) {
      const placed = layoutRowFragmentBounded(
        row,
        structure.columnWidthsPt,
        x,
        first.box.y,
        bottom,
        false,
        head.cursor !== null,
        0,
        deps,
        cursors,
        structure.cellSpacingPt
      );
      const cell = placed.record.cells[0]!;
      blocks = cell.blocks;
      head.complete = placed.remainder === null;
      head.cursor = placed.remainder?.[0] ?? null;
      if (!head.complete && !head.cursor)
        throw new TablePaginationError(
          'table-row-split-unsupported',
          'Merged text lost its continuation cursor'
        );
    }
    const original = group.headRecord!.cells.find((cell) => cell.id === head.cell.id)!;
    return {
      ...original,
      blocks,
      vMergeContinue: false,
      paintInert: false,
      box: {
        ...original.box,
        x: x + original.box.x - group.headRecord!.box.x,
        y: first.box.y,
        height: bottom - first.box.y,
      },
    };
  }

  return {
    rowAt(index, row, top, bottom) {
      const group = atRow.get(index);
      if (!group) return row;
      if (index === group.first && !group.active && prepare(group)) {
        const total = group.heights.reduce((sum, height) => sum + height, 0);
        group.active =
          top + total > bottom + EPSILON &&
          group.heights.every((height) => height <= bottom + EPSILON);
      }
      return group.active ? emptied(group, row) : row;
    },
    optionsAt(index) {
      const group = atRow.get(index);
      return group?.active ? { heightFloorPt: group.heights[index - group.first] } : undefined;
    },
    finish(records) {
      if (!groups.some((group) => group.active)) return records;
      return records.map((row) =>
        isCarriedHeadRow(row) ? { ...row, cells: row.cells.map(withoutZeroHeightBottomEdge) } : row
      );
    },
    publish(records, sources, insets, x) {
      if (!groups.some((group) => group.active)) return { rows: records, sources };
      const insertions = new Map<
        number,
        { record: TableRowFragmentRecord; source: SemanticTableRow }[]
      >();
      const replacements = new Map<number, TableRowFragmentRecord>();
      const coveredByGroup = new Map<Group, number[]>();
      for (const [index, record] of records.entries()) {
        const authored = rowIndex.get(record.id);
        const group = authored === undefined ? undefined : atRow.get(authored);
        if (!record.isHeaderRepeat && group?.active) {
          const covered = coveredByGroup.get(group) ?? [];
          covered.push(index);
          coveredByGroup.set(group, covered);
        }
      }
      for (const group of groups) {
        if (!group.active) continue;
        const covered = coveredByGroup.get(group) ?? [];
        if (!covered.length) continue;
        const firstIndex = covered[0]!;
        const first = records[firstIndex]!;
        const last = records[covered[covered.length - 1]!]!;
        if (!group.headRecord) {
          if (first.id !== group.source.id)
            throw new TablePaginationError(
              'table-row-split-unsupported',
              'Merged text has no placed head row'
            );
          group.headRecord = first;
        }
        const bottom = last.box.y + last.box.height;
        const placed = new Map(
          group.heads.map((head) => [head.cell.id, placeHead(group, head, first, bottom, x)])
        );
        const headRecord: TableRowFragmentRecord = {
          ...group.headRecord,
          ...(first.id === group.source.id ? {} : { isContinuation: true }),
          box: first.id === group.source.id ? first.box : { ...first.box, height: 0 },
          cells: group.headRecord.cells.map(
            (cell) =>
              placed.get(cell.id) ?? {
                ...cell,
                blocks: [],
                box: {
                  ...cell.box,
                  x: cell.box.x + x - group.headRecord!.box.x,
                  y: first.box.y,
                  height: 0,
                },
              }
          ),
        };
        if (first.id === group.source.id) {
          // Nonmerged head-row cells were already placed and keep their original records.
          replacements.set(firstIndex, {
            ...first,
            cells: first.cells.map((cell) => placed.get(cell.id) ?? cell),
          });
        } else {
          const pending = insertions.get(firstIndex) ?? [];
          pending.push({ record: headRecord, source: group.source });
          insertions.set(firstIndex, pending);
          const originalInsets = insets.get(first);
          if (originalInsets) insets.set(headRecord, originalInsets);
        }
        if (rowIndex.get(last.id) === group.last && group.heads.some((head) => !head.complete)) {
          throw new TablePaginationError(
            'table-row-split-unsupported',
            'Merged text remains after its authored row span'
          );
        }
      }
      const outRows: TableRowFragmentRecord[] = [];
      const outSources: SemanticTableRow[] = [];
      for (const [index, record] of records.entries()) {
        for (const insertion of insertions.get(index) ?? []) {
          outRows.push(insertion.record);
          outSources.push(insertion.source);
        }
        const replacement = replacements.get(index) ?? record;
        outRows.push(replacement);
        const sourceIndex = rowIndex.get(record.id);
        outSources.push(
          replacements.has(index) && sourceIndex !== undefined
            ? rows[sourceIndex]!
            : sources[index]!
        );
        const originalInsets = insets.get(record);
        if (replacement !== record && originalInsets) insets.set(replacement, originalInsets);
      }
      return { rows: outRows, sources: outSources };
    },
  };
}
