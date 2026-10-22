// Reserve the supported terminal exact row's outer rule outside its authored content box.
import { canProbeBorderRows, hasTableMerge } from './table-border-probe.ts';
import { effectiveBorderSide } from './table-border-cascade.ts';
import { withoutFloatingTableZones } from './table-float-overlap.ts';
import type { SemanticTableRow, SemanticTableStructure } from './semantic-table.ts';
import type { TableFlowDeps } from './semantic-table-layout.ts';

interface TerminalBorderBudget {
  readonly row: SemanticTableRow;
  readonly widthPt: number;
}

// Resolved structures are immutable. Keep refused shapes cached as well as supported ones.
const budgets = new WeakMap<SemanticTableStructure, TerminalBorderBudget | null>();

export function tableTerminalDeps(
  structure: SemanticTableStructure,
  flowDeps: TableFlowDeps,
  floatingDeps: TableFlowDeps | undefined,
  outOfFlow: boolean
): TableFlowDeps {
  const deps = floatingDeps ?? (outOfFlow ? withoutFloatingTableZones(flowDeps) : flowDeps);
  if (deps.tableNestingOffset || deps.textBoxStory) return deps;
  let budget = budgets.get(structure);
  if (budget === undefined) {
    budget = readTerminalBudget(structure);
    budgets.set(structure, budget);
  }
  return budget ? { ...deps, exactTerminalBorder: budget } : deps;
}

function readTerminalBudget(structure: SemanticTableStructure): TerminalBorderBudget | null {
  const row = structure.rows.at(-1);
  if (
    !row ||
    row.height.rule !== 'exact' ||
    !Number.isFinite(row.height.valuePt) ||
    row.height.valuePt <= 0 ||
    structure.cellSpacingPt !== 0 ||
    structure.float
  )
    return null;

  const first = structure.rows[0]?.cells[0];
  const second = structure.rows[1]?.cells[0];
  // This merge ends before the ordinary body. It cannot own the terminal row's border.
  const prefixMerge =
    structure.rows.length >= 3 &&
    first?.gridColumn === 0 &&
    first.gridSpan === 1 &&
    first.headsVerticalMerge === true &&
    !first.vMergeContinue &&
    second?.gridColumn === 0 &&
    second.gridSpan === 1 &&
    second.vMergeContinue &&
    !second.headsVerticalMerge &&
    structure.rows.every((r) =>
      r.cells.every(
        (c) => (!c.headsVerticalMerge || c === first) && (!c.vMergeContinue || c === second)
      )
    );
  if (
    (hasTableMerge(structure) && !prefixMerge) ||
    structure.rows.some(
      (r, i) =>
        r.isHeader ||
        r.cells.some(
          (c) =>
            (c.headsVerticalMerge && !(prefixMerge && c === first)) ||
            ((i !== 0 || r === row) && c.gridSpan !== 1)
        )
    )
  )
    return null;
  // Only the eligibility view clears this proved prefix continuation. Source rows stay intact.
  const probeRows = prefixMerge
    ? structure.rows.map((r) =>
        r.cells.includes(second!)
          ? {
              ...r,
              cells: r.cells.map((c) => (c === second ? { ...c, vMergeContinue: false } : c)),
            }
          : r
      )
    : structure.rows;
  if (!canProbeBorderRows(probeRows, structure.columnWidthsPt.length)) return null;

  // Unequal widths, compound rules, and fractional-twip widths keep their existing geometry.
  let width: number | undefined;
  for (const [source, side] of [
    [structure.rows[0]!, 'top'],
    [row, 'bottom'],
  ] as const) {
    for (const cell of source.cells) {
      const edge = effectiveBorderSide(
        cell.borders[side],
        (cell.rowTableBorders ?? structure.tableBorders)[side]
      );
      if (
        edge.state !== 'edge' ||
        edge.style !== 'single' ||
        !Number.isFinite(edge.widthPt) ||
        edge.widthPt <= 0 ||
        edge.widthPt > 12
      )
        return null;
      const points = Math.floor(edge.widthPt * 20 + 1e-8) / 20;
      if (points !== edge.widthPt || (width !== undefined && points !== width)) return null;
      width = points;
    }
  }
  return width && width > 0 ? { row, widthPt: width } : null;
}

export function exactTerminalBorderPt(row: SemanticTableRow, deps: TableFlowDeps): number {
  return deps.exactTerminalBorder?.row === row ? deps.exactTerminalBorder.widthPt : 0;
}

export function fittingBorderPt(exactBottom: number, widthPt: number, maxBottom: number): number {
  return exactBottom + widthPt <= maxBottom + 0.001 ? widthPt : 0;
}
