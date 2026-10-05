import type { Model } from '@ironcalc/workbook'
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from '../../utils/sheetWorkbook'

const SHEET_SELECTION_GUARD_INSTALLED = Symbol('sheetSelectionGuardInstalled')

type GuardedSheetModel = Model & {
  [SHEET_SELECTION_GUARD_INSTALLED]?: true
}

function isValidSheetCoordinate(row: number, column: number): boolean {
  return Number.isInteger(row)
    && row >= 1
    && row <= MAX_SHEET_ROWS
    && Number.isInteger(column)
    && column >= 1
    && column <= MAX_SHEET_COLUMNS
}

export function installSheetSelectionGuard(model: Model): void {
  const guardedModel = model as GuardedSheetModel
  if (guardedModel[SHEET_SELECTION_GUARD_INSTALLED]) return

  const setSelectedCell = model.setSelectedCell.bind(model)
  model.setSelectedCell = (row, column) => {
    if (isValidSheetCoordinate(row, column)) setSelectedCell(row, column)
  }
  guardedModel[SHEET_SELECTION_GUARD_INSTALLED] = true
}
