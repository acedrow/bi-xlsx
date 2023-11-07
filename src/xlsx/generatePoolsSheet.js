// TODO: Linden: review this file
import { RESULTS_COLS, getFont } from './sheetConstants.js'

// adding all teams to to the pools
const addTeamsToPoolsSheet = (poolsSheet, teams) => {
  teams.forEach((team, index) => {
    const rowIndex = 3 + index
    // commented out the team + id (maybe we will use it maybe not)
    const addedRow = poolsSheet.insertRow(rowIndex /*, {
      // B: team.id,
      // C: team.name
    } */)
    addedRow.font = getFont(11, false)
    // Merge cells from the second team
    if (index % 2 === 1) {
      const startRow = rowIndex - 1
      const endRow = rowIndex
      poolsSheet.mergeCells(`A${startRow}:A${endRow}`)
    }
  })

  // Check if uneven, yes => add and merge one more row
  if (teams.length % 2 === 1) {
    const lastRowIndex = 3 + teams.length
    poolsSheet.insertRow(lastRowIndex, {
      B: null,
      C: null
    })
    // Merge this empty row with the second to last row
    const startRow = lastRowIndex - 1
    const endRow = lastRowIndex
    poolsSheet.mergeCells(`A${startRow}:A${endRow}`)
  }
}

const generatePoolsSheet = (workbook, { teams }) => {
  const headerFont = getFont(11, true)
  const wrapAlignmentTitle = { vertical: 'middle', horizontal: 'center', wrapText: true }
  const wrapAlignmentValue = { vertical: 'middle', horizontal: 'center', wrapText: false }
  const poolsSheet = workbook.addWorksheet('Pools')
  poolsSheet.columns = RESULTS_COLS

  addTeamsToPoolsSheet(poolsSheet, teams)

  // Merge cells for the header
  poolsSheet.mergeCells('A1:A2')
  poolsSheet.mergeCells('B1:B2')
  poolsSheet.mergeCells('C1:C2')
  poolsSheet.mergeCells('D1:H1')
  poolsSheet.mergeCells('I1:J1')
  poolsSheet.mergeCells('K1:N1')
  poolsSheet.mergeCells('O1:Q1')
  poolsSheet.mergeCells('R1:S1')

  const headerRowValues = {
    A: 'Fight',
    B: 'Team/Fighter ID',
    C: 'Team',
    D: 'Rounds Score',
    I: 'Fight',
    K: 'Rounds',
    O: 'Active/Grounded',
    R: 'Penalties'
  }

  Object.entries(headerRowValues).forEach(([col, value]) => {
    const cell = poolsSheet.getCell(`${col}1`)
    cell.value = value
    cell.font = headerFont
    cell.alignment = wrapAlignmentTitle
  })

  // Set the header values for the second row
  const headerRowValues2 = {
    D: 'R1',
    E: 'R2',
    F: 'R3',
    G: 'R4',
    H: 'R5',
    I: 'Win',
    J: 'Loss',
    K: 'Win',
    L: 'Draw',
    M: 'Loss',
    N: 'Ratio',
    O: 'Active',
    P: 'Grounded',
    Q: 'Ratio',
    R: 'Yellow Card',
    S: 'Red Card'
  }

  Object.entries(headerRowValues2).forEach(([col, value]) => {
    const cell = poolsSheet.getCell(`${col}2`)
    cell.value = value
    cell.font = headerFont
    cell.alignment = wrapAlignmentValue
  })

  // format
  poolsSheet.columns.forEach(column => {
    let maxLength = 0

    column.eachCell({ includeEmpty: true }, cell => {
      let cellLength = (cell.value && cell.value.toString().length) || 0
      cellLength += 2
      if (cellLength > maxLength) {
        maxLength = cellLength
      }
    })
    column.width = maxLength
  })
}

export default generatePoolsSheet
