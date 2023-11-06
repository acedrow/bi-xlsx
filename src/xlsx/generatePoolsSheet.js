// TODO: Linden: review this file
import { RESULTS_COLS, getFont } from './sheetConstants.js'

// sheet 2 Pools

// adding all teams to to the pools
const addTeamsToPoolsSheet = (poolsSheet, teams) => {
  // Assuming row 1 and 2 are for headers
  const currentRow = 3 // Start from row 3
  teams.forEach((team, index) => {
    // The library might allow specifying the row number
    const addedRow = poolsSheet.insertRow(currentRow + index, {
      A: team.id,
      B: team.name
    })
    addedRow.font = getFont(11, false)
  })
}

const generatePoolsSheet = (workbook, { teams }) => {
  const headerFont = getFont(11, true)
  const wrapAlignmentTitle = { vertical: 'middle', horizontal: 'center', wrapText: true }
  const wrapAlignmentValue = { vertical: 'middle', horizontal: 'center', wrapText: false }
  const poolsSheet = workbook.addWorksheet('Pools')
  poolsSheet.columns = RESULTS_COLS

  addTeamsToPoolsSheet(poolsSheet, teams)

  // Merge cells for the header - This should be done only once, not inside the loop
  poolsSheet.mergeCells('A1:A2')
  poolsSheet.mergeCells('B1:B2')
  poolsSheet.mergeCells('C1:C2')
  poolsSheet.mergeCells('D1:H1')
  poolsSheet.mergeCells('I1:J1')
  poolsSheet.mergeCells('K1:N1')
  poolsSheet.mergeCells('O1:Q1')
  poolsSheet.mergeCells('R1:S1')

  // Set the header values for the first row
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

  // Set the second row of header values
  const headerRowValues2 = {
    D: 'R1',
    E: 'R2',
    F: 'R3',
    G: 'R4',
    H: 'R5',
    I: 'Fw',
    J: 'FL',
    K: 'Rw',
    L: 'Rd',
    M: 'RL',
    N: 'R',
    O: 'Active',
    P: 'Grounded',
    Q: 'A/G dif',
    R: 'YK',
    S: 'RK'
  }

  Object.entries(headerRowValues2).forEach(([col, value]) => {
    const cell = poolsSheet.getCell(`${col}2`)
    cell.value = value
    cell.font = headerFont
    cell.alignment = wrapAlignmentValue
  })
}

export default generatePoolsSheet
