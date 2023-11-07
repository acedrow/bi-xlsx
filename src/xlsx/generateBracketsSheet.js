// TODO: Linden: review this file
import { RESULTS_COLS, getFont } from './sheetConstants.js'

// adding all teams to to the bracket
const addTeamsToBracketSheet = (bracketsSheet, teams) => {
  teams.forEach((team, index) => {
    const rowIndex = 3 + index
    // commented out the team + id (maybe we will use it maybe not)
    const addedRow = bracketsSheet.insertRow(rowIndex /*, {
      // B: team.id,
      // C: team.name
    } */)
    addedRow.font = getFont(11, false)
    // Merge cells from the second team
    if (index % 2 === 1) {
      const startRow = rowIndex - 1
      const endRow = rowIndex
      bracketsSheet.mergeCells(`A${startRow}:A${endRow}`)
    }
  })

  // Check if uneven, yes => add and merge one more row
  if (teams.length % 2 === 1) {
    const lastRowIndex = 3 + teams.length
    bracketsSheet.insertRow(lastRowIndex, {
      B: null,
      C: null
    })
    // Merge this empty row with the second to last row
    const startRow = lastRowIndex - 1
    const endRow = lastRowIndex
    bracketsSheet.mergeCells(`A${startRow}:A${endRow}`)
  }
}

const generateBracketsSheet = (workbook, { teams }) => {
  const headerFont = getFont(11, true)
  const wrapAlignmentTitle = { vertical: 'middle', horizontal: 'center', wrapText: true }
  const wrapAlignmentValue = { vertical: 'middle', horizontal: 'center', wrapText: false }
  const bracketsSheet = workbook.addWorksheet('Brackets')

  bracketsSheet.columns = RESULTS_COLS

  addTeamsToBracketSheet(bracketsSheet, teams)

  // Merge cells for the header
  bracketsSheet.mergeCells('A1:A2')
  bracketsSheet.mergeCells('B1:B2')
  bracketsSheet.mergeCells('C1:C2')
  bracketsSheet.mergeCells('D1:H1')
  bracketsSheet.mergeCells('I1:J1')
  bracketsSheet.mergeCells('K1:N1')
  bracketsSheet.mergeCells('O1:Q1')
  bracketsSheet.mergeCells('R1:S1')

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
    const cell = bracketsSheet.getCell(`${col}1`)
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
    const cell = bracketsSheet.getCell(`${col}2`)
    cell.value = value
    cell.font = headerFont
    cell.alignment = wrapAlignmentValue
  })
  bracketsSheet.columns.forEach(column => {
    let maxLength = 0

    column.eachCell({ includeEmpty: true }, cell => {
      const cellValue = cell.value
      let cellLength = (cellValue && cellValue.toString().length) || 0
      cellLength += 2
      if (cellLength > maxLength) {
        maxLength = cellLength
      }
    })
    column.width = maxLength
  })
}

export default generateBracketsSheet
