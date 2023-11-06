// TODO: Linden: review this file
import { RESULTS_COLS, getFont } from './sheetConstants.js'

// sheet 3 Brackets
const generateBracketsSheet = (workbook) => {
  const headerFont = getFont(11, true)
  const wrapAlignmentTitle = { vertical: 'middle', horizontal: 'center', wrapText: true }
  const wrapAlignmentValue = { vertical: 'middle', horizontal: 'center', wrapText: false }
  const bracketsSheet = workbook.addWorksheet('Brackets')

  bracketsSheet.columns = RESULTS_COLS

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
    I: 'FW',
    J: 'FL',
    K: 'Win',
    L: 'Draw',
    M: 'Loss',
    N: 'Round Neutral',
    O: 'Standing',
    P: 'Ground',
    Q: 'Stand/Ground Diff',
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

    // Loop through all cells in a column
    column.eachCell({ includeEmpty: true }, cell => {
      // Calculate the maximum length of cell value
      const cellValue = cell.value
      let cellLength = (cellValue && cellValue.toString().length) || 0

      // Add extra space for aesthetics
      cellLength += 2
      if (cellLength > maxLength) {
        maxLength = cellLength
      }
    })
    column.width = maxLength
  })
}

export default generateBracketsSheet
