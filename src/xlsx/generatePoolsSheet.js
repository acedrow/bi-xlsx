// TODO: Linden: review this file
import { NUMBER_FIGHTS_POOLS, POOLS_TEAM_ROW_START, RESULTS_COLS, getFont } from './sheetConstants.js'

// adding all teams to to the pools
const addTeamsToPoolsSheet = (poolsSheet, teams) => {
  for (let i = 0; i < NUMBER_FIGHTS_POOLS; i++) {
    const rowIndex = POOLS_TEAM_ROW_START + i
    const myIndex = rowIndex
    const oppIndex = (rowIndex % 2) === 0 ? rowIndex - 1 : rowIndex + 1
    const addedRow = poolsSheet.insertRow(rowIndex, {
      A: (i / 2),
      B: '',
      C: '',
      I: { formula: `=IF(K${myIndex}>K${oppIndex},1,0)` },
      J: { formula: `=IF(I${myIndex}=0,1,0)` },
      K: { formula: `=IF(D${myIndex}>D${oppIndex},1,0)+IF(E${myIndex}>E${oppIndex},1,0)+IF(F${myIndex}>F${oppIndex},1,0)+IF(G${myIndex}>G${oppIndex},1,0)+IF(H${myIndex}>H${oppIndex},1,0)` },
      L: { formula: `=N${myIndex}-K${myIndex}-M${myIndex}` },
      M: { formula: `=K${oppIndex}` },
      N: { formula: `=IF(ISBLANK(D${myIndex}),0,1)+IF(ISBLANK(E${myIndex}),0,1)+IF(ISBLA  NK(F${myIndex}),0,1)+IF(ISBLANK(G${myIndex}),0,1)+IF(ISBLANK(H${myIndex}),0,1)` },
      O: { formula: `=SUM(D${myIndex}:H${myIndex})` },
      P: { formula: `=SUM(D${oppIndex}:H${oppIndex})` },
      Q: { formula: `=O${myIndex}-P${myIndex}` }
    })
    addedRow.font = getFont(11, false)
    addedRow.getCell('A').font = getFont(11, true)

    // Merge cells from the second team
    if (i % 2 === 1) {
      const startRow = rowIndex - 1
      const endRow = rowIndex
      poolsSheet.mergeCells(`A${startRow}:A${endRow}`)
    }
  }
}

const generatePoolsSheet = (workbook, { teams }) => {
  const headerFont = getFont(11, true)
  const wrapAlignmentTitle = { vertical: 'middle', horizontal: 'center', wrapText: true }
  const wrapAlignmentValue = { vertical: 'middle', horizontal: 'center', wrapText: false }
  const poolsSheet = workbook.addWorksheet('pools')
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
