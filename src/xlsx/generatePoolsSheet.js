import { NUMBER_FIGHTS_POOLS, POOLS_BRACKETS_COLS, POOLS_TEAM_ROW_START, POOL_BRACKET_HEADER_VALUES, POOL_BRACKET_SUB_HEADER_VALUES, getFont } from './sheetConstants.js'

// adding all teams to to the pools
const addTeamsToPoolsSheet = (poolsSheet, teams) => {
  for (let i = 0; i < NUMBER_FIGHTS_POOLS; i++) {
    const rowIndex = POOLS_TEAM_ROW_START + i
    const myIndex = rowIndex
    const oppIndex = (rowIndex % 2) === 0 ? rowIndex - 1 : rowIndex + 1
    const addedRow = poolsSheet.insertRow(rowIndex, {
      A: Math.floor(i / 2) + 1,
      B: '',
      C: '',
      I: { formula: `=IF(K${myIndex}>K${oppIndex},1,0)` },
      J: { formula: `=IF(I${myIndex}=0,1,0)` },
      K: { formula: `=IF(D${myIndex}>D${oppIndex},1,0)+IF(E${myIndex}>E${oppIndex},1,0)+IF(F${myIndex}>F${oppIndex},1,0)+IF(G${myIndex}>G${oppIndex},1,0)+IF(H${myIndex}>H${oppIndex},1,0)` },
      L: { formula: `=N${myIndex}-K${myIndex}-M${myIndex}` },
      M: { formula: `=K${oppIndex}` },
      N: { formula: `=IF(ISBLANK(D${myIndex}),0,1)+IF(ISBLANK(E${myIndex}),0,1)+IF(ISBLANK(F${myIndex}),0,1)+IF(ISBLANK(G${myIndex}),0,1)+IF(ISBLANK(H${myIndex}),0,1)` },
      O: { formula: `=SUM(D${myIndex}:H${myIndex})` },
      P: { formula: `=SUM(D${oppIndex}:H${oppIndex})` },
      Q: { formula: `=O${myIndex}-P${myIndex}` },
      R: '',
      S: ''
    })

    // Set font styles
    addedRow.font = getFont(11, false)
    addedRow.getCell('A').font = getFont(11, true)

    // Merge cells from the second team
    if (i % 2 === 1) {
      const startRow = rowIndex - 1
      const endRow = rowIndex
      poolsSheet.mergeCells(`A${startRow}:A${endRow}`)

      // Set background color for both rows
      const backgroundColor = i % 4 === 1 ? { argb: 'F2F3F4' } : { argb: 'FFFFFFFF' }
      addedRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: backgroundColor
      }

      // Set background color for the row above
      const aboveRow = poolsSheet.getRow(startRow)
      aboveRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: backgroundColor
      }

      // Set borders for both rows
      addedRow.eachCell({ includeEmpty: true }, cell => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
      })

      aboveRow.eachCell({ includeEmpty: true }, cell => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
      })
    }
  }
}

const generatePoolsSheet = (workbook, { teams }) => {
  const headerFont = getFont(11, true)
  const wrapAlignmentTitle = { vertical: 'middle', horizontal: 'center', wrapText: true }
  const wrapAlignmentValue = { vertical: 'middle', horizontal: 'center', wrapText: false }
  const poolsSheet = workbook.addWorksheet('pools')
  poolsSheet.columns = POOLS_BRACKETS_COLS

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

  Object.entries(POOL_BRACKET_HEADER_VALUES).forEach(([col, value]) => {
    const cell = poolsSheet.getCell(`${col}1`)
    cell.value = value
    cell.font = headerFont
    cell.alignment = wrapAlignmentTitle
  })

  Object.entries(POOL_BRACKET_SUB_HEADER_VALUES).forEach(([col, value]) => {
    const cell = poolsSheet.getCell(`${col}2`)
    cell.value = value
    cell.font = headerFont
    cell.alignment = wrapAlignmentValue
  })
}

export default generatePoolsSheet
