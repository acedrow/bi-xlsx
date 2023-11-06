import { RESULTS_COLS, TITLE_ROW_PROPS, getFont } from './sheetConstants.js'
// sheet 1 Results
const generateTitleRows = (worksheet) => {
// ROW 1
  worksheet.addRow({ A: 'BUHURT INTERNATIONAL' })

  worksheet.getCell('A1').fill = TITLE_ROW_PROPS.fill
  worksheet.getCell('A1').alignment = TITLE_ROW_PROPS.alignment
  worksheet.getCell('A1').font = getFont(16, true)
  // ROW 2
  worksheet.addRow({ A: 'Event Scoring Sheet' })

  worksheet.getCell('A2').fill = TITLE_ROW_PROPS.fill
  worksheet.getCell('A2').alignment = TITLE_ROW_PROPS.alignment
  worksheet.getCell('A2').font = getFont(12, true)

  worksheet.mergeCells('A1:Z1')
  worksheet.mergeCells('A2:Z2')
}

const generateHeaderRows = (worksheet, eventName, date, location) => {
  const row3 = worksheet.addRow({
    A: 'Event Name:',
    B: eventName,
    T: 'Event Date:',
    W: date,
    Y: 'Event Tier:',
    Z: 'Classic'
  })

  row3.font = getFont(11, true)
  worksheet.getCell('Z3').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"Classic,Regional,Conference"']
  }
  worksheet.mergeCells('B3:S3')
  worksheet.mergeCells('T3:V3')
  worksheet.mergeCells('W3:X3')

  const row4 = worksheet.addRow({
    A: 'Event Location:',
    B: location,
    T: 'Finals Type:',
    W: 'Round Robin',
    Y: 'Tier Mult.',
    Z: { formula: '=IF(Z3="Regional",1.5,IF(Z3="Conference",2,1))' }
  })

  row4.font = getFont(11, true)
  worksheet.getCell('W4').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"Round Robin,Bracket"']
  }
  worksheet.mergeCells('B4:S4')
  worksheet.mergeCells('T4:V4')
  worksheet.mergeCells('W4:X4')
}

const generateTeamHeaders = (worksheet) => {
  const headerFont = getFont(11, true)
  const wrapAlignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
  // Add headers for row 5 with common font and alignment
  const row5 = worksheet.addRow({
    A: 'Team/Fighter ID',
    B: 'Team',
    C: 'T',
    D: 'Fights',
    I: 'Rounds',
    N: 'Score',
    W: 'Points',
    X: 'Placement',
    Y: 'Rank adj. Points',
    Z: 'Final Awarded Points'
  });

  ['A', 'B', 'C', 'D', 'I', 'N', 'W', 'X', 'Y', 'Z'].forEach((col) => {
    const cell = row5.getCell(col)
    cell.font = headerFont
    cell.alignment = wrapAlignment
  })

  // Add headers for row 6 and styles
  const headersRow6 = ['Fw', 'FL', 'Ff', 'Fw/R', 'F/T', 'Rw', 'Rd', 'RI', 'Rf', 'Rw/R', 'A', 'A/R', 'G', 'G/R', 'A/G D', 'A/G R', 'YK', 'RK', 'Total']
  const row6 = worksheet.addRow()
  headersRow6.forEach((title, index) => {
    // Get the column letter from the index offset by 4
    const colLetter = String.fromCharCode('D'.charCodeAt(0) + index)
    const cell = row6.getCell(colLetter)
    cell.value = title
    cell.font = headerFont
    cell.alignment = wrapAlignment
  })

  worksheet.mergeCells('A5:A6')
  worksheet.mergeCells('B5:B6')
  worksheet.mergeCells('C5:C6')
  worksheet.mergeCells('D5:H5')
  worksheet.mergeCells('I5:M5')
  worksheet.mergeCells('N5:S5')
  worksheet.mergeCells('W5:W6')
  worksheet.mergeCells('X5:X6')
  worksheet.mergeCells('Y5:Y6')
  worksheet.mergeCells('Z5:Z6')
}
// generates the team names + id
const generateTeamDataRows = (worksheet, teams) => {
  console.log('Adding teams to results  (sheet1):', teams)
  teams.forEach(team => {
    const addedRow = worksheet.addRow({
      A: team.id,
      B: team.name
    })
    addedRow.font = getFont(11, false)
  }
  )
}

// sheet 2 Pools

// adding all teams to to the pools
const addTeamsToPoolsSheet = (poolsSheet, teams) => {
  console.log('Adding teams to pools  (sheet2):', teams)
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

const generatePoolsSheet = (workbook, teams) => {
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

const generateResultsSheet = (workbook, { eventName, date, location, teams }) => {
  const resultsSheet = workbook.addWorksheet('Results')
  resultsSheet.columns = RESULTS_COLS
  generateTitleRows(resultsSheet)
  generateHeaderRows(resultsSheet, eventName, date, location, teams)
  generateTeamHeaders(resultsSheet)
  generateTeamDataRows(resultsSheet, teams)
  generatePoolsSheet(workbook, teams)
  generateBracketsSheet(workbook)
}

export default generateResultsSheet
