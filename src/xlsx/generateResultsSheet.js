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
  const wrapAlignment = { vertical: 'middle', horizontal: 'center', wrapText: false }
  // Add headers for row 5 with common font and alignment
  const row5 = worksheet.addRow({
    A: 'ID',
    B: 'Team',
    C: 'T',
    D: 'Fights',
    I: 'Rounds',
    N: 'Score',
    T: 'Cards',
    W: 'Points',
    X: 'Placement',
    Y: 'Rank adj. Points',
    Z: 'Final Awarded Points'
  })

  row5.font = headerFont
  row5.alignment = wrapAlignment

  // Add headers for row 6 and styles
  const row6 = worksheet.addRow({
    // Fights
    D: 'Won',
    E: 'Loss',
    F: 'Total',
    G: 'Win Ratio',
    // H: '', // will not be used but keep it ?
    // Rounds
    I: 'Win',
    J: 'Draw',
    K: 'Loss',
    L: 'Total',
    M: 'win ratio',
    // Active/Grounded
    N: 'Active',
    O: 'per round',
    P: 'Grounded',
    Q: 'per round',
    R: 'A/G difference', // active versus ground difference
    S: 'A/G Ratio', // active ground ratio read "kill/death Ratio"
    // penaties
    T: 'Yellow',
    U: 'Red',
    V: 'Total'
  })

  row6.font = headerFont
  row6.alignment = wrapAlignment

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
  worksheet.mergeCells('T5:V5')

  worksheet.getColumn('H').hidden = true
}

// generates the team names + id
const generateTeamDataRows = (worksheet, teams) => {
  teams.forEach(team => {
    const addedRow = worksheet.addRow({
      A: team.id,
      B: team.name
    })
    addedRow.font = getFont(11, false)
  }
  )
}

const generateResultsSheet = (workbook, { eventName, date, location, teams }) => {
  const resultsSheet = workbook.addWorksheet('Results')
  resultsSheet.columns = RESULTS_COLS
  generateTitleRows(resultsSheet)
  generateHeaderRows(resultsSheet, eventName, date, location, teams)
  generateTeamHeaders(resultsSheet)
  generateTeamDataRows(resultsSheet, teams)
}

export default generateResultsSheet
