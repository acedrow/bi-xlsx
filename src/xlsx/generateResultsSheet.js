import { RESULTS_COLS, TITLE_ROW_PROPS, getFont } from './sheetConstants.js'

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
  const row5 = worksheet.addRow(
    {
      A: 'Team/Fighter ID',
      B: 'Team',
      C: 'T',
      D: 'Fights',
      I: 'Rounds',
      N: 'Score  '

    })
  row5.font = getFont(11, true)
  worksheet.mergeCells('A5:A6')
  worksheet.mergeCells('B5:B6')
  worksheet.mergeCells('C5:C6')
  worksheet.mergeCells('D5:H5')
  worksheet.mergeCells('I5:M5')
  worksheet.mergeCells('N5:S5')
  worksheet.mergeCells('T5:V5')
}

const generateTeamDataRows = (worksheet, teams) => {
  console.log('generateTeamDataRows teams', teams)
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
