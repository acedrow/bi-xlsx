import { EVENT_TIERS, FINALS_TYPE, RESULTS_COLS, TITLE_ROW_PROPS, getFont } from './sheetConstants.js'

const generateTitleRows = (worksheet) => {
  worksheet.addRow({ A: 'BUHURT INTERNATIONAL' })

  worksheet.getCell('A1').fill = TITLE_ROW_PROPS.fill
  worksheet.getCell('A1').alignment = TITLE_ROW_PROPS.alignment
  worksheet.getCell('A1').font = getFont(16, true)

  worksheet.addRow({ A: 'Event Scoring Sheet' })

  worksheet.getCell('A2').fill = TITLE_ROW_PROPS.fill
  worksheet.getCell('A2').alignment = TITLE_ROW_PROPS.alignment
  worksheet.getCell('A2').font = getFont(12, true)

  worksheet.mergeCells('A1:Z1')
  worksheet.mergeCells('A2:Z2')
}

// TODO: need to pull event name, location, and date in from request object
const generateHeaderRows = (worksheet) => {
  const row3 = worksheet.addRow(
    {
      A: 'Event Name:',
      B: 'TEST EVENT NAME',
      T: 'Event Date:',
      W: 'TEST DATE',
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

  const row4 = worksheet.addRow(
    {
      A: 'Event Location:',
      B: 'TEST EVENT LOC',
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

const generateResultsSheet = (workbook) => {
  const resultsSheet = workbook.addWorksheet('results')
  resultsSheet.columns = RESULTS_COLS
  generateTitleRows(resultsSheet)
  generateHeaderRows(resultsSheet)
}

export default generateResultsSheet
