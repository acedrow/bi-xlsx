import { RESULTS_COLS, TITLE_ROW_PROPS } from './sheetConstants.js'

const generateTitleRows = (worksheet) => {
  worksheet.addRow({ A: 'BUHURT INTERNATIONAL' })

  worksheet.getCell('A1').fill = TITLE_ROW_PROPS.fill
  worksheet.getCell('A1').alignment = TITLE_ROW_PROPS.alignment
  worksheet.getCell('A1').font = TITLE_ROW_PROPS.font

  worksheet.addRow({ A: 'Event Scoring Sheet' })

  worksheet.getCell('A2').fill = TITLE_ROW_PROPS.fill
  worksheet.getCell('A2').alignment = TITLE_ROW_PROPS.alignment
  worksheet.getCell('A2').font = TITLE_ROW_PROPS.font
  worksheet.getCell('A2').font.size = 12

  worksheet.mergeCells('A1:Z1')
  worksheet.mergeCells('A2:Z2')
}

const generateResultsSheet = (workbook) => {
  const resultsSheet = workbook.addWorksheet('results')
  resultsSheet.columns = RESULTS_COLS
  generateTitleRows(resultsSheet)

//   resultsSheet.getCell('B2').dataValidation = {
//     type: 'list',
//     allowBlank: true,
//     formulae: ['One,Two,Three,Forr']
//   }
}

export default generateResultsSheet
