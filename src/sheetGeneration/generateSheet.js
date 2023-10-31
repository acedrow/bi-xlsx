import ExcelJS from 'exceljs'
import { WORKSHEET_COLS } from './sheetConstants.js'

const initWorkbook = () => {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('My Sheet')

  worksheet.columns = WORKSHEET_COLS

  const row0 = worksheet.addRow({ A: 'A!', C: 'C!' })
  row0.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'ff4ccf70' }
  }
  row0.font = {
    name: 'Arial Black',
    color: { argb: '00000000' },
    family: 2,
    size: 14,
    italic: false
  }
  const row1 = worksheet.addRow({ A: 'B should be dropdown', C: 'C!' })

  worksheet.getCell('B2').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['One,Two,Three,Forr']
  }

  return workbook
}

const generateXlsx = () => {
  return initWorkbook()
}

export default generateXlsx
