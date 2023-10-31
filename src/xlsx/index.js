import ExcelJS from 'exceljs'
import generateResultsSheet from './generateResultsSheet.js'

const generateXlsx = () => {
  const workbook = new ExcelJS.Workbook()
  generateResultsSheet(workbook)
  return workbook
}

export default generateXlsx
