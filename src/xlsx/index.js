import ExcelJS from 'exceljs'
import generateResultsSheet from './generateResultsSheet.js'

const generateXlsx = (eventName, { date, location, teams }) => {
  const workbook = new ExcelJS.Workbook()
  generateResultsSheet(workbook, eventName, { date, location, teams })
  return workbook
}

export default generateXlsx
