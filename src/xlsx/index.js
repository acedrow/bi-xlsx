import ExcelJS from 'exceljs'
import generateResultsSheet from './generateResultsSheet.js'
import generatePoolsSheet from './generatePoolsSheet.js'
import generateBracketsSheet from './generateBracketsSheet.js'

const generateXlsx = ({ eventName, date, location, teams }) => {
  const workbook = new ExcelJS.Workbook()
  generateResultsSheet(workbook, { eventName, date, location, teams })
  generatePoolsSheet(workbook, { eventName, date, location, teams })
  generateBracketsSheet(workbook, { eventName, date, location, teams })
  return workbook
}

export default generateXlsx
