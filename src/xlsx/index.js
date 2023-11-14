import ExcelJS from 'exceljs'
import generateResultsSheet from './generateResultsSheet.js'
import generatePoolsSheet from './generatePoolsSheet.js'
import generateBracketsSheet from './generateBracketsSheet.js'

const generateXlsx = ({ eventName, date, location, teams, fighters, tournamentType }) => {
  const workbook = new ExcelJS.Workbook()
  const competitors = teams || fighters
  generateResultsSheet(workbook, { eventName, date, location, teams: competitors, tournamentType })
  generatePoolsSheet(workbook, { eventName, date, location, teams: competitors, tournamentType })
  generateBracketsSheet(workbook, { eventName, date, location, teams: competitors, tournamentType })
  return workbook
}

export default generateXlsx
