import ExcelJS from 'exceljs'
import generateResultsSheet from './generateResultsSheet.js'
import populatePoolsBracketsSheet from './populatePoolsBracketsSheet.js'

const generateXlsx = ({ eventName, date, location, teams, fighters, tournamentType }) => {
  const workbook = new ExcelJS.Workbook()
  const competitors = teams || fighters

  const resultsSheet = workbook.addWorksheet('results')
  generateResultsSheet(resultsSheet, { eventName, date, location, teams: competitors, tournamentType })
  const poolsSheet = workbook.addWorksheet('pools')
  populatePoolsBracketsSheet(poolsSheet, { eventName, date, location, teams: competitors, tournamentType })
  const bracketsSheet = workbook.addWorksheet('brackets')
  populatePoolsBracketsSheet(bracketsSheet, { eventName, date, location, teams: competitors, tournamentType })

  return workbook
}

export default generateXlsx
