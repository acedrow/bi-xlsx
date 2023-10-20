import tempfile from 'tempfile'
import ExcelJS from 'exceljs'
import express from 'express'
const port = process.env.PORT || 8080
const app = express()

const initWorkbook = () => {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('My Sheet')

  worksheet.columns = [
    { key: 'A', width: 10 },
    { key: 'B', width: 32 },
    { key: 'C', width: 10 }
  ]

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

app.get('/', (req, res) => {
  const workbook = initWorkbook()
  res.statusCode = 200
  const tempFilePath = tempfile({ extension: '.xlsx' })
  workbook.xlsx.writeFile(tempFilePath).then(function () {
    res.download(tempFilePath, 'test.xlsx')
  })
})

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})
