import tempfile from 'tempfile'
import ExcelJS from 'exceljs'
import express from 'express'
const port = 3000
const app = express()

const initWorkbook = () => {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('My Sheet')

  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
    { header: 'D.O.B.', key: 'DOB', width: 10 }
  ]
  worksheet.addRow({ id: 1, name: 'John Doe', dob: new Date(1970, 1, 1) })
  worksheet.addRow({ id: 2, name: 'Jane Doe', dob: new Date(1965, 1, 7) })
  return workbook
}

app.get('/', (req, res) => {
  const workbook = initWorkbook()
  res.statusCode = 200
  const tempFilePath = tempfile({ extension: '.xlsx' })
  workbook.xlsx.writeFile(tempFilePath).then(function () {
    res.download(tempFilePath)
  })
})

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})
