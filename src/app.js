import tempfile from 'tempfile'
import express from 'express'
import generateXlsx from './xlsx/index.js'
const port = process.env.PORT || 8080
const app = express()

app.get('/', (req, res) => {
  const workbook = generateXlsx()
  res.statusCode = 200
  const tempFilePath = tempfile({ extension: '.xlsx' })
  workbook.xlsx.writeFile(tempFilePath).then(function () {
    res.download(tempFilePath, 'test.xlsx')
  })
})

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})
