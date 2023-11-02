import tempfile from 'tempfile'
import express from 'express'
import generateXlsx from './xlsx/index.js'
const port = process.env.PORT || 8080
const app = express()
app.use(express.json()) // <==== parse request body as JSON

app.get('/', (req, res) => {
  const workbook = generateXlsx()
  res.statusCode = 200
  const tempFilePath = tempfile({ extension: '.xlsx' })
  workbook.xlsx.writeFile(tempFilePath).then(function () {
    res.download(tempFilePath, 'test.xlsx')
  })
})

app.post('/generateSpreadsheet', (req, res) => {
  console.log('post request', req.body)
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
