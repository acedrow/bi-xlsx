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

/**
expects JSON body in the shape of:
{
  eventName: ${name_string_variable},
  date: ${date_string_variable},
  location: ${location_string_variable},
  teams: [
    "team1", "team2"
      ]
}
*/

app.post('/generateSpreadsheet', (req, res) => {
  console.log('post request', req.body)
  const workbook = generateXlsx()
  res.statusCode = 200
  const tempFilePath = tempfile({ extension: '.xlsx' })
  workbook.xlsx.writeFile(tempFilePath).then(function () {
    res.download(tempFilePath, 'test.xlsx')
  })
})

app.post('/echoBody', (req, res) => {
  console.log('post request', req.body)
  res.json({ requestBody: req.body })
})

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})
