import tempfile from 'tempfile'
import express from 'express'
import cors from 'cors'
import generateXlsx from './xlsx/index.js'
// import fs from 'fs'
// import crypto from 'crypto'
// import path from 'path'

const port = process.env.PORT || 8080

const app = express()

app.use(cors())
app.use(express.json()) // <==== parse request body as JSON

const corsOptions = {
  // origin: 'https://www.buhurtinternational.com/',.
  origin: 'http://localhost:8080',
  optionsSuccessStatus: 200 // some legacy browsers (IE11, various SmartTVs) choke on 204
}

app.post('/generateSpreadsheet', cors(corsOptions), (req, res) => {
  console.log('post request', req.body)
  const workbook = generateXlsx(req.body.eventName, req.body)
  res.statusCode = 200
  const tempFilePath = tempfile({ extension: '.xlsx' })
  workbook.xlsx.writeFile(tempFilePath).then(function () {
    res.download(tempFilePath, 'test.xlsx')
  })
})

app.get('/', (req, res) => {
  const workbook = generateXlsx(req.body.eventName, req.body)
  res.statusCode = 200
  const tempFilePath = tempfile({ extension: '.xlsx' })
  workbook.xlsx.writeFile(tempFilePath).then(function () {
    res.download(tempFilePath, 'test.xlsx')
  })
})

app.post('/echoBody', cors(corsOptions), (req, res) => {
  console.log('post request', req.body)
  res.json({ requestBody: req.body })
})

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})
