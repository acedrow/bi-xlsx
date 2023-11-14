// TODO: set up pre-commit hook
import tempfile from 'tempfile'
import express from 'express'
import cors from 'cors'
import generateXlsx from './xlsx/index.js'

const TEST_DATA = {
  eventName: 'test event',
  date: '1/2/3',
  location: 'St. Paul',
  tournamentType: 'teams',
  teams: [
    { name: 'wyverns', id: 'tha best' },
    { name: 'dfc', id: 'tha' },
    { name: 'knyaz', id: 'big boiz' }
  ],
  duellists: [
    { name: 'Derek', id: 'doc' },
    { name: 'Keegan', id: 'too big' },
    { name: 'Linden', id: 'lich boi' },
    { name: 'Vish', id: '420' },
    { name: 'Spence', id: 'coach' }
  ]
}

const port = process.env.PORT || 8080

const app = express()

app.use(cors())
app.use(express.json()) // <==== parse request body as JSON

const corsOptions = {
  origin: 'https://www.buhurtinternational.com',
  optionsSuccessStatus: 200 // some legacy browsers (IE11, various SmartTVs) choke on 204
}

app.post('/generateSpreadsheet', cors(corsOptions), (req, res) => {
  console.log('post request', req.body)
  const workbook = generateXlsx(req.body)
  res.statusCode = 200
  res.set('Access-Control-Allow-Origin', 'https://www.buhurtinternational.com')
  const tempFilePath = tempfile({ extension: '.xlsx' })
  workbook.xlsx.writeFile(tempFilePath).then(function () {
    res.download(tempFilePath, 'test.xlsx')
  })
})

// TODO: for testing, delete for prod
app.get('/', (req, res) => {
  const workbook = generateXlsx(TEST_DATA)
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
