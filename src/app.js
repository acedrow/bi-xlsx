// TODO: set up pre-commit hook
import tempfile from 'tempfile'
import express from 'express'
import cors from 'cors'
import generateXlsx from './xlsx/index.js'
import { TEST_DATA, TOURNAMENT_TYPE_FIGHTERS, TOURNAMENT_TYPE_TEAMS } from './xlsx/sheetConstants.js'
import { validate, ValidationError, Joi } from 'express-validation'

const port = process.env.PORT || 8080

const app = express()

app.use(cors())
app.use(express.json()) // <==== parse request body as JSON

const corsOptions = {
  origin: 'https://www.buhurtinternational.com',
  optionsSuccessStatus: 200 // some legacy browsers (IE11, various SmartTVs) choke on 204
}

// validation schema for /generateSpreadsheet
const generationValidation = {
  body: Joi.object({
    eventName: Joi.string()
      .required(),
    date: Joi.string()
      .required(),
    location: Joi.string()
      .required(),
    tournamentType: Joi.string()
      .valid(TOURNAMENT_TYPE_FIGHTERS, TOURNAMENT_TYPE_TEAMS).required(),
    teams: Joi.array().items(
      Joi.object({ name: Joi.string().required() }).unknown(true))
  })
}

app.post('/generateSpreadsheet', [cors(corsOptions), validate(generationValidation, {}, {})], (req, res) => {
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

// validation error handling
app.use(function (err, req, res, next) {
  if (err instanceof ValidationError) {
    return res.status(err.statusCode).json(err)
  }

  return res.status(500).json(err)
})

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})
