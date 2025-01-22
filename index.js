const XLSX = require('xlsx')
const path = require('path')

const inputFile = 'eva_lang_sample.xlsx'

function sanitizeString (input) {
  return input.replace(/[^a-zA-Z0-9]/g, '')
}


const inputFilePath = path.join(__dirname, inputFile)
const outputFilePath = path.join(__dirname, 'output.xlsx')

const workbook = XLSX.readFile(inputFilePath)

const sheetName = workbook.SheetNames[0]
const worksheet = workbook.Sheets[sheetName]


const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 })
const occurrenceMap = {}

for (let i = 1; i < data.length; i++) {
  const row = data[i]
  if (row[0]) {
    // Sanitize column 1
    let sanitized = sanitizeString(row[0].toString())

    if (occurrenceMap[sanitized]) {
      occurrenceMap[sanitized] += 1
      sanitized = `${sanitized}(${occurrenceMap[sanitized]})`
    } else {
      occurrenceMap[sanitized] = 1
    }

    // Assign sanitized and unique value to column 5
    row[4] = sanitized
  }
}

const newWorksheet = XLSX.utils.aoa_to_sheet(data)
workbook.Sheets[sheetName] = newWorksheet
XLSX.writeFile(workbook, outputFilePath)

console.log(`Processed Excel file saved to: ${outputFilePath}`)
