#!/usr/bin/env node

const fs = require('fs')
const XLSX = require('xlsx')
const { Command } = require('commander')

const program = new Command()
program
  .option('-o, --out <file>', 'output file')
  .option('--split', 'split by sheet name')
  .option('--multi', 'multi sheet')
  .option('--rawStr', 'parse the input data as plain text')

function main() {
  
  program.parse()
  const options = program.opts()
  const argv = program.args

  if (!argv || argv.length < 1) {
    console.log('basic usage: tool-e2j *.{xlsx,csv}')
  } else {

    for (const inFile of argv) {
      if (!fs.existsSync(inFile)) {
        console.error(`File: ${inFile} in not exist.`)
        continue
      }
      if (inFile.endsWith('.json')) {
        json2excel(inFile, options)
      }
      if (inFile.endsWith('.xlsx') || inFile.endsWith('.csv')) {
        excel2json(inFile, options)
      }
    }

    console.log('Done :)')
  }

}

main()

function json2excel(inFile, options) {

  const outFile = options.out || inFile.replace('.json', '.xlsx')
  const data = JSON.parse(fs.readFileSync(inFile))
  let result = []
  const workbook = XLSX.utils.book_new()

  if (Array.isArray(data)) {
    if (options.multi) {
      for (const item of data) {
        if (item.sheet_name && item.data) {
          const workSheet = XLSX.utils.json_to_sheet(item.data)
          XLSX.utils.book_append_sheet(workbook, workSheet, item.sheet_name)
        }
      }
    } else {
      result = data
    }
  } else {
    for (const key in data) {
      result.push({
        key,
        value: data[key]
      })
    }
  }

  if (result.length) {
    console.log(result.length)
    const workSheet = XLSX.utils.json_to_sheet(result)
    XLSX.utils.book_append_sheet(workbook, workSheet, 'sheet1')
  }

  XLSX.writeFile(workbook, outFile)

  console.log(`Convert ${inFile}`)
}

function excel2json(inFile, options) {

  const outFile = options.out || inFile.replace(/\.(xlsx|csv)/, '.json')

  const readOpts = {
    cellDates: true
  }
  if (options.rawStr) {
    readOpts.type = 'string'
    readOpts.raw = true
  }
  const workbook = XLSX.readFile(inFile, readOpts)

  const result = []
  for (const sheetName of workbook.SheetNames) {
    /* Get worksheet */
    const worksheet = workbook.Sheets[sheetName]
    const data = XLSX.utils.sheet_to_json(worksheet)

    console.log(data.length, sheetName)
    if (options.split) {
      const file = outFile.replace('.json', '-') + sheetName + '.json'
      fs.writeFileSync(file, JSON.stringify(data, null, 2), 'utf8')
      console.log('save ', sheetName)
    } else {
      result.push({
        sheet_name: sheetName,
        data,
      })
    }

  }
  if (!options.split) {
    fs.writeFileSync(outFile, JSON.stringify(result, null, 2), 'utf8')
  }
}