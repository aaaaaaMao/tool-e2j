#!/usr/bin/env node

const fs = require('fs')
const XLSX = require('xlsx')
const { Command } = require('commander')

const program = new Command()
program
  .option('-o, --out <file>', 'output file')
  .option('--split', 'split by sheet name')
  .option('--j2e', 'json (key:value) to excel')
  .option('--multi', 'multi sheet')

function main() {
  
  program.parse()
  const options = program.opts()
    
  console.log(options)

  const _argv = program.args

  if (!_argv || _argv.length < 1) {
    console.log('basic usage: tool-e2j <.xlsx>')
  } else {

    if (options.j2e) {
      json2Excel(_argv, options)
    } else {
      const inFile = _argv[0]
      if (!inFile.endsWith('.xlsx') && !inFile.endsWith('.csv')) {
        throw new Error('Error input file: ' + inFile)
      }
      console.log(options)
      const outFile = options.out || inFile.replace(/\.(xlsx|csv)/, '.json')
    
      const workbook = XLSX.readFile(inFile)
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
    console.log('Done :)')
  }
}

main()

function json2Excel(argv, options) {
  const inFile = argv[0]
  if (!inFile.endsWith('.json')) {
    throw new Error('Error input file: ' + inFile)
  }

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

}