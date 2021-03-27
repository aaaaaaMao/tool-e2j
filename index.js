#!/usr/bin/env node

const fs = require('fs')
const XLSX = require('xlsx')
const {Command} = require('commander')

const program = new Command()
program
    .option('-o, --out <file>', 'output file')
    .option('--split', 'split by sheet name')

function run(argv) {
    
    program.parse()
    const options = program.opts()

    if (!argv || argv.length < 1) {
        console.log("basic usage: tool-e2j <.xlsx>")
    } else {
        const inFile = argv[0]
        if (!inFile.endsWith('.xlsx')) {
            throw new Error('Error input file: ' + inFile)
        }
        console.log(options)
        const outFile = options.out || inFile.replace('.xlsx', '.json')

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
        console.log('Done :)')
    }
}
 
run(process.argv.slice(2));