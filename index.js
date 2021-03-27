#!/usr/bin/env node

const fs = require('fs')
const XLSX = require('xlsx')

function run(argv) {
    if (!argv || argv.length < 2) {
        console.log("basic usage: tool-e2j <.xlsx> <.json>")
    } else {

        const workbook = XLSX.readFile(argv[0])

        const result = []
        for (const sheetName of workbook.SheetNames) {
            /* Get worksheet */
            const worksheet = workbook.Sheets[sheetName]
            const data = XLSX.utils.sheet_to_json(worksheet)
            result.push({
                sheet_name: sheetName,
                data,
            })
            console.log(data.length, sheetName)
        }
        fs.writeFileSync(argv[1], JSON.stringify(result, null, 2), 'utf8')
    
        console.log('Done :)')
    }
}
 
run(process.argv.slice(2));