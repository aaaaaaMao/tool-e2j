#!/usr/bin/env node

const fs = require('fs');
const XLSX = require('xlsx');

function run(argv) {
    if (!argv || argv.length < 2) {
        console.log("Usage: tool-e2j <.xlsx> <.json>")
    } else {
        const workbook = XLSX.readFile(argv[0]);

        const first_sheet_name = workbook.SheetNames[0];
        /* Get worksheet */
        const worksheet = workbook.Sheets[first_sheet_name];

        const data = XLSX.utils.sheet_to_json(worksheet);

        fs.writeFileSync(argv[1], JSON.stringify(data, null, 2), 'utf8');
        console.log(data.length);
        console.log('Done :)');
    }
}
 
run(process.argv.slice(2));