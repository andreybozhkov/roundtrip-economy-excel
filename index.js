const XLSX = require('xlsx');

let readOptions = {
    type: 'string',
    cellFormula: false,
    cellText: false,
    cellDates: true
};

let workbook = XLSX.readFile('./data/initial/import-projects-0.xls', readOptions);
let workbookSheetImportClean = workbook.Sheets['Import_clean'];
let sheetJSON = XLSX.utils.sheet_to_json(workbookSheetImportClean);
console.log(sheetJSON);