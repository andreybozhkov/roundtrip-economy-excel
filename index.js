const XLSX = require('xlsx');

let readOptions = {
    type: 'string',
    cellFormula: false,
    cellText: false,
    cellDates: true
};

let workbookImport = XLSX.readFile('./data/initial/import-projects-0.xls', readOptions);
let importJSON = XLSX.utils.sheet_to_json(workbookImport.Sheets['Import_clean']);
let workbookExport = XLSX.readFile('./data/initial/export-projects-0.xls', readOptions);
let exportJSON = XLSX.utils.sheet_to_json(workbookExport.Sheets['Export_clean']);
let reportJSON = [];

for (importProject of importJSON) {
    for (exportProject of exportJSON) {
        if (exportProject.Trailer === importProject.Trailer && exportProject['Start Date'] > importProject['Start Date']) {
            let newRoundTrip = {...importProject};
            newRoundTrip['Export ID'] = exportProject.ID;
            newRoundTrip['Export Trailer'] = exportProject.Trailer;
            newRoundTrip['Export Project Reporting Date'] = exportProject['Project Reporting Date'];
            newRoundTrip['Export Start Date'] = exportProject['Start Date'];
            newRoundTrip['Export End Date'] = exportProject['End Date'];
            newRoundTrip['Export Traffic Line Group'] = exportProject['Traffic Line Group'];
            newRoundTrip['Export Traffic Line'] = exportProject['Traffic Line'];
            newRoundTrip['Export Estimated Net Profit'] = exportProject['Estimated Net Profit'];
            newRoundTrip['Export Net Profit'] = exportProject['Net Profit'];
            newRoundTrip['Export Customer Companies from Shipments'] = exportProject['Customer Companies from Shipments'];
            reportJSON.push(newRoundTrip);
        }
    }
}

let reportWorksheet = XLSX.utils.json_to_sheet(reportJSON, {
    cellDates: true
});

let newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, reportWorksheet, 'report');

XLSX.writeFile(newWorkbook, 'report.xlsx');