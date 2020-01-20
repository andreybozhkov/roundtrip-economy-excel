const XLSX = require('xlsx');

let readOptions = {
    type: 'string',
    cellFormula: false,
    cellText: false,
    cellDates: false
};

let workbookCustomerReferences = XLSX.readFile('./data/initial/Holmen Paper/2019-11 and 12/edge_protectors_2019_11_and_12.xlsx', readOptions);
let customerRefsJSON = XLSX.utils.sheet_to_json(workbookCustomerReferences.Sheets['edge_protectors']);
let workbookShipments = XLSX.readFile('./data/initial/Holmen Paper/2019-11 and 12/shipments.xls', readOptions);
let shipmentsJSON = XLSX.utils.sheet_to_json(workbookShipments.Sheets['shipments']);
let workbookProjects = XLSX.readFile('./data/initial/Holmen Paper/2019-11 and 12/projects.xls', readOptions);
let projectsJSON = XLSX.utils.sheet_to_json(workbookProjects.Sheets['projects']);
let reportJSON = [];

function addNewRoundTrip (importProject, exportProject) {
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

for (importProject of customerRefsJSON) {
    for (exportProject of shipmentsJSON) {
        if (exportProject.Trailer === importProject.Trailer && exportProject['Start Date'] > importProject['Start Date']) {
            addNewRoundTrip(importProject, exportProject);
            break;
        }
    }
}

let reportWorksheet = XLSX.utils.json_to_sheet(reportJSON, {
    cellDates: false
});

let newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, reportWorksheet, 'data');

XLSX.writeFile(newWorkbook, './data/Holmen Paper report.xlsx');