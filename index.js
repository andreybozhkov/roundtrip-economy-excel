const XLSX = require('xlsx');

let readOptions = {
    type: 'string',
    cellFormula: false,
    cellText: false,
    cellDates: false
};

let workBookShipments = XLSX.readFile('./data/initial/NTG Multimodal costs/shipments_Italy_to_Sweden.xls', readOptions);
let shipmentsJSON = XLSX.utils.sheet_to_json(workBookShipments.Sheets['shipments']);
let workBookProjectsEstimates = XLSX.readFile('./data/initial/NTG Multimodal costs/estimates_projects.xls', readOptions);
let projectsEstimatesJSON = XLSX.utils.sheet_to_json(workBookProjectsEstimates.Sheets['projects_estimates']);
let reportJSON = [];
let projectsList = [];

function generateProjectsList (shipments) {
    for (shipment of shipments) {
        if (!projectsList.includes(shipment['Project ID'])) {
            projectsList.push(shipment['Project ID']);
        }
    }
}

generateProjectsList(shipmentsJSON);

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

for (importProject of shipmentsJSON) {
    for (exportProject of projectsEstimatesJSON) {
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

XLSX.writeFile(newWorkbook, 'report.xlsx');